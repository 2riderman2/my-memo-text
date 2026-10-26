
(async () => {
  "use strict";

  // 全拡張子対象・100日以上前の旧バージョン調査
  // GETのみ。削除処理は一切行わない。

  const DAYS = 100;
  const WORKERS = 2;
  const GAP_MS = 250;

  const match = location.pathname.match(
    /^\/(sites|teams)\/([^/]+)/i
  );

  if (
    !match ||
    !location.hostname.endsWith(".sharepoint.com")
  ) {
    alert("対象SharePointサイトで実行してください。");
    return;
  }

  const site =
    `${location.origin}/${match[1]}/${match[2]}`;

  const sitePrefix =
    new URL(site).pathname.toLowerCase() + "/";

  const sleep = ms =>
    new Promise(r => setTimeout(r, ms));

  const h = {
    Accept: "application/json;odata=verbose"
  };

  // ---------- 操作パネル ----------

  function panel() {
    document.getElementById(
      "sp-all-version-audit"
    )?.remove();

    const el = document.createElement("section");
    el.id = "sp-all-version-audit";

    el.style.cssText = [
      "position:fixed",
      "top:12px",
      "right:12px",
      "z-index:2147483647",
      "width:min(540px,92vw)",
      "max-height:88vh",
      "overflow:auto",
      "background:white",
      "color:#111",
      "border:2px solid #2463a9",
      "border-radius:8px",
      "padding:16px",
      "box-shadow:0 4px 16px #0005",
      "font:14px sans-serif"
    ].join(";");

    el.innerHTML = `
      <strong>
        全ファイル旧バージョン調査（削除なし）
      </strong>
      <p>
        全ファイル一覧CSVを指定してください。
        PDF以外も対象です。
      </p>
      <input data-file type="file" accept=".csv">
      <p data-status>待機中</p>
      <button data-start>調査開始／続きから再開</button>
      <button data-stop disabled>停止</button>
      <button data-csv disabled>調査CSV保存</button>
      <button data-reset>進捗削除</button>
      <p style="font-size:12px">
        同じCSV・同じブラウザでは調査済み結果を
        再利用します。新たな100日判定をしたい
        ときは進捗削除を行ってください。
      </p>
    `;

    document.body.appendChild(el);

    const q = x =>
      el.querySelector(`[data-${x}]`);

    return {
      file: q("file"),
      start: q("start"),
      stop: q("stop"),
      csv: q("csv"),
      reset: q("reset"),
      status: q("status")
    };
  }

  const ui = panel();

  const status = s => {
    ui.status.textContent = s;
    console.log(s);
  };

  let stop = false;
  let running = false;
  let activeDB = null;
  let currentDbName = "";

  // ---------- CSV読み込み ----------

  function parseCSV(s) {
    s = s.replace(/^\uFEFF/, "");

    const out = [];
    let row = [];
    let cell = "";
    let inQuotes = false;

    for (let i = 0; i < s.length; i++) {
      const c = s[i];

      if (inQuotes) {
        if (c === '"' && s[i + 1] === '"') {
          cell += '"';
          i++;
        } else if (c === '"') {
          inQuotes = false;
        } else {
          cell += c;
        }
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === ",") {
        row.push(cell);
        cell = "";
      } else if (c === "\r" || c === "\n") {
        if (c === "\r" && s[i + 1] === "\n") {
          i++;
        }

        row.push(cell);

        if (row.some(v => v !== "")) {
          out.push(row);
        }

        row = [];
        cell = "";
      } else {
        cell += c;
      }
    }

    if (inQuotes) {
      throw Error(
        "CSVの引用符が閉じられていません"
      );
    }

    if (cell !== "" || row.length) {
      row.push(cell);
      out.push(row);
    }

    return out;
  }

  // ---------- SharePoint REST API ----------

  async function getJSON(url) {
    for (let n = 0; n < 7; n++) {
      let r;

      try {
        r = await fetch(url, {
          credentials: "same-origin",
          headers: h
        });
      } catch (e) {
        if (n === 6) throw e;

        await sleep(
          Math.min(1500 * 2 ** n, 60000)
        );
        continue;
      }

      if (r.ok) {
        return r.json();
      }

      if (
        ![429, 500, 502, 503, 504].includes(r.status) ||
        n === 6
      ) {
        throw Error(
          `HTTP ${r.status}: ` +
          (await r.text()).slice(0, 120)
        );
      }

      const ra = Number(
        r.headers.get("Retry-After")
      );

      await sleep(
        ra > 0
          ? ra * 1000
          : Math.min(1500 * 2 ** n, 60000)
      );
    }
  }

  const literal = x =>
    encodeURIComponent(
      x.replace(/'/g, "''")
    );

  const base = p =>
    `${site}/_api/web/` +
    `GetFileByServerRelativePath(` +
    `decodedurl='${literal(p)}')`;

  function nextURL(s) {
    if (!s) return null;

    const u = new URL(s, site);

    if (
      u.origin !== location.origin ||
      !u.pathname.toLowerCase().startsWith(
        sitePrefix + "_api/"
      )
    ) {
      throw Error("不正な次ページURL");
    }

    return u.href;
  }

  async function versions(path) {
    let url =
      base(path) + "/Versions?$top=200";

    let pages = 0;
    const out = [];

    while (url) {
      if (++pages > 500) {
        throw Error(
          "バージョン履歴のページ数上限"
        );
      }

      const data = await getJSON(url);

      if (!Array.isArray(data?.d?.results)) {
        throw Error(
          "バージョン履歴の応答形式が不正"
        );
      }

      out.push(...data.d.results);

      url = nextURL(data.d.__next);
    }

    return out;
  }

  // ---------- IndexedDB進捗保存 ----------

  function dbOpen(name) {
    return new Promise((ok, ng) => {
      const r = indexedDB.open(name, 1);

      r.onupgradeneeded = () => {
        r.result.createObjectStore(
          "results",
          { keyPath: "path" }
        );
      };

      r.onsuccess = () =>
        ok(r.result);

      r.onerror = () =>
        ng(r.error);
    });
  }

  const dbGet = (db, path) =>
    new Promise((ok, ng) => {
      const r = db
        .transaction("results", "readonly")
        .objectStore("results")
        .get(path);

      r.onsuccess = () =>
        ok(r.result);

      r.onerror = () =>
        ng(r.error);
    });

  const dbPut = (db, item) =>
    new Promise((ok, ng) => {
      const tx = db.transaction(
        "results",
        "readwrite"
      );

      tx.objectStore("results").put(item);

      tx.oncomplete = ok;

      tx.onerror = () =>
        ng(tx.error);

      tx.onabort = () =>
        ng(tx.error);
    });

  // ---------- CSV出力 ----------

  function csvCell(v) {
    let s = String(v ?? "");

    if (/^\s*[=+\-@]/.test(s)) {
      s = "'" + s;
    }

    return '"' +
      s.replace(/"/g, '""') +
      '"';
  }

  function saveCSV(filename, head, rows) {
    const batches = [
      "\uFEFF",
      head.map(csvCell).join(",") + "\r\n"
    ];

    for (
      let i = 0;
      i < rows.length;
      i += 3000
    ) {
      batches.push(
        rows.slice(i, i + 3000)
          .map(r => r.map(csvCell).join(","))
          .join("\r\n") + "\r\n"
      );
    }

    const blob = new Blob(batches, {
      type: "text/csv;charset=utf-8"
    });

    const u = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = u;
    a.download = filename;

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => {
      URL.revokeObjectURL(u);
    }, 60000);
  }

  const head = [
    "状態",
    "ライブラリ名",
    "ファイル名",
    "ファイルパス",
    "バージョンID",
    "バージョン番号",
    "バージョン作成日時UTC",
    "バージョンサイズ_バイト",
    "基準日時UTC",
    "備考"
  ];

  // ---------- 入力CSV確認 ----------

  async function parseInput() {
    if (!ui.file.files.length) {
      throw Error(
        "全ファイル一覧CSVを選択してください"
      );
    }

    const text =
      await ui.file.files[0].text();

    const bytes =
      new TextEncoder().encode(text);

    const hash = Array.from(
      new Uint8Array(
        await crypto.subtle.digest(
          "SHA-256",
          bytes
        )
      ),
      x => x.toString(16).padStart(2, "0")
    ).join("");

    const records = parseCSV(text);
    const headers = records.shift();

    if (!headers) {
      throw Error("空のCSVです");
    }

    const p = headers.indexOf("ファイルパス");
    const n = headers.indexOf("ファイル名");
    const l = headers.indexOf("ライブラリ名");

    if ([p, n, l].some(x => x < 0)) {
      throw Error(
        "ライブラリ名・ファイル名・" +
        "ファイルパスの列が必要です"
      );
    }

    const files = new Map();
    const errors = [];

    for (const r of records) {
      const path = r[p] || "";
      const name = r[n] || "";
      const library = r[l] || "";

      if (!path && !name) continue;

      if (
        !path.startsWith("/") ||
        !path.toLowerCase().startsWith(
          sitePrefix
        ) ||
        !name
      ) {
        errors.push([
          "取得失敗",
          library,
          name,
          path,
          "",
          "",
          "",
          "",
          "",
          "サイト外・パス不正"
        ]);
      } else {
        files.set(path, {
          path,
          name,
          library
        });
      }
    }

    return {
      files: [...files.values()],
      errors,
      hash
    };
  }

  async function prepare() {
    const input = await parseInput();

    const name =
      "SP_ALL_VERSION_AUDIT_V3_" +
      site + "_" + input.hash;

    if (
      activeDB &&
      currentDbName !== name
    ) {
      activeDB.close();
      activeDB = null;
    }

    if (!activeDB) {
      activeDB = await dbOpen(name);
      currentDbName = name;
    }

    const meta = await dbGet(
      activeDB,
      "__META__"
    );

    if (!meta) {
      await dbPut(activeDB, {
        path: "__META__",
        cutoff: new Date(
          Date.now() - DAYS * 86400000
        ).toISOString()
      });
    }

    input.cutoff = (
      await dbGet(activeDB, "__META__")
    ).cutoff;

    return input;
  }

  // ---------- 停止ボタン ----------

  ui.stop.onclick = () => {
    stop = true;

    status(
      "停止要求を受け付けました。" +
      "処理中のファイルを終えて停止します。"
    );
  };

  // ---------- 調査開始 ----------

  ui.start.onclick = async () => {
    if (running) return;

    running = true;
    stop = false;

    ui.start.disabled = true;
    ui.reset.disabled = true;
    ui.stop.disabled = false;
    ui.csv.disabled = true;

    try {
      const input = await prepare();

      const cutoff =
        Date.parse(input.cutoff);

      const files = input.files;

      let index = 0;
      let done = 0;
      let fail = 0;
      let candidates = 0;

      status(
        `全形式 ${files.length.toLocaleString()}` +
        `ファイルを調査。基準日時 ${input.cutoff}`
      );

      async function worker() {
        while (!stop && index < files.length) {
          const f = files[index++];

          const cached = await dbGet(
            activeDB,
            f.path
          );

          if (cached?.complete) {
            done++;
            candidates += cached.rows.length;
            continue;
          }

          try {
            const vs = await versions(f.path);
            const rows = [];

            for (const v of vs) {
              if (
                v.IsCurrentVersion === true
              ) {
                continue;
              }

              const id = Number(
                v.ID ?? v.Id
              );

              const created = Date.parse(
                v.Created || ""
              );

              const label = String(
                v.VersionLabel ?? ""
              );

              if (
                !Number.isSafeInteger(id) ||
                id < 0 ||
                !Number.isFinite(created) ||
                !label
              ) {
                throw Error(
                  "バージョンID・日時・" +
                  "番号の取得に失敗"
                );
              }

              if (created <= cutoff) {
                rows.push([
                  "削除候補",
                  f.library,
                  f.name,
                  f.path,
                  id,
                  label,
                  v.Created ?? "",
                  v.Length ?? v.Size ?? "",
                  input.cutoff,
                  ""
                ]);
              }
            }

            await dbPut(activeDB, {
              path: f.path,
              complete: true,
              rows
            });

            candidates += rows.length;

          } catch (e) {
            fail++;

            await dbPut(activeDB, {
              path: f.path,
              complete: false,
              error: String(e.message)
            });
          }

          done++;

          if (
            done % 25 === 0 ||
            done === files.length
          ) {
            status(
              `進捗 ${done.toLocaleString()}/` +
              `${files.length.toLocaleString()}・` +
              `候補 ${candidates.toLocaleString()}・` +
              `今回失敗 ${fail}`
            );
          }

          await sleep(GAP_MS);
        }
      }

      await Promise.all(
        Array.from(
          {
            length: Math.min(
              WORKERS,
              files.length
            )
          },
          worker
        )
      );

      const rows = [...input.errors];

      let failed = input.errors.length;
      let found = 0;
      let bytes = 0;

      for (const f of files) {
        const x = await dbGet(
          activeDB,
          f.path
        );

        if (x?.complete) {
          rows.push(...x.rows);
          found += x.rows.length;

          for (const r of x.rows) {
            bytes += Number(r[7]) || 0;
          }

        } else {
          failed++;

          rows.push([
            x ? "取得失敗" : "未調査",
            f.library,
            f.name,
            f.path,
            "",
            "",
            "",
            "",
            input.cutoff,
            x?.error || "停止・未調査"
          ]);
        }
      }

      ui.csv.disabled = false;

      ui.csv.onclick = () =>
        saveCSV(
          "SharePoint_全形式_旧バージョン調査.csv",
          head,
          rows
        );

      status(
        `調査${stop ? "途中停止" : "終了"}。` +
        `候補${found.toLocaleString()}件・` +
        `推定${(bytes / 1073741824).toFixed(2)}GiB・` +
        `失敗／未調査${failed}件。` +
        `［調査CSV保存］で保存してください。`
      );

      if (failed) {
        console.warn(
          "未取得・未調査のファイルがあります。" +
          "CSVで確認してください。"
        );
      }

    } catch (e) {
      console.error(e);
      status("エラー: " + e.message);

    } finally {
      running = false;
      ui.start.disabled = false;
      ui.reset.disabled = false;
      ui.stop.disabled = true;
    }
  };

  // ---------- 進捗初期化 ----------

  ui.reset.onclick = async () => {
    if (running) return;

    if (!confirm(
      "選択中の全ファイル一覧CSVに対応する" +
      "調査済みデータを削除しますか？"
    )) {
      return;
    }

    try {
      const input = await parseInput();

      const name =
        "SP_ALL_VERSION_AUDIT_V3_" +
        site + "_" + input.hash;

      if (activeDB) {
        activeDB.close();
        activeDB = null;
        currentDbName = "";
      }

      const r =
        indexedDB.deleteDatabase(name);

      r.onsuccess = () =>
        status(
          "このCSVに対応する調査進捗を削除しました"
        );

      r.onerror = () =>
        status(
          "削除失敗: " + r.error
        );

      r.onblocked = () =>
        status(
          "別タブが調査データベースを使用しています"
        );

    } catch (e) {
      status("エラー: " + e.message);
    }
  };

  status(
    "準備完了。CSVを選んで調査開始してください。"
  );
})();
