
(async () => {
  "use strict";

  // PDF旧バージョン調査専用
  // SharePointに対してGETのみ実行。削除しません。

  const siteMatch = location.pathname.match(
    /^\/(sites|teams)\/([^/]+)/i
  );

  if (!siteMatch || !location.hostname.endsWith(".sharepoint.com")) {
    alert("対象のSharePointサイトで実行してください。");
    return;
  }

  const site =
    `${location.origin}/${siteMatch[1]}/${siteMatch[2]}`;

  const sitePath = new URL(site).pathname.toLowerCase() + "/";
  const dbName = "SP_PDF_VERSION_AUDIT_V1_" + site;

  const DAYS = 100;
  const CONCURRENCY = 2;
  const PAUSE_MS = 150;

  const sleep = ms =>
    new Promise(resolve => setTimeout(resolve, ms));

  const headers = {
    Accept: "application/json;odata=verbose"
  };

  // ---------- 操作パネル ----------

  function createPanel() {
    document.getElementById("sp-pdf-version-tool")?.remove();

    const panel = document.createElement("section");
    panel.id = "sp-pdf-version-tool";

    panel.style.cssText = [
      "position:fixed",
      "z-index:2147483647",
      "top:16px",
      "right:16px",
      "width:min(550px,92vw)",
      "max-height:90vh",
      "overflow:auto",
      "background:#fff",
      "color:#111",
      "border:2px solid #2457a5",
      "border-radius:12px",
      "padding:17px",
      "box-shadow:0 5px 28px #0006",
      "font:14px/1.5 sans-serif"
    ].join(";");

    panel.innerHTML = `
      <strong style="font-size:17px">
        PDF旧バージョン調査（削除なし）
      </strong>
      <p>
        以前作成した「SharePoint_全ファイル一覧.csv」を
        選択してください。
      </p>
      <input id="sp-audit-file"
             type="file"
             accept=".csv,text/csv"
             style="width:100%">
      <p id="sp-audit-status">待機中</p>
      <button id="sp-audit-run">調査を開始</button>
      <button id="sp-audit-download" disabled>
        調査結果CSVを保存
      </button>
      <button id="sp-audit-reset">
        保存済み進捗を消去
      </button>
      <p style="font-size:12px">
        取得済みの進捗はブラウザ内に一時保存されます。
        このスクリプトはファイルを削除しません。
      </p>
    `;

    document.body.appendChild(panel);

    return {
      file: panel.querySelector("#sp-audit-file"),
      run: panel.querySelector("#sp-audit-run"),
      download: panel.querySelector("#sp-audit-download"),
      reset: panel.querySelector("#sp-audit-reset"),
      status: panel.querySelector("#sp-audit-status")
    };
  }

  const ui = createPanel();

  function setStatus(message) {
    ui.status.textContent = message;
    console.log(message);
  }

  // ---------- CSV読込 ----------

  function parseCSV(text) {
    text = text.replace(/^\uFEFF/, "");

    const rows = [];
    let row = [];
    let cell = "";
    let quoted = false;

    for (let i = 0; i < text.length; i++) {
      const c = text[i];

      if (quoted) {
        if (c === '"' && text[i + 1] === '"') {
          cell += '"';
          i++;
        } else if (c === '"') {
          quoted = false;
        } else {
          cell += c;
        }
      } else if (c === '"') {
        quoted = true;
      } else if (c === ",") {
        row.push(cell);
        cell = "";
      } else if (c === "\r" || c === "\n") {
        if (c === "\r" && text[i + 1] === "\n") i++;

        row.push(cell);
        cell = "";

        if (row.some(x => x !== "")) {
          rows.push(row);
        }

        row = [];
      } else {
        cell += c;
      }
    }

    if (quoted) {
      throw new Error("CSVの引用符が閉じられていません");
    }

    if (cell.length || row.length) {
      row.push(cell);
      rows.push(row);
    }

    return rows;
  }

  // ---------- REST API ----------

  async function getJson(url) {
    for (let attempt = 0; attempt < 7; attempt++) {
      let response;

      try {
        response = await fetch(url, {
          method: "GET",
          credentials: "same-origin",
          headers
        });
      } catch (error) {
        if (attempt === 6) throw error;

        await sleep(
          Math.min(1500 * 2 ** attempt, 60000)
        );
        continue;
      }

      if (response.ok) {
        return response.json();
      }

      const errorText =
        (await response.text()).slice(0, 180);

      if (
        ![429, 500, 502, 503, 504].includes(response.status) ||
        attempt === 6
      ) {
        const error = new Error(
          `HTTP ${response.status} ${errorText}`
        );
        error.status = response.status;
        throw error;
      }

      const retry = Number(
        response.headers.get("Retry-After")
      );

      await sleep(
        Number.isFinite(retry) && retry > 0
          ? retry * 1000
          : Math.min(1500 * 2 ** attempt, 60000)
      );
    }
  }

  function escapeLiteral(value) {
    return encodeURIComponent(
      value.replace(/'/g, "''")
    ).replace(/'/g, "%27");
  }

  function versionEndpoint(path) {
    return (
      site +
      "/_api/web/GetFileByServerRelativePath(decodedurl='" +
      escapeLiteral(path) +
      "')/Versions"
    );
  }

  function getNext(next) {
    if (!next) return null;

    const url = new URL(next, site);

    if (
      url.origin !== location.origin ||
      !url.pathname.toLowerCase().startsWith(
        sitePath + "_api/"
      )
    ) {
      throw new Error("不正なページ送り先です");
    }

    return url.href;
  }

  async function getVersions(path) {
    const versions = [];
    let url = versionEndpoint(path) + "?$top=200";
    let pageCount = 0;

    while (url) {
      pageCount++;

      if (pageCount > 200) {
        throw new Error(
          "バージョン履歴がページ上限を超えました"
        );
      }

      const data = await getJson(url);

      if (!Array.isArray(data?.d?.results)) {
        throw new Error(
          "バージョンAPIの応答形式が予想外です"
        );
      }

      versions.push(...data.d.results);
      url = getNext(data.d.__next);
    }

    return versions;
  }

  // ---------- 進捗保存 ----------

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(dbName, 1);

      request.onupgradeneeded = () => {
        const db = request.result;
        db.createObjectStore("files", {
          keyPath: "path"
        });
        db.createObjectStore("meta", {
          keyPath: "key"
        });
      };

      request.onsuccess = () =>
        resolve(request.result);

      request.onerror = () =>
        reject(request.error);
    });
  }

  function dbGet(db, store, key) {
    return new Promise((resolve, reject) => {
      const request = db
        .transaction(store, "readonly")
        .objectStore(store)
        .get(key);

      request.onsuccess = () =>
        resolve(request.result);

      request.onerror = () =>
        reject(request.error);
    });
  }

  function dbPut(db, store, value) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(store, "readwrite");
      tx.objectStore(store).put(value);

      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  // ---------- CSV出力 ----------

  function csvCell(value) {
    let s = String(value ?? "");

    if (/^\s*[=+@\-]/.test(s)) {
      s = "'" + s;
    }

    return '"' + s.replace(/"/g, '""') + '"';
  }

  function downloadCsv(filename, header, rows) {
    const content =
      "\uFEFF" +
      [header, ...rows]
        .map(r => r.map(csvCell).join(","))
        .join("\r\n") +
      "\r\n";

    const blob = new Blob([content], {
      type: "text/csv;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download = filename;

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 60000);
  }

  const outputHeader = [
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

  let db = null;
  let outputRows = [];

  // ---------- 調査開始 ----------

  ui.run.onclick = async () => {
    if (!ui.file.files.length) {
      alert("全ファイル一覧CSVを選択してください");
      return;
    }

    ui.run.disabled = true;
    ui.reset.disabled = true;
    ui.download.disabled = true;

    try {
      const text = await ui.file.files[0].text();
      const parsed = parseCSV(text);

      if (!parsed.length) {
        throw new Error("CSVが空です");
      }

      const header = parsed.shift();

      const pathIndex = header.indexOf("ファイルパス");
      const nameIndex = header.indexOf("ファイル名");
      const libIndex = header.indexOf("ライブラリ名");

      if (
        pathIndex < 0 ||
        nameIndex < 0 ||
        libIndex < 0
      ) {
        throw new Error(
          "CSVにファイルパス、ファイル名、ライブラリ名が必要です"
        );
      }

      const invalid = [];
      const selected = new Map();

      for (const row of parsed) {
        const name = row[nameIndex] || "";
        const path = row[pathIndex] || "";

        if (!/\.pdf$/i.test(name)) continue;

        if (
          !path.toLowerCase().startsWith(sitePath) ||
          !/\.pdf$/i.test(path)
        ) {
          invalid.push([
            "取得失敗",
            row[libIndex],
            name,
            path,
            "",
            "",
            "",
            "",
            "",
            "異なるサイトまたは不正なPDFパス"
          ]);
          continue;
        }

        selected.set(path, {
          path,
          name,
          library: row[libIndex] || ""
        });
      }

      const files = [...selected.values()];

      db = await openDb();

      let meta = await dbGet(db, "meta", "job");

      if (!meta) {
        meta = {
          key: "job",
          cutoff: new Date(
            Date.now() - DAYS * 86400000
          ).toISOString(),
          site
        };

        await dbPut(db, "meta", meta);
      }

      if (meta.site !== site) {
        throw new Error(
          "別サイトの保存済み進捗が存在します"
        );
      }

      const cutoffMs = Date.parse(meta.cutoff);

      setStatus(
        `PDF ${files.length.toLocaleString()}件を調査します。` +
        `基準日時：${meta.cutoff}`
      );

      let index = 0;
      let done = 0;
      let errors = 0;
      let found = 0;
      let cached = 0;

      async function worker() {
        while (index < files.length) {
          const file = files[index++];

          const existing = await dbGet(
            db,
            "files",
            file.path
          );

          if (existing?.done) {
            cached++;
            found += existing.candidates.length;
            done++;
            continue;
          }

          try {
            const versions = await getVersions(
              file.path
            );

            const candidates = [];

            for (const version of versions) {
              if (version.IsCurrentVersion === true) {
                continue;
              }

              const created = Date.parse(
                version.Created || ""
              );

              const id = Number(
                version.ID ?? version.Id
              );

              if (
                !Number.isSafeInteger(id) ||
                id <= 0 ||
                !Number.isFinite(created)
              ) {
                throw new Error(
                  "バージョンIDまたは作成日時が不正"
                );
              }

              if (created <= cutoffMs) {
                const size =
                  version.Length ??
                  version.Size ??
                  "";

                candidates.push([
                  "削除候補",
                  file.library,
                  file.name,
                  file.path,
                  id,
                  version.VersionLabel ?? "",
                  version.Created ?? "",
                  size,
                  meta.cutoff,
                  ""
                ]);
              }
            }

            await dbPut(db, "files", {
              path: file.path,
              done: true,
              candidates
            });

            found += candidates.length;

          } catch (error) {
            errors++;

            await dbPut(db, "files", {
              path: file.path,
              done: false,
              error: String(error.message)
            });

            console.warn(
              "取得失敗:",
              file.path,
              error.message
            );
          }

          done++;

          if (
            done % 25 === 0 ||
            done === files.length
          ) {
            setStatus(
              `進捗 ${done.toLocaleString()}/` +
              `${files.length.toLocaleString()} PDF / ` +
              `候補 ${found.toLocaleString()}件 / ` +
              `エラー ${errors}件 / ` +
              `保存済み ${cached}件`
            );
          }

          await sleep(PAUSE_MS);
        }
      }

      await Promise.all(
        Array.from(
          {
            length: Math.min(
              CONCURRENCY,
              files.length
            )
          },
          worker
        )
      );

      outputRows = [...invalid];
      let incomplete = invalid.length;

      for (const file of files) {
        const result = await dbGet(
          db,
          "files",
          file.path
        );

        if (result?.done) {
          outputRows.push(...result.candidates);
        } else {
          incomplete++;

          outputRows.push([
            "取得失敗",
            file.library,
            file.name,
            file.path,
            "",
            "",
            "",
            "",
            meta.cutoff,
            result?.error || "未取得"
          ]);
        }
      }

      const candidates = outputRows.filter(
        r => r[0] === "削除候補"
      );

      const bytes = candidates.reduce(
        (sum, r) => sum + (Number(r[7]) || 0),
        0
      );

      setStatus(
        `調査終了：PDF ${files.length.toLocaleString()}件 / ` +
        `削除候補 ${candidates.length.toLocaleString()}件 / ` +
        `推定対象容量 ` +
        `${(bytes / 1073741824).toFixed(2)} GiB / ` +
        `失敗 ${incomplete}件`
      );

      ui.download.disabled = false;

      ui.download.onclick = () => {
        downloadCsv(
          "SharePoint_PDF旧バージョン削除候補.csv",
          outputHeader,
          outputRows
        );
      };

      console.table({
        PDFファイル数: files.length,
        削除候補数: candidates.length,
        取得失敗数: incomplete
      });

    } catch (error) {
      console.error(error);
      setStatus("調査エラー：" + error.message);

    } finally {
      ui.run.disabled = false;
      ui.reset.disabled = false;
    }
  };

  // ---------- 調査データ初期化 ----------

  ui.reset.onclick = async () => {
    if (!confirm(
      "保存済み調査進捗を消去して、最初から調べ直しますか？"
    )) {
      return;
    }

    if (db) {
      db.close();
      db = null;
    }

    const request = indexedDB.deleteDatabase(
      dbName
    );

    request.onsuccess = () =>
      setStatus("保存済み調査進捗を消去しました");

    request.onerror = () =>
      setStatus("消去失敗：" + request.error);

    request.onblocked = () =>
      setStatus("別のタブがデータベースを使用中です");
  };

  setStatus(
    "準備完了。CSVを選択して調査を開始してください。"
  );
})();
