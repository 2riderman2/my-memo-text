
(async () => {
  "use strict";

  // 全拡張子の旧バージョン削除
  // 調査CSVで承認したバージョンのみ対象

  const DAYS = 100;

  // 初回は最大10件に制限する
  const MAX_DELETE_PER_RUN = 10;

  const GAP_MS = 400;

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

  const prefix =
    new URL(site).pathname.toLowerCase() + "/";

  const sleep = ms =>
    new Promise(r => setTimeout(r, ms));

  const headers = {
    Accept: "application/json;odata=verbose"
  };

  const literal = value =>
    encodeURIComponent(
      value.replace(/'/g, "''")
    );

  const fileApi = path =>
    `${site}/_api/web/` +
    `GetFileByServerRelativePath(` +
    `decodedurl='${literal(path)}')`;

  // ---------- 操作パネル ----------

  document.getElementById(
    "sp-all-version-delete"
  )?.remove();

  const panel = document.createElement("section");

  panel.id = "sp-all-version-delete";

  panel.style.cssText = [
    "position:fixed",
    "top:12px",
    "right:12px",
    "z-index:2147483647",
    "width:min(550px,92vw)",
    "max-height:88vh",
    "overflow:auto",
    "background:white",
    "color:#111",
    "border:2px solid #ad2525",
    "border-radius:8px",
    "padding:16px",
    "box-shadow:0 4px 16px #0005",
    "font:14px sans-serif"
  ].join(";");

  panel.innerHTML = `
    <strong>
      旧バージョン削除（全拡張子）
    </strong>
    <p>
      確認・承認済みの
      「全形式_旧バージョン調査.csv」を
      指定してください。
      <br>
      この版は最大 ${MAX_DELETE_PER_RUN} 件まで
      削除します。
    </p>
    <input data-file type="file" accept=".csv">
    <p data-status>待機中</p>
    <button data-start>候補を確認して削除</button>
    <button data-stop disabled>停止</button>
    <button data-save disabled>実行結果CSV保存</button>
    <p style="font-size:12px;color:#a00">
      旧版の削除は復元できると限りません。
      保存義務・保持ポリシー・承認を
      確認してください。
    </p>
  `;

  document.body.appendChild(panel);

  const q = x =>
    panel.querySelector(`[data-${x}]`);

  const ui = {
    file: q("file"),
    start: q("start"),
    stop: q("stop"),
    save: q("save"),
    status: q("status")
  };

  const status = s => {
    ui.status.textContent = s;
    console.log(s);
  };

  let stop = false;
  let running = false;

  const log = [];

  const logHead = [
    "結果",
    "ライブラリ名",
    "ファイル名",
    "ファイルパス",
    "バージョンID",
    "バージョン番号",
    "バージョン作成日時UTC",
    "詳細"
  ];

  // ---------- CSV読み込み ----------

  function parseCSV(s) {
    s = s.replace(/^\uFEFF/, "");

    const rows = [];
    let row = [];
    let v = "";
    let quoted = false;

    for (let i = 0; i < s.length; i++) {
      const c = s[i];

      if (quoted) {
        if (c === '"' && s[i + 1] === '"') {
          v += '"';
          i++;
        } else if (c === '"') {
          quoted = false;
        } else {
          v += c;
        }
      } else if (c === '"') {
        quoted = true;
      } else if (c === ",") {
        row.push(v);
        v = "";
      } else if (c === "\r" || c === "\n") {
        if (c === "\r" && s[i + 1] === "\n") {
          i++;
        }

        row.push(v);

        if (row.some(x => x !== "")) {
          rows.push(row);
        }

        row = [];
        v = "";
      } else {
        v += c;
      }
    }

    if (quoted) {
      throw Error(
        "CSV引用符が閉じられていません"
      );
    }

    if (v !== "" || row.length) {
      row.push(v);
      rows.push(row);
    }

    return rows;
  }

  function csvCell(v) {
    let s = String(v ?? "");

    if (/^\s*[=+\-@]/.test(s)) {
      s = "'" + s;
    }

    return '"' +
      s.replace(/"/g, '""') +
      '"';
  }

  // ---------- 実行結果CSV保存 ----------

  function saveLog() {
    const parts = [
      "\uFEFF",
      logHead.map(csvCell).join(",") + "\r\n"
    ];

    for (
      let i = 0;
      i < log.length;
      i += 3000
    ) {
      parts.push(
        log.slice(i, i + 3000)
          .map(r => r.map(csvCell).join(","))
          .join("\r\n") + "\r\n"
      );
    }

    const blob = new Blob(parts, {
      type: "text/csv;charset=utf-8"
    });

    const u = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = u;
    a.download =
      "SharePoint_旧バージョン削除結果.csv";

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => {
      URL.revokeObjectURL(u);
    }, 60000);
  }

  ui.save.onclick = saveLog;

  // ---------- 停止 ----------

  ui.stop.onclick = () => {
    stop = true;

    status(
      "停止を要求しました。" +
      "実行中の1件が終わってから停止します。"
    );
  };

  // ---------- REST API GET ----------

  // GETのみ通信エラー時に自動再試行する。
  // 削除POSTは重複防止のため自動再試行しない。

  async function getJSON(url) {
    for (let n = 0; n < 7; n++) {
      let res;

      try {
        res = await fetch(url, {
          credentials: "same-origin",
          headers
        });
      } catch (e) {
        if (n === 6) throw e;

        await sleep(
          Math.min(1500 * 2 ** n, 60000)
        );
        continue;
      }

      if (res.ok) {
        return res.json();
      }

      if (
        ![429, 500, 502, 503, 504].includes(
          res.status
        ) ||
        n === 6
      ) {
        const e = new Error(
          `HTTP ${res.status}: ` +
          (await res.text()).slice(0, 120)
        );

        e.status = res.status;
        throw e;
      }

      const ra = Number(
        res.headers.get("Retry-After")
      );

      await sleep(
        ra > 0
          ? ra * 1000
          : Math.min(1500 * 2 ** n, 60000)
      );
    }
  }

  // ---------- バージョン確認 ----------

  async function versionById(base, id) {
    try {
      const data = await getJSON(
        `${base}/Versions(${id})`
      );

      return data.d;

    } catch (e) {
      if (e.status === 404) {
        return null;
      }

      throw e;
    }
  }

  // 最新版のバージョン番号を取得する

  async function currentLabel(base) {
    const data = await getJSON(
      `${base}/ListItemAllFields`
    );

    const x = data.d;

    const label =
      x?.OData__UIVersionString ??
      x?._UIVersionString ??
      x?.UIVersionString;

    if (!label) {
      throw Error(
        "最新版の番号を確認できません。" +
        "安全のため中止"
      );
    }

    return String(label);
  }

  // ---------- Request Digest ----------

  let digest = "";
  let digestUntil = 0;

  async function requestDigest() {
    if (
      digest &&
      Date.now() < digestUntil
    ) {
      return digest;
    }

    const r = await fetch(
      `${site}/_api/contextinfo`,
      {
        method: "POST",
        credentials: "same-origin",
        headers: {
          Accept: headers.Accept
        }
      }
    );

    if (!r.ok) {
      throw Error(
        `contextinfo失敗 HTTP ${r.status}`
      );
    }

    const data = await r.json();

    const info =
      data?.d?.GetContextWebInformation;

    if (!info?.FormDigestValue) {
      throw Error(
        "RequestDigestを取得できません"
      );
    }

    digest = info.FormDigestValue;

    digestUntil =
      Date.now() +
      Math.max(
        60,
        Number(
          info.FormDigestTimeoutSeconds || 1800
        ) - 120
      ) * 1000;

    return digest;
  }

  // ---------- 実際の削除要求 ----------

  async function deleteVersion(base, id) {
    const token = await requestDigest();

    const res = await fetch(
      `${base}/Versions/DeleteByID(vid=${id})`,
      {
        method: "POST",
        credentials: "same-origin",
        headers: {
          ...headers,
          "X-RequestDigest": token,
          "Content-Type":
            "application/json;odata=verbose"
        },
        body: ""
      }
    );

    if (!res.ok) {
      throw Error(
        `削除要求 HTTP ${res.status}: ` +
        (await res.text()).slice(0, 120)
      );
    }
  }

  // ---------- 削除開始 ----------

  ui.start.onclick = async () => {
    if (running) return;

    if (!ui.file.files.length) {
      alert(
        "調査結果CSVを選んでください"
      );
      return;
    }

    running = true;
    stop = false;

    log.length = 0;

    ui.start.disabled = true;
    ui.stop.disabled = true;
    ui.save.disabled = true;

    try {
      const rows = parseCSV(
        await ui.file.files[0].text()
      );

      const header = rows.shift();

      if (!header) {
        throw Error("CSVが空です");
      }

      const ix = k =>
        header.indexOf(k);

      const required = [
        "状態",
        "ライブラリ名",
        "ファイル名",
        "ファイルパス",
        "バージョンID",
        "バージョン番号",
        "バージョン作成日時UTC",
        "基準日時UTC"
      ];

      if (
        required.some(x => ix(x) < 0)
      ) {
        throw Error(
          "指定した調査結果CSVの列が不足しています"
        );
      }

      const unique = new Map();

      // 「削除候補」だけを選択する

      for (const r of rows) {
        if (
          r[ix("状態")] !== "削除候補"
        ) {
          continue;
        }

        const path =
          r[ix("ファイルパス")];

        const name =
          r[ix("ファイル名")];

        const idstr =
          r[ix("バージョンID")];

        const id = Number(idstr);

        const label =
          r[ix("バージョン番号")];

        const created =
          r[ix("バージョン作成日時UTC")];

        const cutoff =
          r[ix("基準日時UTC")];

        if (
          typeof path !== "string" ||
          !path.startsWith("/") ||
          !path.toLowerCase().startsWith(
            prefix
          ) ||
          !name ||
          !/^\d+$/.test(idstr) ||
          !Number.isSafeInteger(id) ||
          id < 0 ||
          !label ||
          !Number.isFinite(
            Date.parse(created)
          ) ||
          !Number.isFinite(
            Date.parse(cutoff)
          )
        ) {
          throw Error(
            "削除候補CSVに不正な値があります: " +
            (name || path || "(不明)")
          );
        }

        const key =
          path + "\u0000" + id;

        const target = {
          library:
            r[ix("ライブラリ名")],
          name,
          path,
          id,
          label: String(label),
          created,
          cutoff
        };

        if (unique.has(key)) {
          const prev = unique.get(key);

          if (
            prev.label !== target.label ||
            prev.created !== target.created
          ) {
            throw Error(
              "同一IDで異なる候補が重複しています"
            );
          }

        } else {
          unique.set(key, target);
        }
      }

      const selected =
        [...unique.values()];

      if (!selected.length) {
        throw Error(
          "削除候補の行がありません"
        );
      }

      // 初期設定では最大10件だけ処理
      const targets = selected.slice(
        0,
        MAX_DELETE_PER_RUN
      );

      const libCount =
        new Set(
          targets.map(x => x.library)
        ).size;

      status(
        `CSV候補${selected.length}件／` +
        `今回の実行上限${targets.length}件。`
      );

      // ---------- 実行確認 ----------

      const ok = prompt(
        `全形式の旧バージョンを削除します。\n` +
        `CSV候補: ${selected.length}件\n` +
        `今回処理: ${targets.length}件\n` +
        `対象ライブラリ: ${libCount}件\n\n` +
        `最新版および100日未満のものは` +
        `削除しません。\n\n` +
        `承認済みの場合だけ\n` +
        `「削除実行 ${targets.length}」` +
        `と入力してください。`
      );

      if (
        ok !== `削除実行 ${targets.length}`
      ) {
        status(
          "キャンセルしました"
        );
        return;
      }

      ui.stop.disabled = false;

      let success = 0;
      let skip = 0;
      let fail = 0;

      // ---------- 削除処理 ----------

      for (
        let k = 0;
        k < targets.length;
        k++
      ) {
        if (stop) break;

        const t = targets[k];
        const base = fileApi(t.path);

        let result = "見送り";
        let detail = "";

        try {
          // 1. 最新版を確認する

          const latest =
            await currentLabel(base);

          // 2. 旧バージョンが今も存在するか確認

          const v = await versionById(
            base,
            t.id
          );

          if (!v) {
            detail =
              "指定バージョンは存在しません";

          } else if (
            v.IsCurrentVersion === true ||
            String(v.VersionLabel) === latest
          ) {
            detail =
              "最新版の可能性があるため除外";

          } else if (
            String(v.VersionLabel ?? "") !==
            t.label
          ) {
            detail =
              "調査時とバージョン番号が違います";

          } else if (
            Date.parse(v.Created || "") !==
            Date.parse(t.created)
          ) {
            detail =
              "調査時とバージョン作成日時が違います";

          } else if (
            !Number.isFinite(
              Date.parse(v.Created || "")
            ) ||
            Date.parse(v.Created) >
              Date.now() - DAYS * 86400000
          ) {
            detail =
              "100日以上前の条件に合いません";

          } else if (stop) {
            detail =
              "停止指示があったため見送り";

          } else {
            // 3. 条件に合う旧バージョンだけ削除

            await deleteVersion(
              base,
              t.id
            );

            result =
              "削除要求成功";

            detail =
              "REST APIが削除を受理" +
              "（実際の履歴は別途確認）";
          }

        } catch (e) {
          result =
            "削除失敗・要確認";

          detail = e.message;

          console.warn(
            t.path,
            t.id,
            e
          );
        }

        // ---------- ログ記録 ----------

        log.push([
          result,
          t.library,
          t.name,
          t.path,
          t.id,
          t.label,
          t.created,
          detail
        ]);

        if (
          result === "削除要求成功"
        ) {
          success++;

        } else if (
          result === "見送り"
        ) {
          skip++;

        } else {
          fail++;
        }

        status(
          `処理${k + 1}/${targets.length} ` +
          `成功${success} ` +
          `見送り${skip} ` +
          `失敗${fail}`
        );

        await sleep(GAP_MS);
      }

      status(
        `終了：削除要求成功${success}、` +
        `見送り${skip}、` +
        `失敗${fail}。` +
        `必ず実行結果CSVを保存してください。`
      );

      if (log.length) {
        ui.save.disabled = false;

        // 終了時に結果を自動保存する
        saveLog();
      }

    } catch (e) {
      console.error(e);

      status(
        "処理中断: " + e.message
      );

      if (log.length) {
        ui.save.disabled = false;
        saveLog();
      }

    } finally {
      running = false;
      ui.start.disabled = false;
      ui.stop.disabled = true;
    }
  };

  status(
    "準備完了。承認済み調査結果CSVを" +
    "選択してください。"
  );
})();
