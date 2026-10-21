
(async () => {
  "use strict";

  // PDF旧バージョン削除
  // ①の調査CSVにある「削除候補」のみ対象

  const match = location.pathname.match(
    /^\/(sites|teams)\/([^/]+)/i
  );

  if (!match || !location.hostname.endsWith(".sharepoint.com")) {
    alert("対象のSharePointサイトで実行してください");
    return;
  }

  const site =
    `${location.origin}/${match[1]}/${match[2]}`;

  const sitePath =
    new URL(site).pathname.toLowerCase() + "/";

  const DAYS = 100;

  const sleep = ms =>
    new Promise(r => setTimeout(r, ms));

  const headers = {
    Accept: "application/json;odata=verbose"
  };

  function escapeLiteral(value) {
    return encodeURIComponent(
      value.replace(/'/g, "''")
    ).replace(/'/g, "%27");
  }

  function fileApi(path) {
    return (
      site +
      "/_api/web/GetFileByServerRelativePath(decodedurl='" +
      escapeLiteral(path) +
      "')"
    );
  }

  // ---------- 操作パネル ----------

  document.getElementById("sp-pdf-version-tool")?.remove();

  const panel = document.createElement("section");
  panel.id = "sp-pdf-version-tool";

  panel.style.cssText = [
    "position:fixed",
    "z-index:2147483647",
    "top:16px",
    "right:16px",
    "width:min(560px,92vw)",
    "max-height:90vh",
    "overflow:auto",
    "background:#fff",
    "color:#111",
    "border:2px solid #b32929",
    "border-radius:12px",
    "padding:17px",
    "box-shadow:0 5px 28px #0006",
    "font:14px/1.5 sans-serif"
  ].join(";");

  panel.innerHTML = `
    <strong style="font-size:17px">
      PDF旧バージョン削除
    </strong>
    <p>
      確認済みの削除候補CSVを選択してください。
    </p>
    <input id="sp-delete-file"
           type="file"
           accept=".csv,text/csv"
           style="width:100%">
    <p id="sp-delete-status">実行待ち</p>
    <button id="sp-delete-run">
      対象を確認して削除開始
    </button>
    <button id="sp-delete-stop" disabled>
      次の削除前に停止
    </button>
    <button id="sp-delete-log" disabled>
      結果CSVを保存
    </button>
    <p style="font-size:12px">
      削除は元に戻せない場合があります。
      保持ポリシーや権限によって失敗する場合があります。
    </p>
  `;

  document.body.appendChild(panel);

  const ui = {
    file: panel.querySelector("#sp-delete-file"),
    run: panel.querySelector("#sp-delete-run"),
    stop: panel.querySelector("#sp-delete-stop"),
    log: panel.querySelector("#sp-delete-log"),
    status: panel.querySelector("#sp-delete-status")
  };

  function setStatus(message) {
    ui.status.textContent = message;
    console.log(message);
  }

  let shouldStop = false;
  const results = [];

  const resultHeader = [
    "結果",
    "ライブラリ名",
    "ファイル名",
    "ファイルパス",
    "バージョンID",
    "バージョン番号",
    "バージョン作成日時UTC",
    "詳細"
  ];

  // ---------- CSV ----------

  function parseCSV(text) {
    text = text.replace(/^\uFEFF/, "");

    const rows = [];
    let row = [];
    let field = "";
    let quoted = false;

    for (let i = 0; i < text.length; i++) {
      const c = text[i];

      if (quoted) {
        if (c === '"' && text[i + 1] === '"') {
          field += '"';
          i++;
        } else if (c === '"') {
          quoted = false;
        } else {
          field += c;
        }
      } else if (c === '"') {
        quoted = true;
      } else if (c === ",") {
        row.push(field);
        field = "";
      } else if (c === "\r" || c === "\n") {
        if (c === "\r" && text[i + 1] === "\n") i++;

        row.push(field);
        field = "";

        if (row.some(x => x !== "")) {
          rows.push(row);
        }

        row = [];
      } else {
        field += c;
      }
    }

    if (quoted) {
      throw new Error("CSVの引用符が閉じられていません");
    }

    if (field || row.length) {
      row.push(field);
      rows.push(row);
    }

    return rows;
  }

  function csvCell(value) {
    let s = String(value ?? "");

    if (/^\s*[=+@\-]/.test(s)) {
      s = "'" + s;
    }

    return '"' + s.replace(/"/g, '""') + '"';
  }

  function downloadLog() {
    const csv =
      "\uFEFF" +
      [resultHeader, ...results]
        .map(row => row.map(csvCell).join(","))
        .join("\r\n") +
      "\r\n";

    const blob = new Blob([csv], {
      type: "text/csv;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download = "SharePoint_PDF旧バージョン削除結果.csv";

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 60000);
  }

  ui.log.onclick = downloadLog;

  ui.stop.onclick = () => {
    shouldStop = true;
    ui.stop.disabled = true;

    setStatus(
      "停止要求を受け付けました。現在の通信後に停止します。"
    );
  };

  // ---------- REST GET ----------

  async function apiGet(url) {
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

      const message =
        (await response.text()).slice(0, 150);

      if (
        ![429, 500, 502, 503, 504].includes(response.status) ||
        attempt === 6
      ) {
        const error = new Error(
          `HTTP ${response.status}: ${message}`
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

  function nextPage(value) {
    if (!value) return null;

    const url = new URL(value, site);

    if (
      url.origin !== location.origin ||
      !url.pathname.toLowerCase().startsWith(
        sitePath + "_api/"
      )
    ) {
      throw new Error("ページ送り先が不正です");
    }

    return url.href;
  }

  async function getVersions(base) {
    const all = [];
    let url = base + "/Versions?$top=200";
    let pages = 0;

    while (url) {
      pages++;

      if (pages > 200) {
        throw new Error(
          "バージョン履歴のページ上限を超えました"
        );
      }

      const data = await apiGet(url);

      if (!Array.isArray(data?.d?.results)) {
        throw new Error(
          "バージョン履歴の形式が不正です"
        );
      }

      all.push(...data.d.results);
      url = nextPage(data.d.__next);
    }

    return all;
  }

  // ---------- 削除用ダイジェスト ----------

  let digest = "";
  let digestExpiry = 0;

  async function ensureDigest() {
    if (digest && Date.now() < digestExpiry) {
      return digest;
    }

    const response = await fetch(
      site + "/_api/contextinfo",
      {
        method: "POST",
        credentials: "same-origin",
        headers
      }
    );

    if (!response.ok) {
      throw new Error(
        "認証ダイジェスト取得失敗 HTTP " +
        response.status
      );
    }

    const data = await response.json();

    const info =
      data?.d?.GetContextWebInformation;

    if (!info?.FormDigestValue) {
      throw new Error(
        "認証ダイジェストの形式が不正です"
      );
    }

    digest = info.FormDigestValue;

    digestExpiry =
      Date.now() +
      Math.max(
        60,
        Number(info.FormDigestTimeoutSeconds || 1800) - 120
      ) * 1000;

    return digest;
  }

  // ---------- バージョン削除 ----------

  async function deleteVersion(base, id) {
    const url =
      base + `/Versions/deleteByID(${id})`;

    for (let attempt = 0; attempt < 5; attempt++) {
      const d = await ensureDigest();

      let response;

      try {
        response = await fetch(url, {
          method: "POST",
          credentials: "same-origin",
          headers: {
            Accept: "application/json;odata=verbose",
            "X-RequestDigest": d
          }
        });

      } catch (error) {
        throw new Error(
          "通信結果が不明です。再実行前に状態確認：" +
          error.message
        );
      }

      if (response.ok) {
        return;
      }

      if (
        response.status === 403 &&
        attempt === 0
      ) {
        digest = "";

      } else if (
        response.status !== 429 ||
        attempt === 4
      ) {
        throw new Error(
          `HTTP ${response.status}: ` +
          (await response.text()).slice(0, 150)
        );
      }

      const retry = Number(
        response.headers.get("Retry-After")
      );

      await sleep(
        Number.isFinite(retry) && retry > 0
          ? retry * 1000
          : Math.min(2000 * 2 ** attempt, 60000)
      );
    }
  }

  // ---------- 削除開始 ----------

  ui.run.onclick = async () => {
    if (!ui.file.files.length) {
      alert("調査結果CSVを選択してください");
      return;
    }

    ui.run.disabled = true;

    try {
      const text = await ui.file.files[0].text();
      const data = parseCSV(text);
      const header = data.shift();

      const ix = name => header.indexOf(name);

      const required = [
        "状態",
        "ライブラリ名",
        "ファイル名",
        "ファイルパス",
        "バージョンID",
        "バージョン番号",
        "バージョン作成日時UTC"
      ];

      if (required.some(x => ix(x) < 0)) {
        throw new Error(
          "調査用スクリプトのCSVではありません"
        );
      }

      const selected = new Map();

      for (const row of data) {
        if (row[ix("状態")] !== "削除候補") {
          continue;
        }

        const path = row[ix("ファイルパス")];
        const id = Number(row[ix("バージョンID")]);
        const name = row[ix("ファイル名")];

        const created =
          row[ix("バージョン作成日時UTC")];

        if (
          !path?.toLowerCase().startsWith(sitePath) ||
          !/\.pdf$/i.test(path) ||
          !/\.pdf$/i.test(name) ||
          !Number.isSafeInteger(id) ||
          id <= 0 ||
          !Number.isFinite(Date.parse(created))
        ) {
          throw new Error(
            "不正なファイルパスまたはバージョン指定です"
          );
        }

        selected.set(path + "|" + id, {
          library: row[ix("ライブラリ名")],
          name,
          path,
          id,
          label: row[ix("バージョン番号")],
          created
        });
      }

      const targets = [...selected.values()];

      if (!targets.length) {
        throw new Error("削除候補がありません");
      }

      const libraries = [
        ...new Set(targets.map(x => x.library))
      ];

      const confirmation = prompt(
        `削除対象：${targets.length}バージョン\n` +
        `対象ライブラリ：${libraries.length}件\n\n` +
        "最新版は残し、100日以上前の旧バージョンのみ削除します。\n" +
        "保存義務と削除権限を確認してください。\n\n" +
        "実行する場合のみ「削除実行」と入力してください。"
      );

      if (confirmation !== "削除実行") {
        setStatus("削除をキャンセルしました");
        return;
      }

      shouldStop = false;
      ui.stop.disabled = false;
      ui.log.disabled = false;

      let done = 0;
      let success = 0;
      let skipped = 0;
      let failed = 0;

      const grouped = new Map();

      for (const target of targets) {
        if (!grouped.has(target.path)) {
          grouped.set(target.path, []);
        }

        grouped.get(target.path).push(target);
      }

      const cutoff =
        Date.now() - DAYS * 86400000;

      for (const [path, items] of grouped) {
        if (shouldStop) break;

        const base = fileApi(path);

        try {
          // 最新版と旧バージョンを再確認

          const current = await apiGet(
            base + "?$select=UIVersionLabel"
          );

          const currentLabel = String(
            current?.d?.UIVersionLabel || ""
          );

          if (!currentLabel) {
            throw new Error(
              "最新版のバージョン番号を確認できません"
            );
          }

          const versions = await getVersions(base);

          for (const target of items) {
            if (shouldStop) break;

            let status = "見送り";
            let detail = "";

            try {
              const version = versions.find(
                x => Number(x.ID ?? x.Id) === target.id
              );

              if (!version) {
                detail =
                  "既に存在しない、または履歴に見つからない";

              } else if (
                version.IsCurrentVersion === true ||
                String(version.VersionLabel) === currentLabel
              ) {
                detail =
                  "最新版の可能性があるため除外";

              } else if (
                String(version.VersionLabel) !==
                String(target.label)
              ) {
                detail =
                  "バージョン番号が調査時と不一致";

              } else if (
                Date.parse(version.Created) !==
                Date.parse(target.created)
              ) {
                detail =
                  "作成日時が調査時と不一致";

              } else if (
                !Number.isFinite(
                  Date.parse(version.Created)
                ) ||
                Date.parse(version.Created) > cutoff
              ) {
                detail =
                  "100日以上前という条件に合わない";

              } else {
                await deleteVersion(
                  base,
                  target.id
                );

                status = "削除要求成功";
                detail =
                  "SharePointが削除要求を受理";
              }

            } catch (error) {
              status = "削除失敗";
              detail = error.message;
            }

            if (status === "削除要求成功") {
              success++;
            } else if (status === "削除失敗") {
              failed++;
            } else {
              skipped++;
            }

            results.push([
              status,
              target.library,
              target.name,
              target.path,
              target.id,
              target.label,
              target.created,
              detail
            ]);

            done++;

            if (
              done % 10 === 0 ||
              done === targets.length
            ) {
              setStatus(
                `処理 ${done}/${targets.length} / ` +
                `削除要求成功 ${success} / ` +
                `見送り ${skipped} / ` +
                `失敗 ${failed}`
              );
            }

            await sleep(250);
          }

        } catch (error) {
          for (const target of items) {
            if (shouldStop) break;

            results.push([
              "削除失敗",
              target.library,
              target.name,
              target.path,
              target.id,
              target.label,
              target.created,
              "取得・照合失敗：" + error.message
            ]);

            failed++;
            done++;
          }

          console.warn(
            "ファイル単位の処理失敗:",
            path,
            error.message
          );
        }
      }

      setStatus(
        `終了：${done}/${targets.length}件 / ` +
        `削除要求成功 ${success} / ` +
        `見送り ${skipped} / ` +
        `失敗 ${failed}` +
        (shouldStop ? "（途中停止）" : "")
      );

      downloadLog();

    } catch (error) {
      console.error(error);
      setStatus("処理停止：" + error.message);

    } finally {
      ui.run.disabled = false;
      ui.stop.disabled = true;

      if (results.length) {
        ui.log.disabled = false;
      }
    }
  };

  setStatus(
    "削除候補CSVを選択してください。確認なしには削除しません。"
  );
})();
