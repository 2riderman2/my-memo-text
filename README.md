
(async () => {
  // SharePoint全ファイル一覧取得 v2
  // 11ライブラリ対応・権限不足スキップ・更新者取得
  // 読み取り専用 / IndexedDBで進捗保存

  const match = location.pathname.match(/^\/sites\/([^/]+)/i);
  if (!match) {
    console.error("/sites/〇〇 のSharePointサイトで実行してください");
    return;
  }

  const site = location.origin + "/sites/" + match[1];
  const databaseName = "SP_FILE_EXPORT_V2_" + site;
  const pageSize = 100;
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // -------- REST API --------

  async function getJson(url) {
    for (let attempt = 1; attempt <= 7; attempt++) {
      try {
        const res = await fetch(url, {
          method: "GET",
          credentials: "same-origin",
          headers: {
            Accept: "application/json;odata=verbose"
          }
        });

        if (res.ok) return await res.json();

        if (![429, 500, 502, 503, 504].includes(res.status)) {
          const err = new Error("HTTP " + res.status);
          err.status = res.status;
          throw err;
        }

        if (attempt === 7) {
          throw new Error("HTTP " + res.status + " 再試行上限");
        }

        const retry = Number(res.headers.get("Retry-After"));
        const wait = retry > 0
          ? retry * 1000
          : Math.min(2000 * 2 ** (attempt - 1), 60000);

        console.warn(
          `HTTP ${res.status}：${Math.ceil(wait / 1000)}秒後に再試行`
        );
        await sleep(wait);

      } catch (e) {
        if (e.status || attempt === 7 ||
            /再試行上限/.test(e.message)) throw e;

        await sleep(Math.min(2000 * 2 ** (attempt - 1), 60000));
      }
    }
  }

  function nextPage(value) {
    if (!value) return null;

    const u = new URL(value, site);
    if (u.hostname !== location.hostname) {
      throw new Error("ページ送り先のホストが異なります");
    }

    u.protocol = location.protocol;

    if (
      u.origin !== location.origin ||
      !u.pathname.startsWith(
        new URL(site).pathname + "/_api/"
      )
    ) {
      throw new Error("想定外のページ送りURLです");
    }

    return u.href;
  }

  // -------- IndexedDB --------

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(databaseName, 1);

      req.onupgradeneeded = () => {
        const db = req.result;
        db.createObjectStore("records", {keyPath: "key"});
        db.createObjectStore("states", {keyPath: "id"});
      };

      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function readState(db, id) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction("states", "readonly");
      const req = tx.objectStore("states").get(id);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  function savePage(db, records, state) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(
        ["records", "states"], "readwrite"
      );

      const store = tx.objectStore("records");

      for (const record of records) store.put(record);
      tx.objectStore("states").put(state);

      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  // -------- 日時・CSV --------

  function japanTime(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";

    return new Intl.DateTimeFormat("ja-JP", {
      timeZone: "Asia/Tokyo",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23"
    }).format(date);
  }

  function csvCell(value) {
    let s = String(value ?? "");

    // Excelの数式として解釈されることを防ぐ
    if (/^\s*[=+\-@]/.test(s)) s = "'" + s;

    return '"' + s.replace(/"/g, '""') + '"';
  }

  async function exportCsv(db, completedIds) {
    const header = [
      "ライブラリ名",
      "ファイルID",
      "ファイル名",
      "拡張子",
      "サイズ_バイト",
      "サイズ_MiB",
      "ファイルパス",
      "最終更新日時",
      "最終更新者"
    ];

    const chunks = [
      "\uFEFF",
      header.map(csvCell).join(",") + "\r\n"
    ];

    let lines = [];
    let count = 0;

    await new Promise((resolve, reject) => {
      const tx = db.transaction("records", "readonly");
      const req = tx.objectStore("records").openCursor();

      req.onsuccess = () => {
        const cursor = req.result;

        if (!cursor) {
          if (lines.length) {
            chunks.push(lines.join("\r\n") + "\r\n");
          }
          resolve();
          return;
        }

        const r = cursor.value;

        if (completedIds.has(r.libraryId)) {
          const dot = r.name.lastIndexOf(".");
          const ext = dot > 0
            ? r.name.slice(dot + 1).toLowerCase()
            : "";

          const mib = r.size == null
            ? ""
            : (Number(r.size) / 1048576).toFixed(2);

          const values = [
            r.library,
            r.itemId,
            r.name,
            ext,
            r.size ?? "",
            mib,
            r.path,
            japanTime(r.modified),
            r.editor
          ];

          lines.push(values.map(csvCell).join(","));
          count++;

          if (lines.length >= 5000) {
            chunks.push(lines.join("\r\n") + "\r\n");
            lines = [];
          }
        }

        cursor.continue();
      };

      req.onerror = () => reject(req.error);
      tx.onerror = () => reject(tx.error);
    });

    const blob = new Blob(chunks, {
      type: "text/csv;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download = "SharePoint_全ファイル一覧.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 60000);

    console.log("CSV出力開始:", count, "ファイル");
  }

  // -------- メイン処理 --------

  let db;

  try {
    console.log("SharePointライブラリ一覧を取得中...");

    let listUrl = site +
      "/_api/web/lists?" +
      new URLSearchParams({
        "$select": "Id,Title,BaseTemplate,ItemCount",
        "$top": "100"
      });

    const libraries = [];

    while (listUrl) {
      const data = await getJson(listUrl);
      libraries.push(
        ...data.d.results.filter(x => x.BaseTemplate === 101)
      );
      listUrl = nextPage(data.d.__next);
    }

    libraries.sort((a, b) =>
      a.Title.localeCompare(b.Title, "ja")
    );

    console.table(libraries.map((x, i) => ({
      番号: i + 1,
      ライブラリ名: x.Title,
      登録アイテム数: x.ItemCount
    })));

    const choiceText = libraries.map((x, i) =>
      `${i + 1}: ${x.Title}`
    ).join("\n");

    const answer = prompt(
      "取得するライブラリ番号を入力\n" +
      "全部なら ALL\n" +
      "一部なら 1,2,3 の形式\n\n" +
      choiceText
    );

    if (answer === null || !answer.trim()) {
      console.log("キャンセルしました");
      return;
    }

    let selected;

    if (answer.trim().toUpperCase() === "ALL") {
      selected = libraries;
    } else {
      const numbers = [...new Set(
        answer.trim().split(/[,，、\s]+/).map(Number)
      )];

      if (numbers.some(n =>
        !Number.isInteger(n) || n < 1 || n > libraries.length
      )) {
        throw new Error("選択番号が正しくありません");
      }

      selected = numbers.map(n => libraries[n - 1]);
    }

    db = await openDatabase();

    const statuses = [];
    const completedIds = new Set();

    for (const lib of selected) {
      console.log("処理開始:", lib.Title);

      const listApi = site +
        "/_api/web/lists(guid'" + lib.Id + "')";

      try {
        // ライブラリへのアクセスを確認
        await getJson(listApi + "?$select=Id");

        let state = await readState(db, lib.Id);

        if (state?.done) {
          console.log("保存済み:", lib.Title);
          completedIds.add(lib.Id);

          statuses.push({
            ライブラリ: lib.Title,
            状態: "取得済み",
            ファイル数: state.files,
            詳細: "保存データを利用"
          });
          continue;
        }

        const params = new URLSearchParams({
          "$select": [
            "Id",
            "FileLeafRef",
            "FileRef",
            "FSObjType",
            "Modified",
            "File/Length",
            "Editor/Title"
          ].join(","),
          "$expand": "File,Editor",
          "$orderby": "Id asc",
          "$top": String(pageSize)
        });

        let url = state?.nextUrl ||
          listApi + "/items?" + params.toString();

        let pages = state?.pages || 0;
        let seen = state?.seen || 0;
        let files = state?.files || 0;
        let missingSizes = state?.missingSizes || 0;

        while (url) {
          const data = await getJson(url);
          const items = data?.d?.results;

          if (!Array.isArray(items)) {
            throw new Error("想定外のAPIレスポンス");
          }

          const records = [];

          for (const item of items) {
            if (Number(item.FSObjType) !== 0) continue;

            const size = item.File?.Length ?? null;

            if (size === null) missingSizes++;

            records.push({
              key: lib.Id + "|" + item.Id,
              libraryId: lib.Id,
              library: lib.Title,
              itemId: item.Id,
              name: item.FileLeafRef || "",
              size,
              path: item.FileRef || "",
              modified: item.Modified || "",
              editor: item.Editor?.Title || ""
            });
          }

          const following = nextPage(data.d.__next);

          pages++;
          seen += items.length;
          files += records.length;

          state = {
            id: lib.Id,
            nextUrl: following,
            pages,
            seen,
            files,
            missingSizes,
            done: !following
          };

          await savePage(db, records, state);
          url = following;

          if (pages % 10 === 0 || !url) {
            console.log(
              `${lib.Title}: ${seen.toLocaleString()}件確認 / ` +
              `${files.toLocaleString()}ファイル`
            );
          }

          if (pages > 50000) {
            throw new Error("ページ数が安全上限を超えました");
          }

          await sleep(300);
        }

        completedIds.add(lib.Id);

        statuses.push({
          ライブラリ: lib.Title,
          状態: "成功",
          ファイル数: files,
          詳細: missingSizes
            ? `サイズ未取得 ${missingSizes}件`
            : "正常終了"
        });

      } catch (err) {
        // 権限不足でも、ほかのライブラリへ進む
        const denied = [401, 403, 404].includes(err.status);

        statuses.push({
          ライブラリ: lib.Title,
          状態: denied ? "スキップ" : "取得失敗",
          ファイル数: "未確定",
          詳細: denied
            ? `アクセス不可または未検出 HTTP ${err.status}`
            : err.message
        });

        console.warn(
          `${lib.Title}: ${denied ? "スキップ" : "取得失敗"}`,
          err.message
        );
      }
    }

    // -------- 結果の確認 --------

    console.log("===== 取得結果 =====");
    console.table(statuses);

    const failed = statuses.filter(x =>
      x.状態 !== "成功" && x.状態 !== "取得済み"
    );

    if (failed.length) {
      console.warn(
        `${failed.length}ライブラリはCSVに含まれません。`
      );
      console.table(failed);
    }

    if (completedIds.size > 0) {
      await exportCsv(db, completedIds);
    } else {
      console.warn("取得成功分がないためCSVは出力しません");
    }

    console.log("処理終了。CSVと取得結果を確認してください。");

    // CSV確認後の一時データ削除用
    window.SP_EXPORT_RESET = () => {
      db.close();
      const req = indexedDB.deleteDatabase(databaseName);

      req.onsuccess = () =>
        console.log("一時保存データを削除しました");

      req.onerror = () =>
        console.error("一時データ削除失敗", req.error);

      req.onblocked = () =>
        console.warn("ほかのタブがデータベースを使用中");
    };

  } catch (err) {
    console.error("全体処理を中断:", err);
    console.log(
      "保存済みデータは保持されます。" +
      "同じコードを再実行すると再開を試みます。"
    );
  }
})();
