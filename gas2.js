function saveEmailsToSingleSheet() {
    var sheetId = '';  // スプレッドシートのIDを指定
    var sheetName = 'mail'; // シート名を指定
  
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);
  
    var senderEmails = getSenderEmails()
    const today = new Date();
  
    let processedCount = 0;
    const maxCount = 20;
  
    for (var k = 0; k < senderEmails.length && processedCount < maxCount; k++) {
      var senderEmail = senderEmails[k];
      var threads = GmailApp.search('from:' + senderEmail + ' is:unread');
  
      for (var i = 0; i < threads.length && processedCount < maxCount; i++) {
        var messages = threads[i].getMessages();
        
        for (var j = 0; j < messages.length && processedCount < maxCount; j++) {
          var message = messages[j];
          var date = message.getDate();
          var subject = message.getSubject();
          var body = message.getPlainBody();
          var timestamp = new Date().toISOString();
          var messageId = message.getHeader("Message-ID");
  
          var sharedMailLink = messageId
            ? "https://mail.google.com/mail/#search/rfc822msgid:" + messageId.replace(/<|>/g, "")
            : "Message-ID not found";
  
          sheet.appendRow([senderEmail, subject, body, timestamp, sharedMailLink]);
  
          message.markRead(); // 既読にする
          processedCount++;
        }
      }
    }
  }
  
  
  
  function processSheetData() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mailSheet = spreadsheet.getSheetByName("mail");
    var mailData = mailSheet.getRange("A:F").getValues();
    var ankenSheet = spreadsheet.getSheetByName("anken2");
    var youinSheet = spreadsheet.getSheetByName("youin2");
  
    var apiKey = "";  // APIキーを入力
    var apiUrl = "https://api.openai.com/v1/chat/completions";
    var webhookUrl = "";
  
    var totalPromptTokens = 0;
    var totalCompletionTokens = 0;
    var processedCount = 0;
  
    for (var i = 1; i < mailData.length; i++) {
      var row = mailData[i];
      var promptText = row[2];  // C列
  
      if (row[5]) continue;  // F列が空白でないならスキップ
      if (!promptText) continue;  // C列が空白ならスキップ
  
      var response = callChatGPT(apiUrl, apiKey, promptText);
  
      if (response) {
        try {
          var parsedResponse = JSON.parse(response);
          totalPromptTokens += parsedResponse.usage.prompt_tokens;
          totalCompletionTokens += parsedResponse.usage.completion_tokens;
  
          var contentText = parsedResponse.choices[0].message.content.trim();
          var extractedArray = JSON.parse(contentText);
  
          if (!Array.isArray(extractedArray)) extractedArray = [extractedArray];
  
                    extractedArray.forEach(function(extractedData) {
            if (extractedData.type && extractedData.value) {
              var category = extractedData.type;
              var timestamp = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");

              if (category.includes("案件情報")) {
                var ankenRow = [
                  timestamp, // 受信日
                  extractedData["案件名"] || "",
                  extractedData["単価"] || "",
                  extractedData["開始時期"] || "",
                  extractedData["勤務地"] || "",
                  extractedData["必須スキル"] || "",
                  extractedData["尚可スキル"] || "",
                  extractedData["精算幅"] || "",
                  extractedData["面談回数"] || "",
                  extractedData["稼働時間"] || "",
                  extractedData["出社頻度"] || "",
                  extractedData["外国籍可否"] || "",
                  extractedData["年齢制限"] || "",
                  extractedData["備考"] || "",
                  extractedData["送信元"] || row[0], // 送信元
                  row[4] || "" // URL (メールのリンク)
                ];
                var emptyRow = findFirstEmptyRow(ankenSheet);
                ankenSheet.getRange(emptyRow, 1, 1, ankenRow.length).setValues([ankenRow]);
              } else if (category.includes("要員情報")) {
                var youinRow = [
                  timestamp, // 受信日
                  extractedData["氏名"] || "",
                  extractedData["年齢"] || "",
                  extractedData["性別"] || "",
                  extractedData["所属会社"] || "",
                  extractedData["属性"] || "",
                  extractedData["最寄駅"] || "",
                  extractedData["スキル"] || "",
                  extractedData["単価"] || "",
                  extractedData["稼働開始"] || "",
                  extractedData["稼働条件"] || "",
                  extractedData["並行状況"] || "",
                  extractedData["面談可能日"] || "",
                  extractedData["資格"] || "",
                  extractedData["希望条件"] || "",
                  extractedData["備考"] || "",
                  extractedData["送信元"] || row[0], // 送信元
                  row[4] || "" // URL (メールのリンク)
                ];
                var emptyRow = findFirstEmptyRow(youinSheet);
                youinSheet.getRange(emptyRow, 1, 1, youinRow.length).setValues([youinRow]);
              }

              processedCount++;
            }
          });
  
          mailSheet.getRange(i + 1, 6).setValue(1);
        } catch (e) {
          Logger.log("JSONパースエラー: " + e.toString());
        }
      }
    }
  
    var message = {
      text: `処理完了しました。\n処理した案件数: ${processedCount}\n消費トークン数: Prompt: ${totalPromptTokens}, Completion: ${totalCompletionTokens}`
    };
  
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(message)
    });
  }
  
  function callChatGPT(apiUrl, apiKey, info_value) {
    var prompt = `
    以下の内容をJSONの配列形式で返してください（複数ある場合は配列で返す）。
    各JSONオブジェクトには以下を含めてください。「type」の項目は、IT事業に従事する案件情報、または要員情報を判別し、その内容を記載するためのものです。

    【案件情報の場合】：
    - "type": "案件情報"
    - "value": 情報をまとめた文字列
    - "案件名": 案件の名称
    - "単価": 単価または単価範囲
    - "開始時期": 稼働開始時期
    - "勤務地": 勤務地・最寄駅
    - "必須スキル": 必須技術スキル
    - "尚可スキル": あると良いスキル（尚可）
    - "精算幅": 精算時間幅（例：140-180h）
    - "面談回数": 面談の回数
    - "稼働時間": 稼働時間（例：9:00-18:00）
    - "出社頻度": 出社頻度・リモート情報
    - "外国籍可否": 外国籍の可否
    - "年齢制限": 年齢上限・下限
    - "備考": その他特記事項
    - "送信元": メール送信者の会社名と送信者の氏名を記載。会社名がない場合は、送信者の氏名のみ、送信者が個人の場合は、会社名のみを記載。

    【要員情報の場合】：
    - "type": "要員情報"
    - "value": 情報をまとめた文字列
    - "氏名": イニシャルまたは氏名　イニシャルである場合には、他の文字列は入れないこと。
    - "年齢": 年齢 数字のみを半角で記載
    - "性別": 性別
    - "所属会社": 所属会社（弊社、1社先、2社先）
    - "属性": 雇用形態（フリー、プロパー）
    - "最寄駅": 最寄り駅
    - "スキル": 技術スキル一覧
    - "単価": 希望単価または単価範囲
    - "稼働開始": 稼働開始可能時期
    - "稼働条件": 稼働に関する条件（リモート、出社頻度等）
    - "並行状況": 他案件との並行状況
    - "面談可能日": 面談可能な日程
    - "資格": 保有資格
    - "希望条件": 本人の希望する案件条件
    - "備考": その他特記事項
    - "送信元": メール送信者

    ※所属会社：「弊社」「1社先」「2社先」のいずれか
    ※属性：「フリー」「プロパー」のいずれか
    ※稼働条件と希望条件は明確に区別してください
    ※該当しない項目は空文字で返してください
    ※コードブロックは使用しないでください
    `;
  
    var payload = {
      model: "gpt-4o-mini",
      messages: [{ role: "system", content: prompt },
      { role: "user", content: info_value  }],
      temperature: 0.2
    };
  
    var options = {
      method: "post",
      headers: {
        "Authorization": "Bearer " + apiKey,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload)
    };
  
    try {
      var response = UrlFetchApp.fetch(apiUrl, options);
      var jsonText = response.getContentText();
      Logger.log("GPT Response: " + jsonText);
      return jsonText;
    } catch (e) {
      Logger.log("Error: " + e.toString());
      return '{"choices": [{"message": {"content": "[{\"type\": \"エラー\", \"value\": \"APIエラー\", \"プロパー\": \"不明\", \"言語\": \"\", \"必須条件\": \"\", \"場所\": \"\", \"年齢\": \"\"}]"}}]}';
    }
  }
  /**
   * データ列が空白の最初の行を探す
   */
  function findFirstEmptyRow(sheet) {
    const sheetName = sheet.getName();
    let checkRange;
    
    if (sheetName === "anken2") {
      checkRange = "A:P"; // 案件情報は16列
    } else if (sheetName === "youin2") {
      checkRange = "A:R"; // 要員情報は18列
    } else {
      checkRange = "A:J"; // デフォルト
    }
    
    var data = sheet.getRange(checkRange).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i].every(cell => cell === "")) {
        return i + 1; // シートの行番号は1始まり
      }
    }
    return data.length + 1; // すべて埋まっている場合は次の行
  }
  
  
  
  
  // function processSpreadsheetData() {
  //   var apiKey = "";  // OpenAIのAPIキーを設定
  
  //   var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  //   // 「youin」シートのE列とF列を取得しJSON作成
  //   var youinSheet = ss.getSheetByName("youin");
  //   var youinData = youinSheet.getDataRange().getValues();
  //   var youinJSON = {};
  
  //   for (var i = 1; i < youinData.length; i++) { // ヘッダーをスキップ
  //     var key = youinData[i][4]; // E列 (0-based index: 4)
  //     var value = youinData[i][5]; // F列 (0-based index: 5)
  //     if (key) {
  //       youinJSON[key] = value;
  //     }
  //   }
  
  //   // 「anken」シートのB列の値をチェックし、該当行のC列を取得
  //   var ankenSheet = ss.getSheetByName("anken");
  //   var ankenData = ankenSheet.getDataRange().getValues();
  //   var target案件 = "【独占案件】エンド直増員案件・SUN小松(202503101)";
  //   var target案件情報 = null;
  
  //   for (var j = 1; j < ankenData.length; j++) { // ヘッダーをスキップ
  //     if (ankenData[j][1] === target案件) { // B列 (0-based index: 1)
  //       target案件情報 = ankenData[j][2]; // C列 (0-based index: 2)
  //       break;
  //     }
  //   }
  
  //   if (!target案件情報) {
  //     Logger.log("指定の案件が見つかりませんでした。");
  //     return;
  //   }
  
  //   // OpenAI APIに問い合わせ
  //   var prompt = "次のJSONデータの中から、以下の案件情報に最も適合する3件のキーを抽出してください。\n\n案件情報:\n" + target案件情報 + "\n\nJSONデータ:\n" + JSON.stringify(youinJSON);
  
  //   var payload = {
  //     model: "gpt-4o-mini",
  //     messages: [{ role: "user", content: prompt }],
  //     temperature: 0.2
  //   };
  
  //   var options = {
  //     method: "post",
  //     headers: {
  //       "Authorization": "Bearer " + apiKey,
  //       "Content-Type": "application/json"
  //     },
  //     payload: JSON.stringify(payload)
  //   };
  
  //   try {
  //     var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  //     var responseData = JSON.parse(response.getContentText());
  //     var result = responseData.choices[0].message.content;
  
  //     Logger.log("最適な3件のキー: " + result);
  //   } catch (e) {
  //     Logger.log("APIリクエストでエラーが発生しました: " + e.toString());
  //   }
  // }
  
  function onOpen() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // === 1. 「youin2」シートを A列（1列目）で降順ソート ===
    const youinSheetName = 'youin2';
    const youinSheet = ss.getSheetByName(youinSheetName);
    const youinStartRow = 2;
    const youinSortColumn = 1; // A列（受信日）
  
    if (youinSheet) {
      const youinLastRow = youinSheet.getLastRow();
      const youinLastCol = youinSheet.getLastColumn();
  
      if (youinLastRow >= youinStartRow) {
        youinSheet.getRange(youinStartRow, 1, youinLastRow - youinStartRow + 1, youinLastCol)
                  .sort({ column: youinSortColumn, ascending: false });
      }
    }
  
    // === 2. anken2シートの処理 ===
    const sheetName = 'anken2';
    const columnToSort = 1; // A列（受信日）
    const startRow = 2;
  
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
  
    ss.setActiveSheet(sheet);
  
    // A列で実際にデータがある行数を取得
    const lValues = sheet.getRange("A:A").getValues();
    const lastRow = lValues.filter(row => row[0] !== "").length;
  
    const lastCol = sheet.getLastColumn();
  
    if (lastRow >= startRow) {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol)
           .sort({ column: columnToSort, ascending: false });
    }
  
  }
  
  function callOpenAI(messages) {
    const apiKey = ""; // ★ここにOpenAIのAPIキーを記入
  
    const payload = {
      model: "gpt-4.1-mini",
      messages: messages
    };
  
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
  
    const url = "https://api.openai.com/v1/chat/completions";
  
    try {
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      return json.choices[0].message.content;
    } catch (e) {
      return "エラーが発生しました: " + e.toString();
    }
  }
  
  
  
  
  
  
  function onEdit(e) {
    const editedSheet = e.range.getSheet();
    const sheetName1 = "要件定義";
    const sheetName2 = "anken2";
  
    const row = e.range.getRow();
    const col = e.range.getColumn();
  
    // ▼ 要件定義：A列に入力 → F列に日付
    if (editedSheet.getName() === sheetName1) {
      if (col === 1 && e.range.getValue() !== "") {
        const dateCell = editedSheet.getRange(row, 6); // F列
        if (dateCell.getValue() === "") {
          dateCell.setValue(new Date());
        }
      }
      return;
    }
  
    // ▼ anken2：Q列チェック → 募集情報と要因をOpenAIでマッチング → R列に結果（整形せず）
    if (editedSheet.getName() === sheetName2 && col === 17) {
      const isChecked = e.value === "TRUE";
      if (!isChecked) {
        editedSheet.getRange(row, 18).clearContent(); // R列クリア
        return;
      }
  
      // B〜O列のヘッダーと値を取得してJSON化（案件名〜備考まで）
      const headers = editedSheet.getRange(1, 2, 1, 14).getValues()[0]; // B〜O列
      const values = editedSheet.getRange(row, 2, 1, 14).getValues()[0];
      const jobJson = {};
      for (let i = 0; i < headers.length; i++) {
        jobJson[headers[i]] = values[i];
      }
  
      // 要因情報を取得
      const youinData = extractRecentYouinData();
  
      // OpenAI プロンプト
      const prompt = `
  あなたは人材マッチングAIです。
  以下の「募集情報」に対して、最も適合する人材「要員情報」を3人まで選び、
  まず要員情報のJsonをその３名のみのもに絞ってください。（この時点は出力しないでください。）
  次に、マッチングの理由、および該当の人材情報の"メール検索タグ"を改行区切りで出力してください。マッチングの際は、単価の情報を必ず守ること。
  **要員情報は、一つのvalueが一人の要因情報です。複数の要因の情報を混同することが決して無いように注意して取り扱ってください。**
  **募集情報の単価が、要員情報の単価を1.1倍した価格より上である必要があります。**
  **「場所」情報が不明で無い場合には、募集情報の場所と要員情報の場所が離れすぎていないことをマッチング要件とします。通勤が２時間程度を超えないようにしてください。**
  **結果の出力にはマークダウンは使用しないこと。**
  **結果はGoogleスプレッドシートのセルに書き出しますが、その際リンクがクリックできる形式で出力してください。**
  
  【募集情報】
  ${JSON.stringify(jobJson)}
  
  【要員情報一覧】
  ${JSON.stringify(youinData)}
  
  
  `;
  
      // OpenAIへ問い合わせ（整形せずそのまま結果をR列に）
      const response = callOpenAI([{ role: "user", content: prompt }]);
      editedSheet.getRange(row, 18).setValue(response);
    }
  }
  
  
  
  function getSenderEmails() {
    const sheetName = "転送アドレス";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }
  
    // A列のデータを取得（1列で1000行までを目安に取得）
    const data = sheet.getRange("A1:A1000").getValues();
  
    // 空白を除いて1次元配列に整形
    const senderEmails = data
      .flat()
      .filter(email => email && email.toString().trim() !== "");
  
  
    Logger.log(senderEmails); // ログ出力で確認可能
    return senderEmails;
  }
  
  
  
  function extractRecentYouinData() {
    const sheetName = "youin2";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const dataRange = sheet.getRange("A1:R" + sheet.getLastRow()); // A〜R列（18列）
    const values = dataRange.getValues();
  
    const headers = values[0];  // 1行目をカラム名とする
    const today = new Date();
    const oneWeekAgo = new Date(today);
    oneWeekAgo.setDate(today.getDate() - 14);
  
    const results = [];
  
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const dateValue = row[0]; // A列（受信日）
  
      if (dateValue instanceof Date && dateValue >= oneWeekAgo && dateValue <= today) {
        const record = {};
        for (let j = 0; j < headers.length; j++) {
          record[headers[j]] = row[j];
        }
        results.push(record);
      }
    }
  
    Logger.log(JSON.stringify(results, null, 2));
    return results; // 必要に応じて他関数から呼び出せるように
  }
  
/**
 * マッチング機能 - 重複処理付き
 * plan.mdの122行目以降の仕様に基づいて実装
 */

/**
 * 1. youin2シートの重複レコードを削除する関数
 * 氏名、年齢、最寄駅、送信元の4つのカラムが完全一致するものは重複と判断し、
 * 受信日が最も新しいもの一つのみを残して削除する
 */
function removeDuplicateYouinRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const youinSheet = ss.getSheetByName("youin2");
  
  if (!youinSheet) {
    throw new Error("youin2シートが見つかりません");
  }
  
  const dataRange = youinSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // ヘッダー行を除いたデータを取得
  const dataRows = values.slice(1);
  
  // 重複判定のためのマップ作成
  const duplicateMap = new Map();
  
  // 各行をチェックして重複を判定
  dataRows.forEach((row, index) => {
    const 受信日 = row[0];
    const 氏名 = row[1];
    const 年齢 = row[2];
    const 最寄駅 = row[6]; // G列
    const 送信元 = row[16]; // Q列
    
    // 重複判定キーを作成（氏名、年齢、最寄駅、送信元）
    const duplicateKey = `${氏名}_${年齢}_${最寄駅}_${送信元}`;
    
    if (!duplicateMap.has(duplicateKey)) {
      duplicateMap.set(duplicateKey, []);
    }
    
    duplicateMap.get(duplicateKey).push({
      rowIndex: index + 2, // シートの実際の行番号（ヘッダー分+1、0ベース補正+1）
      受信日: 受信日,
      data: row
    });
  });
  
  // 削除対象の行を特定（受信日が古いもの）
  const rowsToDelete = [];
  
  duplicateMap.forEach((records, key) => {
    if (records.length > 1) {
      // 受信日で降順ソート（新しい順）
      records.sort((a, b) => new Date(b.受信日) - new Date(a.受信日));
      
      // 最新以外を削除対象に追加
      for (let i = 1; i < records.length; i++) {
        rowsToDelete.push(records[i].rowIndex);
      }
    }
  });
  
  // 行番号を降順でソートして削除（後ろから削除）
  rowsToDelete.sort((a, b) => b - a);
  
  // 実際に行を削除
  rowsToDelete.forEach(rowIndex => {
    youinSheet.deleteRow(rowIndex);
  });
  
  Logger.log(`重複削除完了: ${rowsToDelete.length}件の重複レコードを削除しました`);
  return rowsToDelete.length;
}

/**
 * 2. 重複処理後のデータから受信日が2週間以内のもののみを抽出してJSON形式で取得
 */
function getRecentYouinDataAsJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const youinSheet = ss.getSheetByName("youin2");
  
  if (!youinSheet) {
    throw new Error("youin2シートが見つかりません");
  }
  
  const dataRange = youinSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // 2週間前の日付を計算
  const today = new Date();
  const twoWeeksAgo = new Date(today);
  twoWeeksAgo.setDate(today.getDate() - 14);
  
  const youin_json = [];
  
  // ヘッダー行を除いたデータをチェック
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const 受信日 = row[0];
    
    // 受信日が2週間以内かチェック
    if (受信日 instanceof Date && 受信日 >= twoWeeksAgo && 受信日 <= today) {
      const record = {};
      
      // 各カラムの値をオブジェクトに格納
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      
      youin_json.push(record);
    }
  }
  
  Logger.log(`要員データ取得: 2週間以内の要員データ${youin_json.length}件を返します`);
  return youin_json;
}

/**
 * 3. anken2シートから前日の受信日の案件を抽出する
 */
function getYesterdayAnkenData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ankenSheet = ss.getSheetByName("anken2");
  
  if (!ankenSheet) {
    throw new Error("anken2シートが見つかりません");
  }
  
  const dataRange = ankenSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // 前日の日付を計算（時間は00:00:00〜23:59:59の範囲）
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStart = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 0, 0, 0);
  const yesterdayEnd = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 23, 59, 59);
  
  const yesterdayAnken = [];
  
  // ヘッダー行を除いたデータをチェック
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const 受信日 = row[0];
    
    // 受信日が前日かチェック
    if (受信日 instanceof Date && 受信日 >= yesterdayStart && 受信日 <= yesterdayEnd) {
      const record = {};
      
      // 各カラムの値をオブジェクトに格納
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      
      // 行番号も保存（後でシートに結果を書き込む際に使用）
      record._rowIndex = i + 1;
      
      yesterdayAnken.push(record);
    }
  }
  
  Logger.log(`前日の案件データ: ${yesterdayAnken.length}件取得しました`);
  return yesterdayAnken;
}

/**
 * 4. OpenAI APIを使用してマッチングを実行する
 */

  
// ===== シンプルなOpenAIマッチング & 通知 =====

/**
 * OpenAI APIでシンプルにマッチング判定（最大3名、氏名・単価・最寄駅・メール検索タグのみ返す）
 */
function simpleOpenAIMatching(anken, youinList) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error("OpenAI APIキーが未設定です");

  const prompt = `
あなたは人材マッチングAIです。
以下の「案件情報」には必ず「URL」フィールドが含まれています。
厳格な条件でマッチする要員情報を最大3名選び、
氏名・単価・最寄駅・メール検索タグ、スキルマッチ理由だけを含むJSON配列で返してください。
理由や説明文は不要です。該当者がいなければ空配列で返してください。

【マッチング条件（すべて満たすこと）】
1. 絶対条件の不一致（例：外国籍不可の場合は外国籍要員を除外、勤務地が一致しない場合も除外）
2. 必須スキルが一定以上一致していること（案件の必須スキルと要員のスキルが複数一致していること）
3. 稼働時期・稼働条件（リモート・出社頻度等）に整合性があること
4. 単価条件は必ず厳格に判定すること。案件の単価が、要員の単価の1.1倍より大きく、1.6倍を下回っていること（要員単価×1.1 < 案件単価 < 要員単価×1.6）を「両方絶対に」満たす場合のみマッチ対象とし、どちらか一方でも満たさない場合は「必ず除外」すること。相談可・応相談・近い・柔軟対応などの文言があっても、数値で厳密に判定し、例外や人間的な柔軟さは一切認めないこと。
- 単価の数値抽出時は「程度」「目安」「応相談」などの文字は無視し、必ず数値部分のみを使うこと。例えば「90万程度」は90万として判定すること。

【数値判定例】
- 要員単価66万の場合、案件単価65万は「要員単価×1.1=72.6 > 65」なので除外
- 要員単価100万の場合、案件単価65万は「要員単価×1.1=110 > 65」なので除外
- 要員単価60万の場合、案件単価65万は「60×1.1=66 < 65」なので除外
- 要員単価58万の場合、案件単価65万は「58×1.1=63.8 < 65」かつ「65<58×1.6=92.8」なのでマッチ
5. 勤務地条件をさらに厳格に判定すること。
   - 案件の勤務地と要員の最寄駅（または勤務地）が地理的に大きく離れている場合（都道府県や都市が異なる場合など）は必ず除外すること。
   - 案件が「完全リモート可」の場合のみ遠方でも許容すること。
   - 「リモート半々」「一部出社」「定例出社なし・用事時のみ出社」など、完全リモートでない場合は、勤務地が大きく離れている要員は必ず除外すること。
   - 「リモートあり」や「リモート半々」などの表現があっても、出社の可能性がある場合は遠方（例：東京と熊本など）は除外すること。

【出力仕様】
- 各要員のJSONオブジェクトには必ず「氏名」「単価」「最寄駅」「メール検索タグ」「スキルマッチ理由」の5つのフィールドを含めること。
- 「スキルマッチ理由」には、案件の必須スキルのうちどれがどのように一致したかを簡潔に記載すること。
- 要員側の「メール検索タグ」フィールドは必ず返すこと。値がない場合は空文字や省略ではなく、必ず"URLなし"という文字列を入れること。
- フィールド名は必ず「メール検索タグ」「スキルマッチ理由」とすること。
- 理由や説明文は一切不要。

【案件情報】
${JSON.stringify(anken)}

【要員情報一覧】
${JSON.stringify(youinList)}
`;

  const payload = {
    model: "gpt-4.1-mini",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.2,
    max_tokens: 1000
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    if (response.getResponseCode() === 429) {
      Logger.log("429エラー: レートリミットに達しました");
      return [];
    }
    const responseData = JSON.parse(response.getContentText());
    let content = responseData.choices[0].message.content.trim();
    content = content.replace(/```json|```/g, "").trim();
    const matchedList = JSON.parse(content);
    if (Array.isArray(matchedList)) {
      return matchedList;
    } else {
      return [];
    }
  } catch (e) {
    Logger.log("OpenAIマッチングAPIエラー: " + e.toString());
    return [];
  }
}

/**
 * Googleチャットにシンプル通知
 */
function simpleSendNotification(anken, matchedList) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('GOOGLE_CHAT_WEBHOOK_URL') || "";
  if (!webhookUrl) return;

  if (matchedList.length === 0) {
    // マッチしない場合は通知しない
    return;
  }

  let message = `【マッチ案件】${anken.案件名 || "案件名不明"}\n`;
  message += `単価: ${anken.単価 || "不明"} / 勤務地: ${anken.勤務地 || "不明"}\n`;
  message += `案件URL: ${anken.URL || anken.url || "URLなし"}\n`;
  message += `---\n`;
  matchedList.forEach(youin => {
    message += `🧑‍💻 ${youin.氏名 || "不明"} / 単価: ${youin.単価 || "不明"} / 最寄駅: ${youin.最寄駅 || "不明"}\n`;
    message += `📩 ${youin.メール検索タグ || "URLなし"}\n`;
    message += `スキルマッチ理由: ${youin.スキルマッチ理由 || "-"}\n`;
    message += `---\n`;
  });

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: message })
  });
}

/**
 * 一括実行：前日の案件×2週間以内の要員でマッチング＆通知
 */
function executeSimpleOpenAIMatchingProcess() {
  // まず重複削除を実行
  removeDuplicateYouinRecords();
  const youinList = getRecentYouinDataAsJson();
  const ankenList = getYesterdayAnkenData();

  ankenList.forEach(anken => {
    const matched = simpleOpenAIMatching(anken, youinList);
    simpleSendNotification(anken, matched);
    Utilities.sleep(20000); // 20秒待機
  });
}
  