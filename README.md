# Google Classroom Discovery Tool (GAS)

**Google Classroom のファイルをスプレッドシート上で効率的に整理・管理するツール**

このツールを使えば、Classroomで配布される授業ワークシートのショートカットから元のファイルを検出し、My Drive 内にあるファイルどこにでも自動で移動させることができます。
## ✨ 主な機能

*   **授業WSの自動検出**: Google ドライブ内の `Classroom` フォルダを自動検索し、授業WS一覧をリストアップします。
*   **ショートカットの完全解決**: ドライブ上のショートカット (`.shortcut`) を認識し、そのリンク先にある**実体ファイル**を操作対象として扱います。
*   **一括移動**: スプレッドシートに移動先フォルダの URL を貼り付けるだけで、指定したフォルダへファイルを一括移動します。
*   **ファイル名の同期**: ショートカット側でファイル名を変更していても安心です。移動時に、実体ファイルの名前をショートカットの名前に自動で同期（リネーム）します。
*   **安全設計**: 操作対象は**自分（実行ユーザー）がオーナーのファイル**に限定されます。共有ファイルや他人のファイルを誤って変更することはありません。

---

## 🛠 セットアップ

このツールは Google スプレッドシートの拡張機能（Apps Script）として動作します。

1.  **新規スプレッドシートの作成**: Google スプレッドシートを新規作成します（名前は任意です。例: `Classroom Manager`）。
2.  **スクリプトエディタを開く**: メニューの **拡張機能 > Apps Script** をクリックします。
3.  **コードの貼り付け**: エディタが開いたら、もともと入っているコードをすべて削除し、**以下のコードをコピーして貼り付けてください**。
4.  **保存**: ツールバーのディスクアイコン（💾）をクリックしてプロジェクトを保存します。

<details>
<summary><strong>👇 クリックしてソースコードを表示</strong></summary>

```javascript
// --- 設定 ---
var ROOT_FOLDER_NAME = "Classroom"; 
// ---------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 授業WSマネージャー')
    .addItem('1. 授業WSを探す', 'listClassFolders')
    .addItem('2. 選択した授業WSからファイルを取得', 'fetchFilesFromSelection')
    .addItem('3. ファイルを移動', 'processMoveQueue')
    .addToUi();
}

function listClassFolders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  sheet.appendRow(["選択", "授業WS名", "フォルダID", "フォルダURL"]);
  sheet.getRange("A1:D1").setFontWeight("bold");
  
  var folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  var rootFolder;
  
  if (folders.hasNext()) {
    rootFolder = folders.next();
  } else {
    SpreadsheetApp.getUi().alert("フォルダ名 '" + ROOT_FOLDER_NAME + "' が見つかりませんでした。");
    return;
  }

  try {
    var subfolders = rootFolder.getFolders();
    var rows = [];
    
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      rows.push(["☐", folder.getName(), folder.getId(), folder.getUrl()]);
    }
    
    if (rows.length > 0) {
      var range = sheet.getRange(2, 1, rows.length, 4);
      range.setValues(rows);
      sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
      SpreadsheetApp.getUi().alert(rows.length + " 件の授業WSが見つかりました。");
    } else {
      SpreadsheetApp.getUi().alert("'Classroom' フォルダは見つかりましたが、中身が空のようです。");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("フォルダへのアクセス中にエラーが発生しました: " + e.message);
  }
}

function fetchFilesFromSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var selectedFolderId = null;
  var selectedClassName = "";

  // ヘッダーをスキップし、チェックされたボックスを探す
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === true) { 
      selectedClassName = data[i][1];
      selectedFolderId = data[i][2];
      break; 
    }
  }

  if (!selectedFolderId) {
    SpreadsheetApp.getUi().alert("まず、授業WSフォルダの横にあるボックスにチェックを入れてください。");
    return;
  }

  // 全てクリアし、エラー防止のため入力規則も明示的に削除
  sheet.clear();
  sheet.getRange("A:E").clearDataValidations(); 
  
  sheet.appendRow(["ソース: " + selectedClassName, "ID: " + selectedFolderId]);
  sheet.appendRow(["ファイル名", "ファイルリンク", "移動先フォルダURL (ここに貼り付け)", "ステータス", "ファイルID"]);
  sheet.getRange("A2:E2").setFontWeight("bold");

  try {
    var folder = DriveApp.getFolderById(selectedFolderId);
    var files = folder.getFiles();
    var allFiles = [];

    // 診断用
    var shortcutCount = 0;
    var brokenShortcutCount = 0;

    // ソート用のフォーマット正規表現: Year_ で始まるもの (例: 2025_...)
    var formatRegex = /^\d{4}_/;

    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName(); // デフォルト: ファイル自身の名前を使用
      var id = file.getId();
      var url = file.getUrl();
      var isShortcut = false;
      var targetAccessError = false;
      
      // --- ショートカットの処理 ---
      if (file.getMimeType() === "application/vnd.google-apps.shortcut") {
        shortcutCount++;
        try {
          var targetId = file.getTargetId(); // ターゲットIDを取得
          
          // 実際のターゲットファイルオブジェクトの取得を試みる
          try {
            var targetFile = DriveApp.getFileById(targetId);
            url = targetFile.getUrl(); // アクセス可能ならターゲットURLを使用
          } catch (accessError) {
             // ターゲットオブジェクトにアクセスできない場合（権限の問題など）、
             // ショートカットのURLを保持しますが、IDはターゲットのものを使用します。
             console.log("ターゲットオブジェクトにアクセスできませんでした: " + name);
             targetAccessError = true;
          }

          // 重要な挙動: 
          // 1. ショートカット名（見た目）を保持
          // 2. ターゲットID（機能）を使用
          // 3. 可能ならターゲットURL、不可ならショートカットURLを使用
          name = file.getName(); 
          id = targetId; 
          isShortcut = true;

        } catch (e) {
          // getTargetId() さえ失敗する場合（非常に稀）、それは完全に壊れています。
          brokenShortcutCount++;
          console.log("完全に壊れたショートカット: " + file.getName());
          continue; 
        }
      }
      // -------------------------

      if (targetAccessError) {
        name = name + " [⚠️ アクセス制限あり]";
      }

      allFiles.push({
        name: name,
        url: url,
        id: id,
        isShortcut: isShortcut
      });
    }

    // --- ソートロジック ---
    var formattedFiles = [];
    var otherFiles = [];

    allFiles.forEach(function(item) {
      if (formatRegex.test(item.name)) {
        formattedFiles.push(item);
      } else {
        otherFiles.push(item);
      }
    });

    // フォーマット済みファイルの自然順ソート
    formattedFiles.sort(function(a, b) {
      return a.name.localeCompare(b.name, undefined, {numeric: true, sensitivity: 'base'});
    });

    // その他はアルファベット順（日本語）
    otherFiles.sort(function(a, b) {
      return a.name.localeCompare(b.name, 'ja');
    });

    var sortedFiles = formattedFiles.concat(otherFiles);
    var rows = [];

    sortedFiles.forEach(function(file) {
      rows.push([file.name, file.url, "", "待機中", file.id]);
    });

    if (rows.length > 0) {
      sheet.getRange(3, 1, rows.length, 5).setValues(rows);
    } else {
      var msg = "ファイルがリストされませんでした。";
      if (shortcutCount > 0) {
        msg += "\n\n診断情報:";
        msg += "\n- 見つかったショートカット数: " + shortcutCount;
        msg += "\n- 完全に壊れたショートカット数: " + brokenShortcutCount;
      }
      SpreadsheetApp.getUi().alert(msg);
      sheet.appendRow(["ファイルが見つかりません。", "", "", "", ""]);
    }
  } catch (e) {
     SpreadsheetApp.getUi().alert("ファイル取得中にエラーが発生しました: " + e.message);
  }
}

function processMoveQueue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; 

  var dataRange = sheet.getRange(3, 1, lastRow - 2, 5);
  var data = dataRange.getValues();
  var ui = SpreadsheetApp.getUi();
  
  var processedCount = 0;
  var myEmail = Session.getActiveUser().getEmail();

  for (var i = 0; i < data.length; i++) {
    var sheetName = data[i][0]; // シートに表示されている名前
    var targetUrl = data[i][2];
    var status = data[i][3];
    var fileId = data[i][4]; // ターゲットID
    
    // 既に完了またはスキップされた行は無視
    if (status === "完了" || status.indexOf("完了") === 0 || status === "スキップ") continue;

    // ターゲットURLが指定されている場合のみ処理
    if (targetUrl !== "") {
      var log = [];
      try {
        var file = DriveApp.getFileById(fileId);
        
        // --- 安全確認: 所有権 ---
        var owner = file.getOwner();
        if (!owner || owner.getEmail() !== myEmail) {
          data[i][3] = "スキップ (所有権なし)";
          continue;
        }

        // 1. 移動
        var targetFolderId = extractIdFromUrl(targetUrl);
        var targetFolder = DriveApp.getFolderById(targetFolderId);
        file.moveTo(targetFolder);
        log.push("移動済み");

        // 2. 名前同期 (実際のファイル名をシート上の名前に変更)
        // これにより、ショートカットに付けた「きれいな名前」が実際のファイルに適用されます。
        if (file.getName() !== sheetName) {
           file.setName(sheetName);
           log.push("名前に同期");
        }

        data[i][3] = "完了 (" + log.join(" & ") + ")";
        processedCount++;

      } catch (e) {
        data[i][3] = "エラー: " + e.message;
      }
    }
  }
  
  // ステータスを書き戻す
  dataRange.setValues(data);
  
  ui.alert(processedCount + ' 件のファイルを処理しました。');
}

function extractIdFromUrl(url) {
  if (!url) throw new Error("URLが空です");
  if (url.length < 50 && url.indexOf("http") === -1) return url;
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
```
</details>

---

## 🚀 使い方

スプレッドシートをリロードすると、メニューバーに **「🎓 授業WSマネージャー」** という項目が追加されます。
※ 初回実行時は、Google アカウントへのアクセス権限の承認（Authorization）画面が表示されます。画面の指示に従って許可してください。

### 1. 授業WSを探す
*   メニューから **「1. 授業WSを探す」** を実行します。
*   `Classroom` フォルダ内にあるサブフォルダ（授業WS）が自動的に検出され、シートに一覧表示されます。

### 2. ファイルを取得
*   A列のチェックボックスを使い、整理したい授業WSを**1つだけ選択**します。
*   メニューから **「2. 選択した授業WSからファイルを取得」** を実行します。
*   その授業WSフォルダ内のファイルがリストアップされます（ショートカットが含まれる場合、自動的に実体ファイルの情報が取得されます）。

### 3. ファイルを移動
*   移動させたいファイルの **C列（移動先フォルダURL）** に、移動先の Google ドライブフォルダの URL を貼り付けます。
*   メニューから **「3. ファイルを移動」** を実行します。
*   処理が完了するとステータスが「完了」に更新され、ファイルが移動します。この際、ショートカット側で名前を変更していた場合は、実体ファイルの名前もそれに合わせて自動的に変更されます。

## ⚠️ 注意点

*   **所有権の確認**: このスクリプトは、データの安全性を考慮し、**あなた自身がオーナー権限を持つファイル**のみを操作します。先生から配布された資料（閲覧権限のみ）などは「スキップ (所有権なし)」と表示され、移動や変更は行われません。
*   **フォルダ名**: デフォルト設定では、マイドライブ直下の「Classroom」という名前のフォルダを検索します。もしフォルダ名が異なる場合は、コード冒頭の `var ROOT_FOLDER_NAME = "Classroom";` の部分を実際のフォルダ名に書き換えてください。

---
*Created by Gemini CLI Agent*
