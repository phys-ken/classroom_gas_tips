function updateLinksAndPermissions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  
  // 実行開始時に全セルの背景色とノートをリセット
  dataRange.setBackground(null);
  dataRange.clearNote();
  
  var richTextValues = dataRange.getRichTextValues();
  var numRows = richTextValues.length;
  var numCols = richTextValues[0].length;
  
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      var cell = sheet.getRange(i + 1, j + 1);
      var richText = richTextValues[i][j];
      if (!richText) continue;
      
      var linkUrl = richText.getLinkUrl();
      if (!linkUrl) continue; // リンク情報がなければスキップ
      
      // Driveファイルリンクの場合
      if (linkUrl.indexOf("drive.google.com/file/d/") > -1) {
        try {
          var fileId = getDriveFileIdFromUrl(linkUrl);
          var file = DriveApp.getFileById(fileId);
          var fileName = file.getName();
          var mimeType = file.getMimeType();
          var icon = "📁"; // デフォルトアイコン
          var permissionChanged = false;
          
          if (mimeType === "application/pdf") {
            // PDFの場合、共有設定を変更
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            icon = "📄";
            permissionChanged = true;
          } else if (mimeType.indexOf("video/") === 0) {
            icon = "🎥";
          }
          
          var newDisplayText = icon + " " + fileName;
          var builder = SpreadsheetApp.newRichTextValue();
          builder.setText(newDisplayText);
          builder.setLinkUrl(0, newDisplayText.length, linkUrl);
          cell.setRichTextValue(builder.build());
          
          // PDFの共有設定変更が成功したセルは背景色を変更
          if (permissionChanged) {
            cell.setBackground("lightgreen");
          }
          
          Logger.log("セル (" + (i+1) + "," + (j+1) + ") 更新: " + newDisplayText);
        } catch (error) {
          Logger.log("セル (" + (i+1) + "," + (j+1) + ") Drive エラー: " + error.message);
          cell.setBackground("red");
          cell.setNote("Driveエラー: " + error.message);
        }
      }
      // YouTubeリンクの場合
      else if (linkUrl.indexOf("youtube.com/watch") > -1 || linkUrl.indexOf("youtu.be/") > -1) {
        try {
          var oembedUrl = "https://www.youtube.com/oembed?url=" + encodeURIComponent(linkUrl) + "&format=json";
          var response = UrlFetchApp.fetch(oembedUrl);
          var json = JSON.parse(response.getContentText());
          var videoTitle = json.title;
          var newDisplayText = "🎥 " + videoTitle;
          var builder = SpreadsheetApp.newRichTextValue();
          builder.setText(newDisplayText);
          builder.setLinkUrl(0, newDisplayText.length, linkUrl);
          cell.setRichTextValue(builder.build());
          
          Logger.log("セル (" + (i+1) + "," + (j+1) + ") 更新 (YouTube): " + newDisplayText);
        } catch (error) {
          Logger.log("セル (" + (i+1) + "," + (j+1) + ") YouTube エラー: " + error.message);
          cell.setBackground("red");
          cell.setNote("YouTubeエラー: " + error.message);
        }
      }
      // その他のリンクは変更しない
    }
  }
}

function getDriveFileIdFromUrl(url) {
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  } else {
    throw new Error("URLからファイルIDを抽出できませんでした: " + url);
  }
}
