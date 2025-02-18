/**
 * スプレッドシートが開かれたときにメニューを追加
 */
function onOpen() {
  // メニューを追加
  SpreadsheetApp.getUi()
    .createMenu('Classroom操作')
    .addItem('初期設定を実行', 'initSetup')
    .addItem('クラス一覧を取得', 'getClassroomList')
    .addItem('投稿一覧を取得', 'getPostList')
    .addItem('投稿を再投稿', 'repostAssignments')
    .addToUi();
}

/**
 * 初期設定を行う関数
 */
function initSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 必要なシート名のリスト
  var sheetNames = ['ホーム', 'クラス一覧', '投稿一覧', '再投稿', '設定'];

  // 既存のシートを削除して新規作成
  sheetNames.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    ss.insertSheet(sheetName);
  });

  // 「ホーム」シートの設定
  var homeSheet = ss.getSheetByName('ホーム');
  homeSheet.getRange('A1').setValue('Google Classroom 操作用スプレッドシート');
  homeSheet.getRange('A2').setValue('以下にCourse IDを入力してください。');
  homeSheet.getRange('A4').setValue('Course ID:');
  homeSheet.getRange('B4').setValue('');
  homeSheet.getRange('A1:B1').setFontWeight('bold');

  // 「設定」シートの設定
  var settingsSheet = ss.getSheetByName('設定');
  // 投稿種別のリストを作成
  settingsSheet.getRange('A1').setValue('投稿種別');
  settingsSheet.getRange('A2').setValue('課題'); // Assignment
  settingsSheet.getRange('A3').setValue('質問'); // Question
  settingsSheet.getRange('A4').setValue('資料'); // Material

  // 「再投稿」シートの設定
  var repostSheet = ss.getSheetByName('再投稿');
  var headers = ['投稿先クラス', '投稿種別', '投稿名', '投稿 ID', '本文', 'トピック', '締切日', '添付ファイルID/URL'];
  repostSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 「再投稿」シートにデータバリデーションを設定
  // 投稿種別のプルダウン設定
  var postTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingsSheet.getRange('A2:A4'), true)
    .build();
  repostSheet.getRange('B2:B').setDataValidation(postTypeRule);

  // クラス一覧を取得して、クラス名のプルダウンを設定
  var optionalArgs = {
    'teacherId': 'me'
  };
  var courses = Classroom.Courses.list(optionalArgs).courses;
  if (courses && courses.length > 0) {
    var classNames = courses.map(function(course) {
      return [course.name];
    });
    settingsSheet.getRange(1, 2).setValue('クラス名');
    settingsSheet.getRange(2, 2, classNames.length, 1).setValues(classNames);

    var classNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(settingsSheet.getRange(2, 2, classNames.length, 1), true)
      .build();
    repostSheet.getRange('A2:A').setDataValidation(classNameRule);
  }
}

/**
 * 教師として参加しているクラス一覧を取得し、「クラス一覧」シートに表示
 */
function getClassroomList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'クラス一覧';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Classroom APIを使用してクラス一覧を取得
  var optionalArgs = {
    'teacherId': 'me'
  };
  var courses = Classroom.Courses.list(optionalArgs).courses;

  if (courses && courses.length > 0) {
    // シートにヘッダーを追加
    sheet.appendRow(['クラス名', 'Course ID', '作成日', '更新日']);

    // 各クラスの情報をシートに書き込む
    courses.forEach(function(course) {
      sheet.appendRow([
        course.name,
        course.id,
        course.creationTime,
        course.updateTime
      ]);
    });
  } else {
    SpreadsheetApp.getUi().alert('参加しているクラスが見つかりませんでした。');
  }
}

/**
 * 指定したクラスの投稿一覧を取得し、「投稿一覧」シートに表示
 */
function getPostList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var homeSheet = ss.getSheetByName('ホーム');
  var courseId = homeSheet.getRange('B4').getValue();

  if (!courseId) {
    SpreadsheetApp.getUi().alert('ホームシートのCourse IDを入力してください。');
    return;
  }

  var sheetName = '投稿一覧';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // シートにヘッダーを追加
  var headers = ['投稿種別', '投稿名', '投稿 ID', '本文', 'トピック', '締切日', '投稿日'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var row = 2; // データの開始行

  // Classroom APIを使用して投稿一覧を取得
  var courseWork = Classroom.Courses.CourseWork.list(courseId).courseWork || [];
  var announcements = Classroom.Courses.Announcements.list(courseId).announcements || [];
  var materials = Classroom.Courses.CourseWorkMaterials.list(courseId).courseWorkMaterial || [];

  // 全ての投稿を一つの配列にまとめる
  var posts = [];

  // 課題と質問をpostsに追加
  courseWork.forEach(function(work) {
    var type = work.workType === 'ASSIGNMENT' ? '課題' : '質問';
    posts.push({
      type: type,
      title: work.title,
      id: work.id,
      description: work.description || '',
      topicId: work.topicId || '',
      dueDate: work.dueDate ? `${work.dueDate.year}/${work.dueDate.month}/${work.dueDate.day}` : '',
      creationTime: work.creationTime,
      materials: work.materials || []
    });
  });

  // お知らせをpostsに追加
  announcements.forEach(function(announcement) {
    posts.push({
      type: 'お知らせ',
      title: announcement.text || '',
      id: announcement.id,
      description: announcement.text || '',
      topicId: announcement.topicId || '',
      dueDate: '',
      creationTime: announcement.creationTime,
      materials: announcement.materials || []
    });
  });

  // 資料をpostsに追加
  materials.forEach(function(material) {
    posts.push({
      type: '資料',
      title: material.title || '',
      id: material.id,
      description: material.description || '',
      topicId: material.topicId || '',
      dueDate: '',
      creationTime: material.creationTime,
      materials: material.materials || []
    });
  });

  // 投稿を投稿日でソート
  posts.sort(function(a, b) {
    return new Date(a.creationTime) - new Date(b.creationTime);
  });

  // トピック一覧を取得し、「設定」シートに記載
  var topics = Classroom.Courses.Topics.list(courseId).topic || [];
  var settingsSheet = ss.getSheetByName('設定');
  settingsSheet.getRange('D1').setValue('トピック名');
  var topicNames = [];
  if (topics.length > 0) {
    topicNames = topics.map(function(topic) {
      return [topic.name];
    });
    settingsSheet.getRange(2, 4, topicNames.length, 1).setValues(topicNames);

    // 「再投稿」シートのトピック列にデータバリデーションを設定
    var repostSheet = ss.getSheetByName('再投稿');
    var topicRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(settingsSheet.getRange(2, 4, topicNames.length, 1), true)
      .build();
    repostSheet.getRange('F2:F').setDataValidation(topicRule);
  }

  // 各投稿をシートに書き込む
  posts.forEach(function(post) {
    var rowData = [
      post.type,
      post.title,
      post.id,
      post.description,
      post.topicId,
      post.dueDate,
      post.creationTime
    ];

    // 添付ファイルを列方向に展開
    var attachments = [];
    if (post.materials) {
      post.materials.forEach(function(material) {
        if (material.driveFile) {
          attachments.push(material.driveFile.driveFile.alternateLink);
        } else if (material.link) {
          attachments.push(material.link.url);
        } else if (material.youtubeVideo) {
          attachments.push(material.youtubeVideo.alternateLink);
        }
        // 他の素材タイプがあればここに追加
      });
    }

    // シートにデータを書き込む
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // 添付ファイルを追加
    if (attachments.length > 0) {
      sheet.getRange(row, headers.length + 1, 1, attachments.length).setValues([attachments]);
    }

    row++;
  });

  // 行の高さを調整（一括で1行の高さに設定）
  sheet.setRowHeights(2, sheet.getLastRow(), 21); // デフォルトの行の高さに設定
}

/**
 * スプレッドシート上のデータを元に投稿を再投稿する関数
 */
function repostAssignments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var repostSheet = ss.getSheetByName('再投稿');
  var dataRange = repostSheet.getDataRange();
  var data = dataRange.getValues();
  var headers = data.shift(); // ヘッダー行を取得

  data.forEach(function(row) {
    var postData = {};
    var postId = '';
    var courseId = '';
    var postType = '';
    var updateExisting = false;
    var skipPost = false;

    headers.forEach(function(header, index) {
      switch (header) {
        case '投稿先クラス':
          var classInput = row[index];
          courseId = getCourseIdByNameOrId(classInput);
          break;
        case '投稿種別':
          postType = row[index];
          break;
        case '投稿名':
          postData.title = row[index];
          break;
        case '投稿 ID':
          postId = row[index];
          break;
        case '本文':
          postData.description = row[index];
          break;
        case 'トピック':
          var topicInput = row[index];
          if (topicInput) {
            postData.topicId = getTopicIdByNameOrId(courseId, topicInput);
          }
          break;
        case '締切日':
          if (row[index]) {
            var dateParts = row[index].split('/');
            postData.dueDate = {
              year: parseInt(dateParts[0]),
              month: parseInt(dateParts[1]),
              day: parseInt(dateParts[2])
            };
          }
          break;
        default:
          break;
      }
    });

    // 添付ファイルの取得と処理
    var attachmentStartIndex = headers.indexOf('添付ファイルID/URL');
    var attachmentIds = row.slice(attachmentStartIndex).filter(function(cell) { return cell; });
    var validAttachments = [];
    var invalidAttachmentUrls = []; // 添付できなかったファイルのURLを保持

    if (attachmentIds.length > 0) {
      attachmentIds.forEach(function(idOrUrl) {
        var attachment = {};
        if (isValidUrl(idOrUrl)) {
          if (isYouTubeUrl(idOrUrl)) {
            // YouTubeのURLの場合
            attachment = {
              youtubeVideo: {
                id: extractYouTubeId(idOrUrl)
              }
            };
          } else {
            // 通常のリンクの場合
            attachment = {
              link: {
                url: idOrUrl
              }
            };
          }
        } else {
          // ファイルIDとして処理
          var fileId = extractFileId(idOrUrl);
          attachment = {
            driveFile: {
              driveFile: {
                id: fileId
              }
            }
          };
        }

        // 個別に添付ファイルを試してみる
        try {
          // 仮の投稿データを作成
          var tempPostData = {
            materials: [attachment],
            title: postData.title || 'Temporary Title'
          };

          // 仮の投稿を試す（非公開のクラスで投稿し、エラーが出るか確認）
          var tempMaterial = Classroom.Courses.CourseWorkMaterials.create(tempPostData, courseId);

          // エラーが出なければ有効な添付ファイルとして追加
          validAttachments.push(attachment);

          // 仮の投稿を削除
          Classroom.Courses.CourseWorkMaterials.remove(courseId, tempMaterial.id);

        } catch (e) {
          if (e.toString().indexOf('AttachmentNotVisible') !== -1) {
            // 添付できなかったファイルのURLを保持
            invalidAttachmentUrls.push(idOrUrl);
          } else {
            // その他のエラーは再スロー
            throw e;
          }
        }
      });

      // 有効な添付ファイルがあれば追加
      if (validAttachments.length > 0) {
        postData.materials = validAttachments;
      }

      // 添付できなかったファイルのURLを本文に追記
      if (invalidAttachmentUrls.length > 0) {
        var urlsText = '\n添付できなかったファイル:\n' + invalidAttachmentUrls.join('\n');
        postData.description = (postData.description || '') + urlsText;
      }
    }

    if (!courseId || !postType) {
      // 必要な情報がない場合はスキップ
      return;
    }

    if (postId) {
      // 投稿IDが指定されている場合
      var existingPost = getPostById(courseId, postId, postType);
      if (existingPost) {
        // 投稿IDが重複している場合、更新する
        updateExisting = true;
      } else {
        // 投稿IDが存在しない場合、処理を行わない
        skipPost = true;
      }
    } else {
      // 投稿IDが空欄の場合、新規投稿
      updateExisting = false;
    }

    if (skipPost) {
      return;
    }

    // 投稿種別に応じて処理を行う
    if (postType === '課題') {
      postData.workType = 'ASSIGNMENT';
      if (updateExisting) {
        Classroom.Courses.CourseWork.patch(postData, courseId, postId);
      } else {
        Classroom.Courses.CourseWork.create(postData, courseId);
      }
    } else if (postType === '質問') {
      postData.workType = 'SHORT_ANSWER_QUESTION'; // 質問の種類を指定
      if (updateExisting) {
        Classroom.Courses.CourseWork.patch(postData, courseId, postId);
      } else {
        Classroom.Courses.CourseWork.create(postData, courseId);
      }
    } else if (postType === '資料' || postType === 'お知らせ') {
      if (updateExisting) {
        Classroom.Courses.CourseWorkMaterials.patch(postData, courseId, postId);
      } else {
        Classroom.Courses.CourseWorkMaterials.create(postData, courseId);
      }
    }
  });

  SpreadsheetApp.getUi().alert('再投稿が完了しました。');
}


/**
 * クラス名またはIDからCourse IDを取得する関数
 * @param {string} input - クラス名またはCourse ID
 * @return {string} Course ID
 */
function getCourseIdByNameOrId(input) {
  var optionalArgs = {
    'teacherId': 'me'
  };
  var courses = Classroom.Courses.list(optionalArgs).courses;
  if (courses && courses.length > 0) {
    for (var i = 0; i < courses.length; i++) {
      if (courses[i].name === input || courses[i].id === input) {
        return courses[i].id;
      }
    }
  }
  // クラスが見つからない場合は入力をそのまま返す（IDとして扱う）
  return input;
}

/**
 * トピック名またはIDからトピックIDを取得する関数
 * @param {string} courseId - コースID
 * @param {string} input - トピック名またはトピックID
 * @return {string} トピックID
 */
function getTopicIdByNameOrId(courseId, input) {
  var topics = Classroom.Courses.Topics.list(courseId).topic || [];
  for (var i = 0; i < topics.length; i++) {
    if (topics[i].name === input || topics[i].topicId === input) {
      return topics[i].topicId;
    }
  }
  // トピックが存在しない場合は新規作成
  var newTopic = Classroom.Courses.Topics.create({name: input}, courseId);
  return newTopic.topicId;
}

/**
 * 投稿IDから既存の投稿を取得する関数
 * @param {string} courseId - コースID
 * @param {string} postId - 投稿ID
 * @param {string} postType - 投稿種別
 * @return {Object|null} 投稿オブジェクトまたはnull
 */
function getPostById(courseId, postId, postType) {
  try {
    if (postType === '課題' || postType === '質問') {
      var post = Classroom.Courses.CourseWork.get(courseId, postId);
      return post;
    } else if (postType === '資料' || postType === 'お知らせ') {
      var post = Classroom.Courses.CourseWorkMaterials.get(courseId, postId);
      return post;
    }
  } catch (e) {
    // エラーが発生した場合はnullを返す
    return null;
  }
  return null;
}

/**
 * URLが有効かどうかを確認する関数
 * @param {string} url - チェックするURL
 * @return {boolean} 有効なURLであればtrue、そうでなければfalse
 */
function isValidUrl(url) {
  try {
    new URL(url);
    return true;
  } catch (_) {
    return false;
  }
}

/**
 * URLがYouTubeのURLかどうかを確認する関数
 * @param {string} url - チェックするURL
 * @return {boolean} YouTubeのURLであればtrue、そうでなければfalse
 */
function isYouTubeUrl(url) {
  var pattern = /^(https?\:\/\/)?(www\.youtube\.com|youtu\.?be)\/.+$/;
  return pattern.test(url);
}

/**
 * YouTubeのURLから動画IDを抽出する関数
 * @param {string} url - YouTubeのURL
 * @return {string} 動画ID
 */
function extractYouTubeId(url) {
  var videoId = '';
  var regex = /(?:v=|\/)([0-9A-Za-z_-]{11}).*/;
  var match = url.match(regex);
  if (match && match[1]) {
    videoId = match[1];
  }
  return videoId;
}

/**
 * GoogleドライブのURLからファイルIDを抽出する関数
 * @param {string} url - GoogleドライブのファイルURLまたはファイルID
 * @return {string} ファイルID
 */
function extractFileId(url) {
  var id = '';
  var regex = /[-\w]{25,}/;
  var match = url.match(regex);
  if (match && match[0]) {
    id = match[0];
  } else {
    // URLではなくファイルIDが直接渡された場合
    id = url;
  }
  return id;
}
