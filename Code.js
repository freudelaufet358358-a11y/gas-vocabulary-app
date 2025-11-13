// ===================================
// グローバル設定
// ===================================
var SPREADSHEET_ID = ''; // 空欄の場合は現在のスプレッドシートを使用

function getSpreadsheet() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// ===================================
// Webアプリのエントリーポイント
// ===================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('英単語学習アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===================================
// 単語帳の種類を取得
// ===================================
function getAvailableBooks() {
  var ss = getSpreadsheet();
  var books = [];
  
  // ターゲット1900（単語リスト）
  if (ss.getSheetByName('単語リスト')) {
    books.push({
      id: 'target1900',
      name: 'ターゲット1900',
      sheetName: '単語リスト',
      logSheetName: 'AllLog',
      type: 'word'
    });
  }
  
  // 速読熟語（熟語リスト）
  if (ss.getSheetByName('熟語リスト')) {
    books.push({
      id: 'sokudoku',
      name: '速読熟語',
      sheetName: '熟語リスト',
      logSheetName: '熟語学習ログ',
      type: 'idiom'
    });
  }
  
  return books;
}

// ===================================
// 単語リストシートの初期化と更新
// ===================================
function initializeWordListSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = ss.insertSheet(sheetName);
    
    // ヘッダー行を追加
    var headers = ['WordID', 'English', 'Japanese', '学習済み', '正解数', '不正解数'];
    sheet.appendRow(headers);
    
    // ヘッダー行を装飾
    sheet.getRange(1, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    // サンプルデータを追加（熟語リストの場合）
    if (sheetName === '熟語リスト') {
      var sampleData = [
        [1, 'in order to', '〜するために'],
        [2, 'look forward to', '〜を楽しみにする'],
        [3, 'take care of', '〜の世話をする'],
        [4, 'as soon as', '〜するとすぐに'],
        [5, 'make use of', '〜を利用する']
      ];
      
      for (var i = 0; i < sampleData.length; i++) {
        sheet.appendRow(sampleData[i]);
      }
    }
    
    Logger.log(sheetName + 'を作成しました。');
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var needsUpdate = false;
  
  // 必要な列が存在するか確認
  if (headers.length < 3 || headers[0] !== 'WordID' || headers[1] !== 'English' || headers[2] !== 'Japanese') {
    throw new Error(sheetName + 'のフォーマットが正しくありません。A列:WordID, B列:English, C列:Japaneseが必要です。');
  }
  
  // D列（学習済み）が存在しない場合は追加
  if (headers.length < 4 || headers[3] !== '学習済み') {
    sheet.getRange(1, 4).setValue('学習済み');
    needsUpdate = true;
  }
  
  // E列（正解数）が存在しない場合は追加
  if (headers.length < 5 || headers[4] !== '正解数') {
    sheet.getRange(1, 5).setValue('正解数');
    needsUpdate = true;
  }
  
  // F列（不正解数）が存在しない場合は追加
  if (headers.length < 6 || headers[5] !== '不正解数') {
    sheet.getRange(1, 6).setValue('不正解数');
    needsUpdate = true;
  }
  
  if (needsUpdate) {
    // ヘッダー行を装飾
    sheet.getRange(1, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    Logger.log(sheetName + 'に必要な列を追加しました。');
  }
  
  return sheet;
}

// ===================================
// AllLogから統計を更新
// ===================================
function updateWordStatisticsFromAllLog(bookId) {
  try {
    var ss = getSpreadsheet();
    var books = getAvailableBooks();
    var book = books.find(b => b.id === bookId) || books[0];
    
    var allLogSheet = ss.getSheetByName(book.logSheetName);
    var wordListSheet = initializeWordListSheet(book.sheetName);
    
    if (!allLogSheet) {
      Logger.log(book.logSheetName + 'シートが見つかりません。統計の更新をスキップします。');
      return { success: true, message: book.logSheetName + 'シートがないため統計更新をスキップしました。' };
    }
    
    // AllLogからデータを取得
    var allLogData = allLogSheet.getDataRange().getValues();
    
    if (allLogData.length <= 1) {
      Logger.log(book.logSheetName + 'にデータがありません。');
      return { success: true, message: book.logSheetName + 'にデータがありません。' };
    }
    
    // 単語リストからデータを取得
    var wordListData = wordListSheet.getDataRange().getValues();
    var wordStats = {};
    
    // 各単語の統計を初期化（0から開始 - 既存のカウントは無視）
    for (var i = 1; i < wordListData.length; i++) {
      var wordId = wordListData[i][0];
      if (wordId) {
        wordStats[wordId] = {
          row: i + 1,
          correct: 0,  // 常に0から開始
          incorrect: 0  // 常に0から開始
        };
      }
    }
    
    // AllLogから統計を完全に再計算
    for (var i = 1; i < allLogData.length; i++) {
      var result = String(allLogData[i][1]).trim(); // 正誤
      var wordId = allLogData[i][2]; // WordID
      
      if (wordId && wordStats[wordId]) {
        if (result === '正解') {
          wordStats[wordId].correct++;
        } else if (result === '不正解') {
          wordStats[wordId].incorrect++;
        }
      }
    }
    
    // 単語リストシートに統計を書き込み
    var updates = [];
    for (var wordId in wordStats) {
      var stat = wordStats[wordId];
      updates.push({
        row: stat.row,
        correct: stat.correct,
        incorrect: stat.incorrect,
        learned: stat.correct > 0 // 1回以上正解したら学習済み
      });
    }
    
    // バッチ更新
    updates.forEach(function(update) {
      wordListSheet.getRange(update.row, 4).setValue(update.learned ? '○' : ''); // 学習済み
      wordListSheet.getRange(update.row, 5).setValue(update.correct); // 正解数
      wordListSheet.getRange(update.row, 6).setValue(update.incorrect); // 不正解数
    });
    
    Logger.log(book.sheetName + 'の統計を更新しました。更新件数: ' + updates.length);
    
    // 全体統計も計算して返す
    var stats = calculateStatisticsFromAllLog(book.logSheetName);
    
    return {
      success: true,
      message: '統計を更新しました。更新件数: ' + updates.length,
      updatedCount: updates.length,
      stats: stats
    };
    
  } catch (error) {
    Logger.log('統計更新エラー: ' + error.message);
    return {
      success: false,
      message: '統計更新エラー: ' + error.message
    };
  }
}

// ===================================
// 単語データの取得
// ===================================
function getWordLists(bookId) {
  try {
    var books = getAvailableBooks();
    var book = books.find(b => b.id === bookId) || books[0];
    
    if (!book) {
      return {
        success: false,
        message: '指定された単語帳が見つかりません。'
      };
    }
    
    var sheet = initializeWordListSheet(book.sheetName);
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: false,
        message: book.name + 'にデータがありません。'
      };
    }
    
    var words = [];
    for (var i = 1; i < data.length; i++) {
      // IDと英語と日本語が存在するかチェック（falseという単語も含める）
      var hasId = data[i][0] !== undefined && data[i][0] !== null && data[i][0] !== '';
      var hasEnglish = data[i][1] !== undefined && data[i][1] !== null && String(data[i][1]).trim() !== '';
      var hasJapanese = data[i][2] !== undefined && data[i][2] !== null && String(data[i][2]).trim() !== '';
      
      if (hasId && hasEnglish && hasJapanese) {
        words.push({
          id: Number(data[i][0]),
          english: String(data[i][1]).trim(),
          japanese: String(data[i][2]).trim(),
          learned: data[i][3] === true || data[i][3] === 'TRUE' || data[i][3] === '○',
          correctCount: Number(data[i][4]) || 0,
          incorrectCount: Number(data[i][5]) || 0
        });
      }
    }
    
    // 軽量な統計計算（単語リストから計算）
    var stats = calculateStatisticsFromWordList(words, book.logSheetName);
    
    // 過去30問で出題された単語IDを取得
    var recentWordIds = getRecentlyTestedWordIds(30, book.logSheetName);
    
    return {
      success: true,
      words: words,
      stats: stats,
      recentlyTestedIds: recentWordIds,
      bookInfo: book,
      availableBooks: books
    };
    
  } catch (error) {
    Logger.log('getWordLists エラー: ' + error.message);
    return {
      success: false,
      message: 'データの取得に失敗しました: ' + error.message
    };
  }
}

// ===================================
// 過去N問で出題された単語IDを取得（最適化版）
// ===================================
function getRecentlyTestedWordIds(count, logSheetName) {
  try {
    var sheet = getOrCreateAllLogSheet(logSheetName);
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    // 最後のN行だけを取得（全データを読まないため高速）
    var startRow = Math.max(2, lastRow - count + 1);
    var numRows = lastRow - startRow + 1;
    
    if (numRows <= 0) {
      return [];
    }
    
    var data = sheet.getRange(startRow, 1, numRows, 8).getValues();
    
    var recentWordIds = [];
    for (var i = data.length - 1; i >= 0 && recentWordIds.length < count; i--) {
      var wordId = data[i][2]; // WordID列（3列目、インデックス2）
      if (wordId) {
        recentWordIds.push(Number(wordId));
      }
    }
    
    return recentWordIds;
    
  } catch (error) {
    Logger.log('getRecentlyTestedWordIds エラー: ' + error.message);
    return [];
  }
}

// ===================================
// AllLogシートの取得・作成
// ===================================
function getOrCreateAllLogSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['LogID', '正誤', 'WordID', 'English', 'Japanese', 'Direction', 'UserAnswer', 'Date/Time']);
    
    var headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setFontWeight('bold')
               .setBackground('#4285f4')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    Logger.log(sheetName + 'シートを作成しました。');
  }
  
  return sheet;
}

// ===================================
// テスト結果の1問ごとの記録
// ===================================
function logTestResult(wordId, isCorrect, direction, userAnswer, english, japanese, bookId) {
  try {
    var books = getAvailableBooks();
    var book = books.find(b => b.id === bookId) || books[0];
    var sheet = getOrCreateAllLogSheet(book.logSheetName);
    
    var lastRow = sheet.getLastRow();
    var lastLogId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
    var newLogId = lastLogId + 1;
    
    var resultText = isCorrect ? '正解' : '不正解';
    var timestamp = new Date();
    
    sheet.appendRow([
      newLogId,
      resultText,
      wordId,
      english,
      japanese,
      direction,
      String(userAnswer),
      timestamp
    ]);
    
    return { success: true };
    
  } catch (error) {
    Logger.log('logTestResult エラー: ' + error.message);
    return { success: false, message: error.message };
  }
}

// ===================================
// テスト結果の保存（まとめて処理）
// ===================================
function saveTestResults(resultsData, bookId) {
  try {
    var ss = getSpreadsheet();
    var books = getAvailableBooks();
    var book = books.find(b => b.id === bookId) || books[0];
    
    var wordListSheet = initializeWordListSheet(book.sheetName);
    var wordListData = wordListSheet.getDataRange().getValues();
    
    // 単語IDをキーとした辞書を作成
    var wordMap = {};
    for (var i = 1; i < wordListData.length; i++) {
      var wordId = wordListData[i][0];
      if (wordId) {
        wordMap[wordId] = {
          row: i + 1,
          english: wordListData[i][1],
          japanese: wordListData[i][2],
          learned: wordListData[i][3],
          correctCount: Number(wordListData[i][4]) || 0,
          incorrectCount: Number(wordListData[i][5]) || 0
        };
      }
    }
    
    // 各問題の結果を処理
    resultsData.wordResults.forEach(function(result) {
      var wordId = result.wordId;
      var isCorrect = result.correct;
      var direction = result.direction;
      var userAnswer = result.userAnswer;
      
      if (wordMap[wordId]) {
        var word = wordMap[wordId];
        
        // AllLogに記録
        logTestResult(wordId, isCorrect, direction, userAnswer, word.english, word.japanese, bookId);
        
        // 正解数・不正解数を更新
        if (isCorrect) {
          word.correctCount++;
        } else {
          word.incorrectCount++;
        }
        
        // 学習済みフラグ（1回以上正解したら学習済み）
        if (word.correctCount > 0) {
          word.learned = true;
        }
      }
    });
    
    // 単語リストシートを更新
    for (var wordId in wordMap) {
      var word = wordMap[wordId];
      wordListSheet.getRange(word.row, 4).setValue(word.learned ? '○' : '');
      wordListSheet.getRange(word.row, 5).setValue(word.correctCount);
      wordListSheet.getRange(word.row, 6).setValue(word.incorrectCount);
    }
    
    // 統計を計算
    var stats = calculateStatisticsFromAllLog(book.logSheetName);
    updateStatisticsSheet(stats);
    
    // 獲得XPを計算（1問あたり1XP）
    var xpGained = resultsData.correctAnswers;
    
    return {
      success: true,
      stats: stats,
      xpGained: xpGained
    };
    
  } catch (error) {
    Logger.log('saveTestResults エラー: ' + error.message);
    return {
      success: false,
      message: 'テスト結果の保存に失敗しました: ' + error.message
    };
  }
}

// ===================================
// AllLogから統計を計算
// ===================================
function calculateStatisticsFromAllLog(logSheetName) {
  try {
    var sheet = getOrCreateAllLogSheet(logSheetName);
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        totalTests: 0,
        totalCorrect: 0,
        totalIncorrect: 0,
        currentStreak: 0,
        longestStreak: 0,
        lastStudyDate: '',
        level: 1,
        xp: 0
      };
    }
    
    var totalCorrect = 0;
    var totalIncorrect = 0;
    var dates = [];
    
    // データを集計
    for (var i = 1; i < data.length; i++) {
      var result = String(data[i][1]).trim();
      var timestamp = data[i][7];
      
      if (result === '正解') {
        totalCorrect++;
      } else if (result === '不正解') {
        totalIncorrect++;
      }
      
      if (timestamp) {
        var dateOnly = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (dates.indexOf(dateOnly) === -1) {
          dates.push(dateOnly);
        }
      }
    }
    
    // 日付を昇順にソート
    dates.sort();
    
    // 連続学習日数を計算
    var currentStreak = 0;
    var longestStreak = 0;
    var tempStreak = 0;
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // 現在のストリークを計算
    if (dates.length > 0) {
      var lastDate = dates[dates.length - 1];
      if (lastDate === today || lastDate === yesterday) {
        currentStreak = 1;
        for (var i = dates.length - 2; i >= 0; i--) {
          var currentDate = new Date(dates[i + 1]);
          var prevDate = new Date(dates[i]);
          var diffDays = Math.floor((currentDate - prevDate) / 86400000);
          
          if (diffDays === 1) {
            currentStreak++;
          } else {
            break;
          }
        }
      }
    }
    
    // 最長ストリークを計算
    if (dates.length > 0) {
      tempStreak = 1;
      longestStreak = 1;
      
      for (var i = 1; i < dates.length; i++) {
        var currentDate = new Date(dates[i]);
        var prevDate = new Date(dates[i - 1]);
        var diffDays = Math.floor((currentDate - prevDate) / 86400000);
        
        if (diffDays === 1) {
          tempStreak++;
          longestStreak = Math.max(longestStreak, tempStreak);
        } else {
          tempStreak = 1;
        }
      }
    }
    
    // XPとレベルの計算
    var xp = totalCorrect * 10 + Math.ceil((totalCorrect + totalIncorrect) / 10) * 2;
    var level = Math.floor(xp / 100) + 1;
    
    // 総テスト数（10問を1テストとして計算）
    var totalTests = Math.ceil((totalCorrect + totalIncorrect) / 10);
    
    return {
      totalTests: totalTests,
      totalCorrect: totalCorrect,
      totalIncorrect: totalIncorrect,
      currentStreak: currentStreak,
      longestStreak: Math.max(longestStreak, currentStreak),
      lastStudyDate: dates.length > 0 ? dates[dates.length - 1] : '',
      level: level,
      xp: xp
    };
    
  } catch (error) {
    Logger.log('calculateStatisticsFromAllLog エラー: ' + error.message);
    return {
      totalTests: 0,
      totalCorrect: 0,
      totalIncorrect: 0,
      currentStreak: 0,
      longestStreak: 0,
      lastStudyDate: '',
      level: 1,
      xp: 0
    };
  }
}

// ===================================
// 単語リストから統計を軽量計算（高速化版）
// ===================================
function calculateStatisticsFromWordList(words, logSheetName) {
  try {
    // 単語リストの統計から集計（AllLogを読まないため高速）
    var totalCorrect = 0;
    var totalIncorrect = 0;
    
    for (var i = 0; i < words.length; i++) {
      totalCorrect += words[i].correctCount || 0;
      totalIncorrect += words[i].incorrectCount || 0;
    }
    
    // XPとレベルを計算（calculateStatisticsFromAllLogと同じ計算式）
    var xp = totalCorrect * 10 + Math.ceil((totalCorrect + totalIncorrect) / 10) * 2;
    var level = Math.floor(xp / 100) + 1;
    
    // 連続記録を効率的に計算
    var streakData = calculateStreakEfficiently(logSheetName);
    
    return {
      totalTests: 0, // テスト回数はAllLogから計算する必要がある
      totalCorrect: totalCorrect,
      totalIncorrect: totalIncorrect,
      currentStreak: streakData.currentStreak,
      longestStreak: streakData.longestStreak,
      lastStudyDate: streakData.lastStudyDate,
      level: level,
      xp: xp
    };
    
  } catch (error) {
    Logger.log('統計計算エラー（軽量版）: ' + error.message);
    return {
      totalTests: 0,
      totalCorrect: 0,
      totalIncorrect: 0,
      currentStreak: 0,
      longestStreak: 0,
      lastStudyDate: '',
      level: 1,
      xp: 0
    };
  }
}

// ===================================
// 連続記録を効率的に計算（軽量版）
// ===================================
function calculateStreakEfficiently(logSheetName) {
  try {
    var sheet = getOrCreateAllLogSheet(logSheetName);
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        currentStreak: 0,
        longestStreak: 0,
        lastStudyDate: ''
      };
    }
    
    // 直近100行のタイムスタンプのみ取得（効率化）
    var numRows = Math.min(100, lastRow - 1);
    var startRow = lastRow - numRows + 1;
    var timestamps = sheet.getRange(startRow, 8, numRows, 1).getValues(); // 8列目はタイムスタンプ
    
    // 日付のみ抽出（重複除去）
    var dates = [];
    var dateMap = {};
    
    for (var i = 0; i < timestamps.length; i++) {
      if (timestamps[i][0]) {
        var dateOnly = Utilities.formatDate(new Date(timestamps[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (!dateMap[dateOnly]) {
          dateMap[dateOnly] = true;
          dates.push(dateOnly);
        }
      }
    }
    
    // 日付を昇順にソート
    dates.sort();
    
    if (dates.length === 0) {
      return {
        currentStreak: 0,
        longestStreak: 0,
        lastStudyDate: ''
      };
    }
    
    // 現在のストリークを計算
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var currentStreak = 0;
    var longestStreak = 0;
    
    var lastDate = dates[dates.length - 1];
    
    // 今日または昨日に学習している場合のみ連続記録を計算
    if (lastDate === today || lastDate === yesterday) {
      currentStreak = 1;
      
      // 連続日数を計算
      for (var i = dates.length - 2; i >= 0; i--) {
        var currentDate = new Date(dates[i + 1]);
        var prevDate = new Date(dates[i]);
        var diffDays = Math.floor((currentDate - prevDate) / 86400000);
        
        if (diffDays === 1) {
          currentStreak++;
        } else {
          break;
        }
      }
    }
    
    // 最長ストリークを計算
    var tempStreak = 1;
    for (var i = 1; i < dates.length; i++) {
      var currentDate = new Date(dates[i]);
      var prevDate = new Date(dates[i - 1]);
      var diffDays = Math.floor((currentDate - prevDate) / 86400000);
      
      if (diffDays === 1) {
        tempStreak++;
        longestStreak = Math.max(longestStreak, tempStreak);
      } else {
        tempStreak = 1;
      }
    }
    
    longestStreak = Math.max(longestStreak, currentStreak);
    
    return {
      currentStreak: currentStreak,
      longestStreak: longestStreak,
      lastStudyDate: lastDate
    };
    
  } catch (error) {
    Logger.log('連続記録計算エラー: ' + error.message);
    return {
      currentStreak: 0,
      longestStreak: 0,
      lastStudyDate: ''
    };
  }
}

// ===================================
// 統計シートの更新
// ===================================
function updateStatisticsSheet(stats) {
  try {
    var ss = getSpreadsheet();
    var statsSheet = ss.getSheetByName('統計');
    
    if (!statsSheet) {
      statsSheet = ss.insertSheet('統計');
      var headers = ['総テスト数', '正解数', '不正解数', '現在のストリーク', '最長ストリーク', '最終学習日', 'レベル', '経験値'];
      statsSheet.appendRow(headers);
      
      var headerRange = statsSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold')
                 .setBackground('#4285f4')
                 .setFontColor('#ffffff')
                 .setHorizontalAlignment('center');
    }
    
    var statsRow = [
      stats.totalTests,
      stats.totalCorrect,
      stats.totalIncorrect,
      stats.currentStreak,
      stats.longestStreak,
      stats.lastStudyDate,
      stats.level,
      stats.xp
    ];
    
    if (statsSheet.getLastRow() < 2) {
      statsSheet.appendRow(statsRow);
    } else {
      statsSheet.getRange(2, 1, 1, 8).setValues([statsRow]);
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log('updateStatisticsSheet エラー: ' + error.message);
    return { success: false, message: error.message };
  }
}

// ===================================
// AllLogシートのデータをエクスポートする関数（オプション）
// ===================================
function exportAllLog() {
  try {
    var sheet = getOrCreateAllLogSheet('AllLog');
    var data = sheet.getDataRange().getValues();
    
    return {
      success: true,
      data: data
    };
  } catch (error) {
    Logger.log('AllLogエクスポートエラー: ' + error.message);
    return {
      success: false,
      message: 'AllLogのエクスポートに失敗しました: ' + error.message
    };
  }
}