// UI研究のアンケート（Google フォーム4つ）と、回答がたまるスプレッドシートを作るスクリプト
// docs/questionnaire/設問一覧.xlsx から tools/make_forms.py で作る。設問は Excel を直して作り直す。
//
// 1. Google ドライブ →「新規 → その他 → Google Apps Script」
// 2. この中身をぜんぶ貼り付けて保存
// 3. 関数 whoami を実行し、ログに出たアカウントを確かめる（フォームとシートはこのアカウントのドライブにできる）
// 4. 関数 createAll を実行（1〜2分かかる）
// 5. 実行ログの「==== ここから下を…」の行から下をすべてコピーしてチャットに貼る
// 6. 回答用URLを Google にログインしていないブラウザで開き、答えられるか確かめる
//
// ★createAll を2回実行すると、フォーム4つと回答シートがもう1組できる（いらないほうはゴミ箱に入れる）。
// ★前に作った2つのフォーム（1回目・2回目のアンケート）と回答表「事後アンケート回答」は使わない。

function whoami() {
  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());
}

var NAMES = ['1回目_桃山台', '1回目_南口', '2回目_桃山台', '2回目_南口'];
var SHEET_TITLE = 'UI研究 アンケート回答';

function createAll() {
  var makers = [make1Momoyamadai_, make1Minamiguchi_, make2Momoyamadai_, make2Minamiguchi_];
  var forms = [], ss = null;
  try {
    // 作る関数は FormApp.create の直後にフォームを forms に入れる。途中で止まっても、作りかけを下のログに出せる
    for (var i = 0; i < makers.length; i++) {
      makers[i](forms);
    }
    ss = SpreadsheetApp.create(SHEET_TITLE);
    for (var j = 0; j < forms.length; j++) {
      forms[j].setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      var sh = findSheet_(ss.getId(), forms[j]);
      if (!sh) {
        throw new Error('「' + NAMES[j] + '」の回答シートが見つかりません');
      }
      sh.setName(NAMES[j]);
    }
    tidySheets_(ss.getId());
  } catch (e) {
    var msg = ['途中で止まりました: ' + e.message,
               '作りかけのものはゴミ箱に入れてから、もう一度 createAll を実行してください。'];
    for (var k = 0; k < forms.length; k++) {
      msg.push('  ' + NAMES[k] + '（作りかけ） ' + forms[k].getEditUrl());
    }
    if (ss) {
      msg.push('  回答シート（作りかけ） ' + ss.getUrl());
    }
    Logger.log(msg.join('\n'));
    throw e;
  }
  // 1回の Logger.log にまとめる（URL の行ごとにログの時刻が付かないように）。区切りはタブ
  var lines = ['==== ここから下をすべてコピーしてチャットに貼ってください ===='];
  for (var n = 0; n < forms.length; n++) {
    lines.push(NAMES[n] + '\t' + forms[n].getPublishedUrl() + '\t' + forms[n].getEditUrl());
  }
  lines.push('回答シート\t' + ss.getUrl());
  Logger.log(lines.join('\n'));
}

// 回答先にしたときに自動でできるシートを探す。すぐには見えないことがあるので、数回待つ
function findSheet_(ssId, form) {
  for (var n = 0; n < 10; n++) {
    SpreadsheetApp.flush();
    var sheets = SpreadsheetApp.openById(ssId).getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (isSheetOf_(sheets[i], form)) {
        return sheets[i];
      }
    }
    Utilities.sleep(2000);
  }
  return null;
}

// シートにつながったフォームの URL に、フォームの ID（または回答用URLの ID）が含まれるか
function isSheetOf_(sheet, form) {
  var url = sheet.getFormUrl();
  if (!url) {
    return false;
  }
  if (url.indexOf(form.getId()) >= 0) {
    return true;
  }
  var m = form.getPublishedUrl().match(/\/forms\/d\/e\/([^\/?#]+)/);
  return !!(m && url.indexOf(m[1]) >= 0);
}

// シートを NAMES の順に並べ、最初からある空のシート（「シート1」など）を消す。
// 4枚がそろってから消すので、最後の1枚を消すことにはならない
function tidySheets_(ssId) {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.openById(ssId);
  for (var i = 0; i < NAMES.length; i++) {
    ss.setActiveSheet(ss.getSheetByName(NAMES[i]));
    ss.moveActiveSheet(i + 1);
  }
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    var sh = sheets[j];
    if (NAMES.indexOf(sh.getName()) < 0 && !sh.getFormUrl() && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  }
}

// 1回目_桃山台
function make1Momoyamadai_(forms) {
  var f = FormApp.create('バスの人数アプリ　1回目　ゆきのアンケート');
  forms.push(f);
  f.setDescription('桃山台駅で降りたら答えてください（1分ほど）。思ったとおりに選んでください。');
  f.setConfirmationMessage('ありがとうございました。かえりのバスでも、バス停を発車するたびに人数を答えてください。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問2 ゆきのバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 ゆきのバスで人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 ゆきのバスで、数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 ゆきで使った数字キーの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('0〜9 のキーで人数を打つ画面');
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問6 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
}

// 1回目_南口
function make1Minamiguchi_(forms) {
  var f = FormApp.create('バスの人数アプリ　1回目　かえりのアンケート');
  forms.push(f);
  f.setDescription('南口に着いたら答えてください（4分ほど）。思ったとおりに選んでください。');
  f.setConfirmationMessage('ご協力ありがとうございました。今回はこれで終わりです。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問2 かえりのバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 かえりのバスで人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 かえりのバスで、数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 かえりで使った＋−ボタンの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('−5・−1・＋1・＋5 で人数を合わせる画面');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('アプリ全体について');
  it.setHelpText('今回の往復で使ってみた、アプリ全体の印象をお答えください。');
  it = f.addScaleItem().setTitle('問6 このアプリを何度も使いたいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問7 このアプリは必要以上に複雑だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問8 このアプリは簡単に使えると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問9 このアプリを使うには、詳しい人の助けが必要だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問10 このアプリのいろいろな機能は、うまくまとまっていると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問11 このアプリには一貫性のないところが多いと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問12 たいていの人は、このアプリの使い方をすぐ覚えられると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問13 このアプリはとても使いにくいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問14 このアプリを自信をもって使えた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問15 このアプリを使う前に、いろいろ覚える必要があった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問16 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
}

// 2回目_桃山台
function make2Momoyamadai_(forms) {
  var f = FormApp.create('バスの人数アプリ　2回目　ゆきのアンケート');
  forms.push(f);
  f.setDescription('桃山台駅で降りたら答えてください（1分ほど）。思ったとおりに選んでください。');
  f.setConfirmationMessage('ありがとうございました。かえりのバスでも、バス停を発車するたびに人数を答えてください。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問2 ゆきのバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 ゆきのバスで人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 ゆきのバスで、数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 ゆきで使った＋−ボタンの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('−5・−1・＋1・＋5 で人数を合わせる画面');
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問6 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
}

// 2回目_南口
function make2Minamiguchi_(forms) {
  var f = FormApp.create('バスの人数アプリ　2回目　かえりのアンケート');
  forms.push(f);
  f.setDescription('南口に着いたら答えてください（5分ほど）。思ったとおりに選んでください。');
  f.setConfirmationMessage('ご協力ありがとうございました。これで終わりです。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問2 かえりのバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 かえりのバスで人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 かえりのバスで、数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 かえりで使った数字キーの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('0〜9 のキーで人数を打つ画面');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('アプリ全体について');
  it.setHelpText('今回の往復で使ってみた、アプリ全体の印象をお答えください。');
  it = f.addScaleItem().setTitle('問6 このアプリを何度も使いたいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問7 このアプリは必要以上に複雑だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問8 このアプリは簡単に使えると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問9 このアプリを使うには、詳しい人の助けが必要だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問10 このアプリのいろいろな機能は、うまくまとまっていると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問11 このアプリには一貫性のないところが多いと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問12 たいていの人は、このアプリの使い方をすぐ覚えられると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問13 このアプリはとても使いにくいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問14 このアプリを自信をもって使えた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問15 このアプリを使う前に、いろいろ覚える必要があった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('「わからない」ボタンについて');
  it.setHelpText('1回目と2回目のどちらか一方だけ、人数を入れるところの下に「わからない」ボタンがありました。');
  it = f.addMultipleChoiceItem().setTitle('問16 「わからない」ボタンがあったのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '覚えていない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('問17 人数を答えやすかったのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '変わらない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('問18 数えたとおりの人数を答えられたのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '変わらない']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問19 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
}
