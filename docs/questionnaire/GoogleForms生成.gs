// 事後アンケートの Google フォームを作るスクリプト
// docs/questionnaire/設問一覧.xlsx から自動生成。直接ここを直さず、Excel を直して作り直すこと。
//
// 使い方：
//   1. Google ドライブで「新規 → その他 → Google Apps Script」を開く
//   2. このファイルの中身をぜんぶ貼り付ける
//   3. 関数 createAll を選んで実行する（初回は権限の確認が出る）
//   4. 実行ログに5つのフォームのURLが出るので控える
//
// うまくいかないときは、createAll ではなく下の5つを1つずつ実行してもよい：
//   makeDaily1 / makeDaily2_AriFirst / makeDaily2_NashiFirst
//   makeFinal_StepperFirst / makeFinal_NumpadFirst

function createAll() {
  var urls = [];
  urls.push(['①前半（全員・解散直後）', makeDaily1()]);
  urls.push(['①後半 S4 … v1 v2 の人に渡す', makeDaily2_AriFirst()]);
  urls.push(['①後半 S4 … v3 v4 の人に渡す', makeDaily2_NashiFirst()]);
  urls.push(['②まとめ … v1 v3 の人に渡す', makeFinal_StepperFirst()]);
  urls.push(['②まとめ … v2 v4 の人に渡す', makeFinal_NumpadFirst()]);
  for (var i = 0; i < urls.length; i++) {
    Logger.log(urls[i][0] + '\n  回答用: ' + urls[i][1][0] + '\n  編集用: ' + urls[i][1][1]);
  }
}

function makeDaily1() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（前半）');
  f.setDescription('今日の乗車についてお答えください。所要5分ほどです。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addSectionHeaderItem().setTitle('はじめに');
  it.setHelpText('今日の乗車についてお答えください。所要5分ほどです。');
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setHelpText('カードの左上に書いてあります');
  it.setRequired(true);
  it = f.addDateItem().setTitle('0-2 今日の日付');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-3 今日は何日目ですか').setChoiceValues(['1日目', '2日目']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-4 この用紙の版').setChoiceValues(['v1', 'v2', 'v3', 'v4']);
  it.setHelpText('配布カードに書いてあります');
  it.setRequired(true);
  it = f.addTextItem().setTitle('0-5 今日乗った1本目の発車時刻');
  it.setHelpText('覚えている範囲で。空欄でも構いません');
  it.setRequired(false);
  it = f.addTextItem().setTitle('0-6 今日乗った2本目の発車時刻');
  it.setHelpText('覚えている範囲で。空欄でも構いません');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('普段のことについて');
  it.setHelpText('★1日目の方だけお答えください。2日目の方はそのまま次へ進んでください。');
  it = f.addMultipleChoiceItem().setTitle('1-1 年代').setChoiceValues(['10代', '20代', '30代', '40代以上']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('1-2 普段バスを利用する頻度').setChoiceValues(['ほぼ毎日', '週に数回', '月に数回', 'ほとんど乗らない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('1-3 今回の路線に乗ったことがありましたか').setChoiceValues(['よく乗る', '数回ある', '初めて']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('1-4 実験で使った端末').setChoiceValues(['iPhone', 'Android', 'その他']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('1-5 利き手').setChoiceValues(['右', '左']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('1-6 実験中、スマホは主にどちらの手で操作しましたか').setChoiceValues(['片手（親指）', '両手', 'その時による']);
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('今日、人数を答えるときのことについて');
  it.setHelpText('正解・不正解を見るものではありません。思ったとおりにお答えください。');
  it = f.addMultipleChoiceItem().setTitle('3-1 今日、人数を答えるとき、実際に一人ずつ数えていましたか').setChoiceValues(['毎回数えた', 'だいたい数えた', '目分量が多かった', 'ほとんど目分量']);
  it.setRequired(true);
  it = f.addCheckboxItem().setTitle('3-2 数えるのが難しいと感じたのはどんなときですか（いくつでも）').setChoiceValues(['混んでいた', '自分が立っていた', '揺れた', '降車が近かった', '暗かった', 'その他']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('3-3 今日、数えきれていないのに、それらしい数を入れて送ったことがありますか').setChoiceValues(['何度もある', '数回ある', 'ない', '覚えていない']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('3-4 （3-3で「ある」と答えた方）そうしたのはどんなときですか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('3-5 人数を入れずに空欄のまま送ったことがありますか。あればその理由も');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('3-6 今日の2便のあいだで、答えやすさは違いましたか').setBounds(1, 5).setLabels('違わない', '大きく違った');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('3-6b （3-6について）どう違ったか、よければ教えてください');
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('3-7 今日、座っていた割合はどのくらいですか').setChoiceValues(['ほぼ座席', '半々', 'ほぼ立席']);
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function makeDaily2_AriFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（後半）A');
  f.setDescription('前半を送信された方にお渡しするものです。引き続きお答えください。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setHelpText('前半と同じIDを入れてください');
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('今日の「わからない」ボタンについて');
  it.setHelpText('今日の1本目と2本目で、画面に1か所だけ違いがありました。');
  it = f.addMultipleChoiceItem().setTitle('4-1 今日の1本目と2本目で、画面に違いがあったことに気づきましたか').setChoiceValues(['はっきり気づいた', 'なんとなく気づいた', '気づかなかった']);
  it.setRequired(true);
  f.addSectionHeaderItem().setTitle('違いはここです').setHelpText('人数を入れる欄のすぐ下に――');
  f.addSectionHeaderItem().setTitle('【「わからない」ボタン あり】').setHelpText('「わからない」というボタンが置かれていた便');
  f.addSectionHeaderItem().setTitle('【「わからない」ボタン なし】').setHelpText('同じ場所に何も置かれていなかった便');
  it = f.addMultipleChoiceItem().setTitle('4-2 「わからない」ボタンがあったほうの乗車を覚えていますか').setChoiceValues(['1本目', '2本目', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-3 ボタンがあったとき、数えきれないときはどうしていましたか').setChoiceValues(['ボタンを押した', 'それらしい数を入れた', '空欄で送った', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-4 ボタンがなかったとき、数えきれないときはどうしていましたか').setChoiceValues(['それらしい数を入れた', '空欄で送った', 'クラスだけ選んで送った', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-5 どちらのほうが正直に答えられたと感じますか').setChoiceValues(['ボタンがあった便', 'ボタンがなかった便', '変わらない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('4-6 そう思う理由');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('4-7 「わからない」と答えるのに、抵抗や後ろめたさはありましたか').setBounds(1, 5).setLabels('まったくなかった', 'とてもあった');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('4-7b （4-7について）よければ理由を教えてください');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function makeDaily2_NashiFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（後半）B');
  f.setDescription('前半を送信された方にお渡しするものです。引き続きお答えください。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setHelpText('前半と同じIDを入れてください');
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('今日の「わからない」ボタンについて');
  it.setHelpText('今日の1本目と2本目で、画面に1か所だけ違いがありました。');
  it = f.addMultipleChoiceItem().setTitle('4-1 今日の1本目と2本目で、画面に違いがあったことに気づきましたか').setChoiceValues(['はっきり気づいた', 'なんとなく気づいた', '気づかなかった']);
  it.setRequired(true);
  f.addSectionHeaderItem().setTitle('違いはここです').setHelpText('人数を入れる欄のすぐ下に――');
  f.addSectionHeaderItem().setTitle('【「わからない」ボタン なし】').setHelpText('同じ場所に何も置かれていなかった便');
  f.addSectionHeaderItem().setTitle('【「わからない」ボタン あり】').setHelpText('「わからない」というボタンが置かれていた便');
  it = f.addMultipleChoiceItem().setTitle('4-2 「わからない」ボタンがあったほうの乗車を覚えていますか').setChoiceValues(['1本目', '2本目', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-3 ボタンがあったとき、数えきれないときはどうしていましたか').setChoiceValues(['ボタンを押した', 'それらしい数を入れた', '空欄で送った', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-4 ボタンがなかったとき、数えきれないときはどうしていましたか').setChoiceValues(['それらしい数を入れた', '空欄で送った', 'クラスだけ選んで送った', '覚えていない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('4-5 どちらのほうが正直に答えられたと感じますか').setChoiceValues(['ボタンがなかった便', 'ボタンがあった便', '変わらない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('4-6 そう思う理由');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('4-7 「わからない」と答えるのに、抵抗や後ろめたさはありましたか').setBounds(1, 5).setLabels('まったくなかった', 'とてもあった');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('4-7b （4-7について）よければ理由を教えてください');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function makeFinal_StepperFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― まとめ A');
  f.setDescription('2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addSectionHeaderItem().setTitle('はじめに');
  it.setHelpText('2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。');
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-2 この用紙の版').setChoiceValues(['v1', 'v2', 'v3', 'v4']);
  it.setHelpText('配布カードに書いてあります');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('2つの入力方式について');
  it.setHelpText('2日間で、人数の入れ方が日によって違いました。次の2つです。');
  it = f.addMultipleChoiceItem().setTitle('2-0 1日目に使ったのはどちらの方式でしたか').setChoiceValues(['ステッパー方式', 'テンキー方式', '覚えていない']);
  it.setHelpText('思い出せる範囲で構いません');
  it.setRequired(false);
  f.addSectionHeaderItem().setTitle('【ステッパー方式】').setHelpText('「−5」「−1」「＋1」「＋5」のボタンで、数を増やしたり減らしたりして合わせる');
  it = f.addScaleItem().setTitle('2-1 ステッパー方式は押しやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-2 ステッパー方式は素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-3 ステッパー方式では自分が思った通りの人数を入力できた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-4 ステッパー方式は頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  f.addSectionHeaderItem().setTitle('【テンキー方式】').setHelpText('数字のキーを押して、人数を直接入力する');
  it = f.addScaleItem().setTitle('2-1 テンキー方式は押しやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-2 テンキー方式は素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-3 テンキー方式では自分が思った通りの人数を入力できた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-4 テンキー方式は頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('2-5 2つのうち、答えやすかったのはどちらですか').setChoiceValues(['ステッパー方式', 'テンキー方式', 'どちらとも言えない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('2-6 それはなぜですか');
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('2-7 揺れている車内では、どちらが押し間違えにくかったですか').setChoiceValues(['ステッパー方式', 'テンキー方式', 'どちらとも言えない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('2-8 それはなぜですか');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('アプリ全体の使いやすさ');
  it.setHelpText('国際的に使われている標準の10項目です。項目の追加・削除・並べ替えはしないでください。');
  it = f.addScaleItem().setTitle('5-1 このアプリを頻繁に使いたいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-2 このアプリは不必要に複雑だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-3 このアプリは簡単に使えると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-4 このアプリを使うには、詳しい人のサポートが必要だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-5 このアプリのさまざまな機能は、うまくまとまっていると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-6 このアプリには一貫性のないところが多いと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-7 たいていの人はこのアプリの使い方をすぐに覚えられると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-8 このアプリはとても使いにくいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-9 このアプリを自信をもって使えた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-10 このアプリを使い始める前に、いろいろ覚える必要があった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('自由にお書きください');
  it = f.addParagraphTextItem().setTitle('6-1 人数を答えるとき、いちばん面倒だと感じたことは何ですか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-2 「次に停まるバス停」を選ぶとき、困ったことはありましたか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-3 このアプリをこう変えたら答えやすくなる、という案があれば');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-4 実験全体について、気づいたことがあれば');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function makeFinal_NumpadFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― まとめ B');
  f.setDescription('2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addSectionHeaderItem().setTitle('はじめに');
  it.setHelpText('2日間すべての乗車を終えた方にお答えいただきます。所要5分ほどです。');
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-2 この用紙の版').setChoiceValues(['v1', 'v2', 'v3', 'v4']);
  it.setHelpText('配布カードに書いてあります');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('2つの入力方式について');
  it.setHelpText('2日間で、人数の入れ方が日によって違いました。次の2つです。');
  it = f.addMultipleChoiceItem().setTitle('2-0 1日目に使ったのはどちらの方式でしたか').setChoiceValues(['ステッパー方式', 'テンキー方式', '覚えていない']);
  it.setHelpText('思い出せる範囲で構いません');
  it.setRequired(false);
  f.addSectionHeaderItem().setTitle('【テンキー方式】').setHelpText('数字のキーを押して、人数を直接入力する');
  it = f.addScaleItem().setTitle('2-1 テンキー方式は押しやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-2 テンキー方式は素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-3 テンキー方式では自分が思った通りの人数を入力できた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-4 テンキー方式は頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  f.addSectionHeaderItem().setTitle('【ステッパー方式】').setHelpText('「−5」「−1」「＋1」「＋5」のボタンで、数を増やしたり減らしたりして合わせる');
  it = f.addScaleItem().setTitle('2-1 ステッパー方式は押しやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-2 ステッパー方式は素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-3 ステッパー方式では自分が思った通りの人数を入力できた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-4 ステッパー方式は頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('2-5 2つのうち、答えやすかったのはどちらですか').setChoiceValues(['テンキー方式', 'ステッパー方式', 'どちらとも言えない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('2-6 それはなぜですか');
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('2-7 揺れている車内では、どちらが押し間違えにくかったですか').setChoiceValues(['テンキー方式', 'ステッパー方式', 'どちらとも言えない']);
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('2-8 それはなぜですか');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('アプリ全体の使いやすさ');
  it.setHelpText('国際的に使われている標準の10項目です。項目の追加・削除・並べ替えはしないでください。');
  it = f.addScaleItem().setTitle('5-1 このアプリを頻繁に使いたいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-2 このアプリは不必要に複雑だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-3 このアプリは簡単に使えると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-4 このアプリを使うには、詳しい人のサポートが必要だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-5 このアプリのさまざまな機能は、うまくまとまっていると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-6 このアプリには一貫性のないところが多いと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-7 たいていの人はこのアプリの使い方をすぐに覚えられると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-8 このアプリはとても使いにくいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-9 このアプリを自信をもって使えた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('5-10 このアプリを使い始める前に、いろいろ覚える必要があった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('自由にお書きください');
  it = f.addParagraphTextItem().setTitle('6-1 人数を答えるとき、いちばん面倒だと感じたことは何ですか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-2 「次に停まるバス停」を選ぶとき、困ったことはありましたか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-3 このアプリをこう変えたら答えやすくなる、という案があれば');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('6-4 実験全体について、気づいたことがあれば');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}
