// 事後アンケートの Google フォームを作るスクリプト
// docs/questionnaire/設問一覧.xlsx から自動生成。直接ここを直さず、Excel を直して作り直すこと。
//
// 使い方：
//   1. Google ドライブで「新規 → その他 → Google Apps Script」を開く
//   2. このファイルの中身をぜんぶ貼り付ける
//   3. いったん保存する（保存しないと関数の一覧に出ない）
//   4. 関数 createAll を選んで実行する（初回は権限の確認が出る）
//   5. 実行ログに3つのフォームのURLが出るので控える
//
// うまくいかないときは、createAll ではなく下の3つを1つずつ実行してもよい：
//   makeFront / makeBack_AriFirst / makeBack_NashiFirst

function createAll() {
  var urls = [];
  urls.push(['A 前半（全員・毎回）', makeFront()]);
  urls.push(['B 後半 … v1 v2 の人に渡す', makeBack_AriFirst()]);
  urls.push(['B 後半 … v3 v4 の人に渡す', makeBack_NashiFirst()]);
  for (var i = 0; i < urls.length; i++) {
    Logger.log(urls[i][0] + '\n  回答用: ' + urls[i][1][0] + '\n  編集用: ' + urls[i][1][1]);
  }
}

function makeFront() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（前半）');
  f.setDescription('今日の乗車についてお答えください。所要7分ほどです。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addSectionHeaderItem().setTitle('はじめに');
  it.setHelpText('今日の乗車についてお答えください。所要7分ほどです。');
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
  it = f.addPageBreakItem().setTitle('今日の人数の入れ方について');
  it.setHelpText('今日ずっと使っていた入れ方について、そのまま思ったとおりにお答えください。');
  it = f.addMultipleChoiceItem().setTitle('2-0 今日の人数の入れ方はどちらでしたか').setChoiceValues(['ステッパー方式', 'テンキー方式', 'わからない']);
  it.setHelpText('ステッパー方式＝「−5」「−1」「＋1」「＋5」のボタンで数を増減させるもの／テンキー方式＝数字のキーを押して直接入力するもの');
  it.setRequired(true);
  f.addSectionHeaderItem().setTitle('【ステッパー方式】').setHelpText('「−5」「−1」「＋1」「＋5」のボタンで、数を増やしたり減らしたりして合わせる');
  f.addSectionHeaderItem().setTitle('【テンキー方式】').setHelpText('数字のキーを押して、人数を直接入力する');
  it = f.addScaleItem().setTitle('2-1 今日の入れ方は、全体として答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('2-2 今日の入れ方は押しやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-3 今日の入れ方は素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-4 今日の入れ方では、自分が思った通りの人数を入力できた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-5 揺れている車内でも、今日の入れ方は押し間違えにくかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('2-6 今日の入れ方は、頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('2-7 今日の人数の入れ方について、気づいたことがあれば自由にお書きください');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function makeBack_AriFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（後半）A');
  f.setDescription('前半を送信された方にお渡しするものです。引き続きお答えください。2日目の方は、最後に追加の質問があります。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setHelpText('前半と同じIDを入れてください');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-3 今日は何日目ですか').setChoiceValues(['1日目', '2日目']);
  it.setHelpText('★2日目の方は、最後に追加の質問があります');
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
  it = f.addPageBreakItem().setTitle('★2日目の方だけお答えください');
  it.setHelpText('1日目の方は、ここで送信して終わりです。ご協力ありがとうございました。');
  it = f.addSectionHeaderItem().setTitle('アプリ全体の使いやすさ');
  it.setHelpText('2日間を通してのアプリ全体の印象をお答えください。国際的に使われている標準の10項目です。');
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
  it = f.addSectionHeaderItem().setTitle('最後に、自由にお書きください');
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

function makeBack_NashiFirst() {
  var f = FormApp.create('バス車内の人数を答えるアプリについて ― 今日の分（後半）B');
  f.setDescription('前半を送信された方にお渡しするものです。引き続きお答えください。2日目の方は、最後に追加の質問があります。');
  try { f.setCollectEmail(false); } catch (e) {}   // 新しいGoogleフォームでは使えないことがある
  var it;
  it = f.addTextItem().setTitle('0-1 配布カードに書かれたログインID（p01 など）');
  it.setHelpText('前半と同じIDを入れてください');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('0-3 今日は何日目ですか').setChoiceValues(['1日目', '2日目']);
  it.setHelpText('★2日目の方は、最後に追加の質問があります');
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
  it = f.addPageBreakItem().setTitle('★2日目の方だけお答えください');
  it.setHelpText('1日目の方は、ここで送信して終わりです。ご協力ありがとうございました。');
  it = f.addSectionHeaderItem().setTitle('アプリ全体の使いやすさ');
  it.setHelpText('2日間を通してのアプリ全体の印象をお答えください。国際的に使われている標準の10項目です。');
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
  it = f.addSectionHeaderItem().setTitle('最後に、自由にお書きください');
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
