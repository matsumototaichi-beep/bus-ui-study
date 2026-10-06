// 事後アンケートの Google フォームを作るスクリプト
// docs/questionnaire/設問一覧.xlsx から自動生成。Excel を直して作り直すこと。
//
// 1. Google ドライブ →「新規 → その他 → Google Apps Script」
// 2. この中身をぜんぶ貼り付けて保存
// 3. 関数 createAll を実行
// 4. 実行ログに2つのURLが出る
//
// ★先に whoami を実行して、どのアカウントで動くか確かめること。
//   フォームは、ここに出たアカウントの Google ドライブに作られる。

function whoami() {
  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());
}

function createAll() {
  var a = make1(), b = make2();
  Logger.log('1回目\n  回答用: ' + a[0] + '\n  編集用: ' + a[1]);
  Logger.log('2回目\n  回答用: ' + b[0] + '\n  編集用: ' + b[1]);
}

function make1() {
  var f = FormApp.create('バスの人数アプリ ― 1回目');
  f.setDescription('ゆきとかえりの2本に乗ったあとにお答えください。5分ほどで終わります。');
  try { f.setCollectEmail(false); } catch (e) {}
  var it;
  it = f.addSectionHeaderItem().setTitle('1回目のアンケート');
  it.setHelpText('ゆきとかえりの2本に乗ったあとにお答えください。5分ほどで終わります。');
  it = f.addTextItem().setTitle('初-1 カードに書いてあるID');
  it.setHelpText('p01 のような形です');
  it.setRequired(true);
  it = f.addDateItem().setTitle('初-2 今日の日付');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('初-3 今日乗ったのはどちらですか').setChoiceValues(['昼', '夕方']);
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('あなたについて');
  it.setHelpText('1回目だけおうかがいします。');
  it = f.addMultipleChoiceItem().setTitle('個-1 年代').setChoiceValues(['10代', '20代', '30代', '40代以上']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('個-2 普段バスに乗る頻度').setChoiceValues(['ほぼ毎日', '週に数回', '月に数回', 'ほとんど乗らない']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('個-3 この路線に乗ったことがありますか').setChoiceValues(['よく乗る', '数回ある', '初めて']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('個-4 使ったスマホ').setChoiceValues(['iPhone', 'Android', 'その他']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('個-5 スマホはどのように持って操作しましたか').setChoiceValues(['片手（親指）', '両手', 'その時による']);
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('今日、人数を数えたときのこと');
  it.setHelpText('正解・不正解を見るものではありません。思ったとおりにお答えください。');
  it = f.addMultipleChoiceItem().setTitle('数-1 人数を答えるとき、一人ずつ数えていましたか').setChoiceValues(['毎回数えた', 'だいたい数えた', '目分量が多かった', 'ほとんど目分量']);
  it.setRequired(true);
  it = f.addCheckboxItem().setTitle('数-2 数えにくかったのはどんなときですか（いくつでも）').setChoiceValues(['混んでいた', '自分が立っていた', '揺れた', '降りる人が多かった', '暗かった', 'その他']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('数-3 数えきれなかったとき、どうしましたか').setChoiceValues(['だいたいの数を入れて送った', '空欄のまま送った', '数えきれないことはなかった', '覚えていない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('数-4 座っていましたか、立っていましたか').setChoiceValues(['ほぼ座っていた', '半々', 'ほぼ立っていた']);
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('ゆき（数字キー）について');
  it.setHelpText('数字のキーを押して人数を直接入力する画面です。');
  it = f.addScaleItem().setTitle('ゆ-1 全体として答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('ゆ-2 素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-3 押し間違えにくかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-4 思ったとおりの人数を入れられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-5 頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('かえり（＋−ボタン）について');
  it.setHelpText('「−5」「−1」「＋1」「＋5」のボタンで数を合わせる画面です。');
  it = f.addScaleItem().setTitle('か-1 全体として答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('か-2 素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-3 押し間違えにくかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-4 思ったとおりの人数を入れられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-5 頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('ゆきとかえりをくらべて');
  it = f.addMultipleChoiceItem().setTitle('比-1 どちらが答えやすかったですか').setChoiceValues(['ゆきの数字キー', 'かえりの＋−ボタン', '変わらない']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('比-2 そう思った理由があれば教えてください');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('最後に');
  it = f.addParagraphTextItem().setTitle('終-1 人数を答えるとき、いちばん面倒だったことは何ですか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('終-2 こうしたら答えやすくなる、という案があれば');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function make2() {
  var f = FormApp.create('バスの人数アプリ ― 2回目');
  f.setDescription('ゆきとかえりの2本に乗ったあとにお答えください。8分ほどで終わります。');
  try { f.setCollectEmail(false); } catch (e) {}
  var it;
  it = f.addSectionHeaderItem().setTitle('2回目のアンケート');
  it.setHelpText('ゆきとかえりの2本に乗ったあとにお答えください。8分ほどで終わります。');
  it = f.addTextItem().setTitle('初-1 カードに書いてあるID');
  it.setHelpText('1回目と同じIDを入れてください');
  it.setRequired(true);
  it = f.addDateItem().setTitle('初-2 今日の日付');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('初-3 今日乗ったのはどちらですか').setChoiceValues(['昼', '夕方']);
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('今日、人数を数えたときのこと');
  it.setHelpText('正解・不正解を見るものではありません。思ったとおりにお答えください。');
  it = f.addMultipleChoiceItem().setTitle('数-1 人数を答えるとき、一人ずつ数えていましたか').setChoiceValues(['毎回数えた', 'だいたい数えた', '目分量が多かった', 'ほとんど目分量']);
  it.setRequired(true);
  it = f.addCheckboxItem().setTitle('数-2 数えにくかったのはどんなときですか（いくつでも）').setChoiceValues(['混んでいた', '自分が立っていた', '揺れた', '降りる人が多かった', '暗かった', 'その他']);
  it.setRequired(false);
  it = f.addMultipleChoiceItem().setTitle('数-3 数えきれなかったとき、どうしましたか').setChoiceValues(['だいたいの数を入れて送った', '空欄のまま送った', '数えきれないことはなかった', '覚えていない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('数-4 座っていましたか、立っていましたか').setChoiceValues(['ほぼ座っていた', '半々', 'ほぼ立っていた']);
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('ゆき（＋−ボタン）について');
  it.setHelpText('「−5」「−1」「＋1」「＋5」のボタンで数を合わせる画面です。');
  it = f.addScaleItem().setTitle('ゆ-1 全体として答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('ゆ-2 素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-3 押し間違えにくかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-4 思ったとおりの人数を入れられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('ゆ-5 頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('かえり（数字キー）について');
  it.setHelpText('数字のキーを押して人数を直接入力する画面です。');
  it = f.addScaleItem().setTitle('か-1 全体として答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('か-2 素早く答えられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-3 押し間違えにくかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-4 思ったとおりの人数を入れられた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('か-5 頭を使う・疲れると感じた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('ゆきとかえりをくらべて');
  it = f.addMultipleChoiceItem().setTitle('比-1 どちらが答えやすかったですか').setChoiceValues(['ゆきの＋−ボタン', 'かえりの数字キー', '変わらない']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('比-2 そう思った理由があれば教えてください');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('「わからない」ボタンについて');
  it.setHelpText('今日は、人数を入れる欄のすぐ下に「わからない」というボタンがありました。');
  it = f.addMultipleChoiceItem().setTitle('分-1 「わからない」ボタンを押したことがありますか').setChoiceValues(['何度もある', '数回ある', '一度もない', '気づかなかった']);
  it.setRequired(true);
  it = f.addScaleItem().setTitle('分-2 「わからない」ボタンがあったほうが答えやすかった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('分-3 ボタンが無かった1回目とくらべて、正直に答えられましたか').setChoiceValues(['今日のほうが正直に答えられた', '変わらない', '1回目のほうが正直に答えられた']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('分-4 そう思った理由があれば教えてください');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('アプリ全体について');
  it.setHelpText('2回使ってみた全体の印象をお答えください。');
  it = f.addScaleItem().setTitle('全-1 このアプリを何度も使いたいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-2 このアプリは必要以上に複雑だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-3 このアプリは簡単に使えると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-4 このアプリを使うには、詳しい人の助けが必要だと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-5 このアプリのいろいろな機能は、うまくまとまっていると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-6 このアプリには一貫性のないところが多いと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-7 たいていの人は、このアプリの使い方をすぐ覚えられると思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-8 このアプリはとても使いにくいと思う').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-9 このアプリを自信をもって使えた').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addScaleItem().setTitle('全-10 このアプリを使う前に、いろいろ覚える必要があった').setBounds(1, 5).setLabels('まったくそう思わない', 'とてもそう思う');
  it.setRequired(false);
  it = f.addPageBreakItem().setTitle('最後に');
  it = f.addParagraphTextItem().setTitle('終-1 人数を答えるとき、いちばん面倒だったことは何ですか');
  it.setRequired(false);
  it = f.addParagraphTextItem().setTitle('終-2 こうしたら答えやすくなる、という案があれば');
  it.setRequired(false);
  return [f.getPublishedUrl(), f.getEditUrl()];
}
