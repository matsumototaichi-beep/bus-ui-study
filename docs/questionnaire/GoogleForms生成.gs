// 事後アンケートの Google フォームを作るスクリプト
// docs/questionnaire/設問一覧.xlsx から自動生成。Excel を直して作り直すこと。
//
// 1. Google ドライブ →「新規 → その他 → Google Apps Script」
// 2. この中身をぜんぶ貼り付けて保存
// 3. 関数 createAll を実行
// 4. 実行ログに2つのURLが出る
// 5. 回答用URLを Google にログインしていないブラウザで開き、答えられるか確かめる
//
// ★先に whoami を実行して、どのアカウントで動くか確かめること。
//   フォームは、ここに出たアカウントの Google ドライブに作られる。
//
// ★2026-10-07 に設問を減らした（1回目8問・2回目11問）。前に作った2つのフォームはゴミ箱に入れること。

function whoami() {
  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());
}

function createAll() {
  var a = make1(), b = make2();
  Logger.log('1回目\n  回答用: ' + a[0] + '\n  編集用: ' + a[1]);
  Logger.log('2回目\n  回答用: ' + b[0] + '\n  編集用: ' + b[1]);
}

function make1() {
  var f = FormApp.create('バスの人数アプリ　1回目のアンケート');
  f.setDescription('往復が終わったら答えてください（2分ほど）。思ったとおりに選んでください。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('今回の往復について');
  it = f.addScaleItem().setTitle('問2 ゆき（吹田駅 → 桃山台駅）のバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 かえり（桃山台駅 → 吹田駅）のバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('人数の入れ方について');
  it = f.addScaleItem().setTitle('問6 数字キーの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('0〜9 のキーで人数を打つ画面');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問7 ＋−ボタンの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('−5・−1・＋1・＋5 で人数を合わせる画面');
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問8 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
  return [f.getPublishedUrl(), f.getEditUrl()];
}

function make2() {
  var f = FormApp.create('バスの人数アプリ　2回目のアンケート');
  f.setDescription('往復が終わったら答えてください（3分ほど）。思ったとおりに選んでください。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;
  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('今回の往復について');
  it = f.addScaleItem().setTitle('問2 ゆき（吹田駅 → 桃山台駅）のバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問3 かえり（桃山台駅 → 吹田駅）のバスの混み具合').setBounds(1, 5).setLabels('空いていた', '混んでいた');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問4 人数を数えきれないことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問5 数えきれないまま、だいたいの人数を入れて送ったことがありましたか').setBounds(1, 5).setLabels('なかった', '何度もあった');
  it.setRequired(true);
  it = f.addSectionHeaderItem().setTitle('人数の入れ方について');
  it = f.addScaleItem().setTitle('問6 数字キーの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('0〜9 のキーで人数を打つ画面');
  it.setRequired(true);
  it = f.addScaleItem().setTitle('問7 ＋−ボタンの答えやすさ').setBounds(1, 5).setLabels('答えにくかった', '答えやすかった');
  it.setHelpText('−5・−1・＋1・＋5 で人数を合わせる画面');
  it.setRequired(true);
  it = f.addPageBreakItem().setTitle('「わからない」ボタンについて');
  it.setHelpText('1回目と2回目のどちらか一方だけ、人数を入れるところの下に「わからない」ボタンがありました。');
  it = f.addMultipleChoiceItem().setTitle('問8 「わからない」ボタンがあったのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '覚えていない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('問9 人数を答えやすかったのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '変わらない']);
  it.setRequired(true);
  it = f.addMultipleChoiceItem().setTitle('問10 数えたとおりの人数を答えられたのは、どちらの回ですか').setChoiceValues(['1回目', '2回目', '変わらない']);
  it.setRequired(true);
  it = f.addParagraphTextItem().setTitle('問11 答えにくかったことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);
  return [f.getPublishedUrl(), f.getEditUrl()];
}
