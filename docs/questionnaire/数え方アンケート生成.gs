// 「数え方」のアンケートの Google フォームを作るスクリプト（UI研究の追加調査）
//
// なぜ作るか：
//   参加者から「始発から乗るので最初は人が少なくて数えやすく、あとは乗ってきた人と
//   降りた人を足し引きすればいい」という話が出た（2026-10-17 松本が聞き取り）。
//   本当に足し引きで答えているなら、人数が増えても所要時間が伸びない（相関0.015）ことや、
//   30人以上でも人どうしの答えが中央1人しか違わない（ICC 0.95）ことの説明になる。
//   計数研究（1から数える vs 乗り降りを足し引き）の読み方も変わる。
//
//   既存の4フォームには足さない。足すと結合の道具の列対応が変わるうえ、
//   1回目のアンケートで数え方を聞くと2回目の数え方に影響してしまうため。
//   独立した短いフォームにして、**往復がすべて終わった人**にだけ答えてもらう。
//   すでに終わった6人（A・D・E・F・G・H）には、あとからチャットで依頼する。
//
// 1. Google ドライブ →「新規 → その他 → Google Apps Script」
// 2. この中身をぜんぶ貼り付けて保存
// 3. 関数 createCounting を実行
// 4. 実行ログに2つのURLが出る
// 5. 回答用URLを Google にログインしていないブラウザで開き、答えられるか確かめる
//
// ★先に whoami を実行して、どのアカウントで動くか確かめること。
//   フォームは、ここに出たアカウントの Google ドライブに作られる。

function whoami() {
  Logger.log('いま動いているアカウント: ' + Session.getEffectiveUser().getEmail());
}

function createCounting() {
  var f = FormApp.create('バスの人数アプリ　数え方のアンケート');
  f.setDescription('2回の往復がすべて終わった方にうかがいます（1分ほど）。正解はありません。実際にやっていたとおりにお答えください。');
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setRequireLogin(false); } catch (e) {}
  try { f.setPublished(true); } catch (e) {}
  var it;

  it = f.addTextItem().setTitle('問1 カードに書いてあるID（例：A）');
  it.setValidation(FormApp.createTextValidation().setHelpText('カードのIDを1文字で入れてください（例：A）').requireTextMatchesPattern('^\\s*[A-Xa-xＡ-Ｘａ-ｘ]\\s*$').build());
  it.setRequired(true);

  it = f.addMultipleChoiceItem().setTitle('問2 車内の人数を答えるとき、どうやって数えていましたか')
    .setChoiceValues([
      '毎回、車内の人を1から数え直していた',
      '前に答えた人数に、乗ってきた人と降りた人を足し引きしていた',
      'その時によって変えていた'
    ]);
  it.setRequired(true);

  it = f.addMultipleChoiceItem().setTitle('問3 そのやり方は、4回の区間（1回目のゆき・かえり、2回目のゆき・かえり）でずっと同じでしたか')
    .setChoiceValues(['ずっと同じだった', '途中で変わった', '覚えていない']);
  it.setRequired(true);

  // 問4 は「変化」を直接聞く。4回の区間のどれも始発から出発するので、
  // 「出発直後は」と聞くと「4回のうちどれのことか」が決まらない（松本の指摘 2026-10-17）。
  // 知りたいのは「最初とあとの差」なので、差そのものを 1問で聞く。
  // 混まなかった便に当たった人も「変わらなかった」で答えられる
  it = f.addScaleItem().setTitle('問4 車内に人が少ないときに比べて、人が増えてくると、車内の人数を数えるのは難しくなりましたか')
    .setBounds(1, 5).setLabels('ぜんぜん変わらなかった', 'とても難しくなった');
  it.setHelpText('4回の区間全体を通して、だいたいのところでお答えください');
  it.setRequired(true);

  it = f.addScaleItem().setTitle('問5 車内の人数を数えるのは、全体として大変でしたか')
    .setBounds(1, 5).setLabels('楽だった', '大変だった');
  it.setRequired(true);

  it = f.addParagraphTextItem().setTitle('問6 数え方について、気づいたことや工夫したことがあれば書いてください（なければ「なし」）');
  it.setRequired(true);

  Logger.log('数え方のアンケート\n  回答用: ' + f.getPublishedUrl() + '\n  編集用: ' + f.getEditUrl());
  return [f.getPublishedUrl(), f.getEditUrl()];
}
