/**
 * ご入会のしおり 確認フォームを作成する Google Apps Script
 *
 * 使い方：
 *   1. https://script.google.com で新しいプロジェクトを作り、このファイルの中身を貼り付ける
 *   2. 関数「createGuideForm」を選んで実行（初回は権限の許可を求められる）
 *   3. ログに表示されるフォームの URL・回答用スプレッドシートの URL を確認する
 *
 * しおりの内容を変えたときは GUIDE_VERSION と SECTIONS を直して再実行し、新しいフォームを作る。
 * （どの版のしおりに同意したかを、フォーム名と回答シートで区別するため）
 */

// しおりの版。元にしたファイル：入会のしおり（2026.10版・Word）
const GUIDE_VERSION = '2026.10';

// 章ごとのチェック項目。checks は全部チェックしないと次のページへ進めない。
const SECTIONS = [
  {
    title: '必要書類・初回契約金',
    description: '月額会費はご利用開始日やお支払い方法（クレジットカード可）によって異なります。',
    checks: [
      '入会金 22,000円（税込／事務手数料・カード発行手数料込）',
      '月額会費 1ヵ月分〜3ヵ月分',
      '〈A〉銀行口座情報（通帳またはキャッシュカード）と銀行印、または〈B〉クレジットカード',
      '身分証明書',
    ],
  },
  {
    title: 'ご入会の流れ',
    description:
      '入会手続き時に口座情報・銀行口座印をお持ちでない方は、次回ご来店時にお持ちいただき、記入・捺印をお願いします。' +
      '場合によっては、口座引落の開始が遅くなる場合がございます。',
    checks: [
      'ご入会コース・オプションの確認／費用のご説明',
      '契約書類の記載',
      '身分証明書のコピー',
      '各種アプリのダウンロード（Legame 公式アプリ／公式 LINE アカウント／れんらくアプリ）',
      '写真撮影',
      '入会のしおり、全体の解説',
      'お会計',
      '会員証のお渡し、ご入館の流れ',
    ],
  },
  {
    title: '重要な会員規約',
    description: '後々の行き違いを防ぐため、特に大切な規約です。1つずつご説明します。',
    checks: [
      '第13条（コース変更・休会・退会手続き）：会員契約書頭書の「変更受付開始日」以降に申し出が可能。' +
        '退会は手続きした月の翌月末日、コース変更・休会は手続きした月の翌々月から適用。' +
        '休会は原則2ヵ月間で、その後は元のコースに自動復帰。手続きは店頭の所定書式のみで、電話や LINE では受け付けない。',
      '第18条（営業日および営業時間）：イベント等で営業日時が変わる場合は公式アプリで事前にお知らせするため、通知設定をする。',
      '第23条（盗難および紛失）・第24条（怪我、事故等を回避するための諸注意）：' +
        'パーソナルトレーニング中に身体に違和感があれば直ちにトレーナーに申し出る。' +
        '後日の不調は、指導が要因である証明ができない限り責任を負いかねる。第23条・第24条を家でも改めて確認する。',
      '第31条（施設レンタル品の紛失・破損等による賠償）：故意・過失による紛失・破損は相当の弁済が発生。' +
        '有料ロッカーの鍵の再作成費 3,300円（税込）。',
      '第32条（予約のキャンセル、遅刻）：予約日の前営業日の営業終了以降のキャンセルは原則できず、権利を消化したものとする' +
        '（二親等以内の方のご不幸を除く）。事前連絡がなく予約時間に遅れた場合は、いかなる事情があろうとも権利を消化したものとする。遅刻した場合も、原則として予約の時間枠内でのトレーニングとなる。',
      '第33条（ピラティスコースの終了および代替措置について）：外部講師の退任等でコースを終了する場合は、終了日の2ヶ月前までに通知し、' +
        '在籍会員は終了日以降に自動的に A コースへ移行する。所定の契約期間（第13条2項および第34条）中は、コース終了を理由に退会できない。',
    ],
  },
  {
    title: 'その他の大事なポイント',
    description: '',
    checks: [
      'カフェバースペースへの飲食の持ち込み不可（2階トレーニングスペースでの飲み物は可、食べ物は不可）',
      'パーソナルトレーニングの予約可能回数はコースごとに異なり、上限を超える予約はできない。毎月16日に翌月分の予約カレンダーが公開される',
      '月額会費は前納制。口座振替は毎月27日（土日祝は翌平日）に翌月分を引落し（ネット銀行では「アプラス」名義で承認要請が届く）。クレジットカードは毎月末日に翌月分を決済',
      '年会費 11,000円（税込）：口座振替は初年度分を月額会費の初回振替と一緒に、クレジットカードは入会月の末日に決済。次年度以降は入会月の月額会費と一緒',
      '体組成計は無料で利用でき、データ保存用のカードは500円で販売している',
    ],
  },
];

const STAFF_LIST = ['（スタッフ名を設定してください）'];

function createGuideForm() {
  const form = FormApp.create(`ご入会のしおり 確認（${GUIDE_VERSION}版）`);
  form
    .setDescription(
      'スタッフの説明を聞きながら、各項目を確認してチェックを入れてください。\n' +
        'すべての項目にチェックを入れると、次のページへ進めます。'
    )
    .setProgressBar(true)
    .setShowLinkToRespondAgain(false)
    .setAllowResponseEdits(false)
    .setConfirmationMessage('ご確認ありがとうございました。スタッフにお声がけください。');

  // 1ページ目：基本情報
  form.addTextItem().setTitle('お名前（フルネーム）').setRequired(true);
  form.addTextItem().setTitle('会員番号').setHelpText('分かる場合のみ').setRequired(false);
  form.addListItem().setTitle('担当スタッフ').setChoiceValues(STAFF_LIST).setRequired(true);

  // 章ごとに1ページ。全項目チェック必須
  SECTIONS.forEach((section) => {
    form.addPageBreakItem().setTitle(section.title).setHelpText(section.description);
    const validation = FormApp.createCheckboxValidation()
      .requireSelectExactly(section.checks.length)
      .setHelpText('すべての項目にチェックを入れてください。')
      .build();
    form
      .addCheckboxItem()
      .setTitle(`「${section.title}」の説明を受け、確認しました`)
      .setChoiceValues(section.checks)
      .setValidation(validation)
      .setRequired(true);
  });

  // 最終ページ：同意と署名（氏名の入力）
  form.addPageBreakItem().setTitle('最終確認');
  form
    .addCheckboxItem()
    .setTitle('同意事項')
    .setChoiceValues([`ご入会のしおり（${GUIDE_VERSION}版）の内容について説明を受け、理解しました`])
    .setRequired(true);
  form.addTextItem().setTitle('ご署名（お名前をもう一度入力してください）').setRequired(true);

  // 回答の保存先スプレッドシート（回答のたびに1行追加される。日時は自動で記録される）
  const sheet = SpreadsheetApp.create(`ご入会のしおり 確認記録（${GUIDE_VERSION}版）`);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  Logger.log('フォーム（回答用）: ' + form.getPublishedUrl());
  Logger.log('フォーム（編集用）: ' + form.getEditUrl());
  Logger.log('確認記録シート: ' + sheet.getUrl());
}
