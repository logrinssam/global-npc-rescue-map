/**
 * 지구촌 NPC 구조대 — 구글폼 자동 생성 스크립트
 *
 * 사용법
 *  1. 크롬에서 script.google.com → "새 프로젝트" → 이 파일 내용을 전부 붙여 넣고 저장.
 *  2. 위 함수 선택 상자에서 makeForm 을 고르고 ▶ 실행. 처음 한 번 권한 허용.
 *  3. 실행 로그(Ctrl+Enter)에 폼 편집 주소·학생용 주소·응답 시트 주소가 찍힌다.
 *
 * 별칭은 학생이 직접 쓴다(모둠은 1~5모둠 드롭다운, 별칭은 글자 입력). 교사가 미리 넣을 것은 없다.
 * 문항 글자는 사후_회고검사지.docx 와 똑같다(같은 목록에서 생성).
 * 응답 시트의 열 순서는 사후_분석.xlsx 의 "폼붙여넣기" 시트가 기대하는 순서이므로
 * 폼에서 문항을 추가·삭제·이동하지 말 것. (문구만 고치는 것은 괜찮음)
 */

var TEAMS = ["1모둠", "2모둠", "3모둠", "4모둠", "5모둠"]; // NPC 여권의 모둠 선택(1~5모둠)과 같음

var A_ITEMS = [{"q": "테아가 사는 투발루에서 바닷물이 점점 높아지는 가장 큰 까닭은 무엇일까요?", "opts": ["① 지구가 더워져서", "② 비가 많이 와서", "③ 땅이 아래로 꺼져서", "④ 배가 많이 다녀서"]}, {"q": "바닷물이 밭과 우물에 들어오면 생기는 일로 알맞지 않은 것은 무엇일까요?", "opts": ["① 마실 물이 부족해진다", "② 농사가 잘 안 된다", "③ 먹을 것이 부족해진다", "④ 눈이 많이 온다"]}, {"q": "말리의 세이두가 겪는 가뭄은 비가 적게 와서 농사가 안 되고, 식량과 돈이 줄어드는 문제예요.", "opts": ["O", "X"]}, {"q": "가뭄이 계속될 때 '오래 이어지는 도움'에 가장 가까운 것은 무엇일까요?", "opts": ["① 당장 먹을 식량 한 번 보내기", "② 물 모으는 시설과 가뭄에 강한 작물", "③ 큰 잔치 열어 주기", "④ 새 옷 보내 주기"]}, {"q": "미얀마의 민처럼 분쟁(다툼)이 있는 곳의 어린이는 학교가 자주 멈춰서 배울 권리를 지키기 어려워요.", "opts": ["O", "X"]}, {"q": "분쟁 지역의 어린이에게 가장 먼저 필요한 것은 무엇일까요?", "opts": ["① 안전하게 보호받고 배우는 것", "② 새 스마트폰", "③ 더 많은 시험", "④ 혼자서 멀리 떠나기"]}, {"q": "방글라데시의 리마와 마을 사람들이 겪는 문제는 무엇일까요?", "opts": ["① 눈사태로 길이 막힘", "② 홍수로 집과 길이 물에 잠김", "③ 화산이 터짐", "④ 사막이 넓어짐"]}, {"q": "홍수가 자주 나면 살던 곳에 남을지, 떠날지 고민하는 가족이 생겨요.", "opts": ["O", "X"]}, {"q": "여러 나라가 함께 모여 지구촌 문제를 의논하고 돕는 조직(예: 국제연합 UN, 유니세프)을 무엇이라고 할까요?", "opts": ["① 국제기구", "② 마을 회의", "③ 학급 회의", "④ 회사"]}, {"q": "지구촌 문제는 그 나라 사람들끼리만 해결하면 되고, 다른 나라나 우리는 상관이 없어요.", "opts": ["O", "X"]}];
var B_ROWS = ["B1. 나는 멀리 있는 나라 사람들이 겪는 어려움이 나와도 관계있다고 느낀다.", "B2. 나는 지구촌 문제를 겪는 사람의 마음(걱정, 두려움, 희망 등)을 떠올릴 수 있다.", "B3. 나는 다른 나라의 어려운 소식을 보면 그 사람들이 어떻게 지내는지 궁금하다.", "B4. 나는 어려움을 겪는 사람도 자기 삶을 열심히 꾸려 가는 사람이라고 생각한다.", "B5. 나는 '내가 그 사람이라면 어떻게 할까?'를 생각해 볼 수 있다.", "B6. 나는 같은 문제라도 사람마다 바라는 해결 방법이 다를 수 있다고 생각한다.", "B7. 나는 지구촌 문제를 겪는 사람에게 무엇이 가장 필요한지 그 사람의 입장에서 말할 수 있다.", "B8. 나는 어떤 선택에도 좋은 점과 어려운 점이 함께 있다는 것을 안다.", "B9. 나는 친구가 나와 다른 해결 방법을 골라도 그 까닭을 들어 보고 싶다."];
var C_ROWS = ["C1. 나는 지구촌 문제를 해결하는 데 나도 할 수 있는 일이 있다고 생각한다.", "C2. 나는 지구촌 문제와 해결 방법을 친구들에게 설명할 수 있다.", "C3. 나는 지구촌 문제에 대해 더 알아보고 싶다.", "C4. 나는 지구촌 문제를 돕는 활동(알리기, 캠페인, 기부 등)에 참여하고 싶다.", "C5. 나는 지구촌 문제를 위해 우리 모둠이나 학급이 할 수 있는 일을 말할 수 있다.", "C6. 나는 지구촌 문제를 도울 때 그곳 주민의 의견을 먼저 들어 보려고 할 것이다."];
var D_ITEMS = ["① '지구촌 문제' 하면 떠오르는 것 3가지를 쓰고, 왜 떠올랐는지 까닭도 써 보세요.", "② \"내가 지구촌 문제를 겪는 친구라면, ______\"  빈칸을 이어서 써 보세요. (그 친구의 마음, 필요한 것, 바라는 것을 그 친구가 되어 써요.)", "③ 지구촌 문제는 누가, 어떻게 해결해야 할까요? (함께할 사람과 각자 할 일을 써 보세요.)", "④ 성찰문: \"수업 전에는 지구촌 문제를 ______라고 생각했는데, 지금은 ______\"  두 줄로 써 보세요."];
var SCALE = ["① 전혀 아니다", "② 아니다", "③ 보통이다", "④ 그렇다", "⑤ 매우 그렇다"];

function makeForm() {
  var r1 = makeRetroForm();
  Logger.log("\n===== 완료 =====\n[회고검사지 입력용 폼]\n 편집: " + r1.edit + "\n 학생용: " + r1.live + "\n 응답 시트: " + r1.sheet);
}

function addA(form) {
  form.addSectionHeaderItem().setTitle("A. 지구촌 문제 알기")
      .setHelpText("종이에 ○ 한 답을 그대로 고르세요. 종이에 안 쓴 문항은 비워 두세요.");
  A_ITEMS.forEach(function (it, i) {
    form.addMultipleChoiceItem().setTitle((i + 1) + ". " + it.q).setChoiceValues(it.opts).setRequired(false);
  });
}

function addGrid(form, title, rows) {
  form.addGridItem().setTitle(title).setRows(rows).setColumns(SCALE).setRequired(false)
      .setHelpText("종이에 ○ 한 번호를 그대로 고르세요. 종이에 안 쓴 문항은 비워 두세요.");
}

function addD(form, count) {
  form.addSectionHeaderItem().setTitle("D. 생각 쓰기").setHelpText("종이에 쓴 글을 그대로 옮겨 쓰세요. 맞춤법은 고치지 않아도 돼요.");
  for (var i = 0; i < count; i++) {
    form.addParagraphTextItem().setTitle(D_ITEMS[i]).setRequired(false);
  }
}

function linkSheet(form, title) {
  var ss = SpreadsheetApp.create(title + " (응답)");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  return ss.getUrl();
}

function makeRetroForm() {
  var form = FormApp.create("지구촌 NPC 구조대 — 수업을 마치고 나를 돌아보는 검사지 (입력용)");
  form.setDescription("종이 검사지에 쓴 답을 그대로 옮겨 적는 곳이에요. 종이에 쓴 것과 똑같이 입력하고, 종이에 안 쓴 문항은 비워 두세요.")
      .setCollectEmail(false).setLimitOneResponsePerUser(false).setProgressBar(true).setAllowResponseEdits(false);

  form.addListItem().setTitle("모둠").setChoiceValues(TEAMS).setRequired(true);
  form.addTextItem().setTitle("별칭").setHelpText("종이 검사지와 챗봇에 쓴 별칭을 똑같이 쓰세요. 모둠 번호는 빼고 별칭만 (예: 바다)").setRequired(true);

  form.addPageBreakItem().setTitle("A. 지구촌 문제 알기");
  addA(form);

  form.addPageBreakItem().setTitle("B. 마음 알기");
  form.addSectionHeaderItem().setTitle("B. 마음 알기").setHelpText("종이의 왼쪽 칸(수업 전의 나)과 오른쪽 칸(지금의 나)을 각각 옮겨요.");
  addGrid(form, "B — 수업 전의 나", B_ROWS);
  addGrid(form, "B — 지금의 나", B_ROWS);

  form.addPageBreakItem().setTitle("C. 할 수 있는 일");
  form.addSectionHeaderItem().setTitle("C. 할 수 있는 일").setHelpText("종이의 왼쪽 칸(수업 전의 나)과 오른쪽 칸(지금의 나)을 각각 옮겨요.");
  addGrid(form, "C — 수업 전의 나", C_ROWS);
  addGrid(form, "C — 지금의 나", C_ROWS);

  form.addPageBreakItem().setTitle("D. 생각 쓰기");
  addD(form, 4);

  form.setConfirmationMessage("수고했어요! 종이 검사지는 선생님께 내 주세요.");
  var sheet = linkSheet(form, "회고검사지");
  return { edit: form.getEditUrl(), live: form.getPublishedUrl(), sheet: sheet };
}
