/*
웹UI에서 클릭 시 호출 예시:
ppttpl://insert?src=<구글URL>&template=<id>&pos=end

pos 옵션:
- end (기본): 맨 끝에 삽입
- afterCurrent: 현재 선택 슬라이드 뒤에 삽입
- at:<n> : n번 슬라이드 뒤에 삽입 (예: at:3)
*/
function openInPowerPoint(src, id, pos="end"){
  const u = "ppttpl://insert?src=" + encodeURIComponent(src)
          + "&template=" + encodeURIComponent(id || "")
          + "&pos=" + encodeURIComponent(pos || "end");
  window.location.href = u;
}
