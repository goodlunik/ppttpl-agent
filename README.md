# PptTplAgent (.NET) - AHK 대체 프로토콜 핸들러

웹UI에서 `ppttpl://insert?src=...` 를 호출하면,
Windows 프로토콜 핸들러가 실행되어 Google Slides/Drive에서 PPTX를 다운로드하고,
PowerPoint에 슬라이드를 삽입합니다.

## 요구사항
- Windows
- PowerPoint 설치
- .NET 8 SDK (빌드할 때만 필요)

## 빌드 (단일 exe 권장)
PowerShell / CMD:

```bash
cd src\PptTplAgent
dotnet publish -c Release -r win-x64 /p:PublishSingleFile=true /p:SelfContained=true
```

출력:
`src\PptTplAgent\bin\Release\net8.0-windows\win-x64\publish\PptTplAgent.exe`

## 설치(프로토콜 등록)
### 방법 1) PowerShell 스크립트(권장)
`scripts\install_protocol.ps1` 실행 → exe 경로를 입력하면 HKCU(현재 사용자)에 등록합니다.

### 방법 2) reg 파일
`scripts\register_ppttpl_CURRENT_USER.reg` 안의 `PATH_TO_EXE` 를 실제 exe 경로로 바꾼 뒤 실행

## 테스트
크롬 주소창:
`ppttpl://ping`

또는 웹UI 버튼 클릭.

## 웹UI에서 호출 예시
```js
const url = "ppttpl://insert?src=" + encodeURIComponent(src) + "&template=" + encodeURIComponent(id) + "&pos=end";
window.location.href = url;
```

## 주의: 구글 공유 권한
이 프로그램은 로그인 쿠키 없이 다운로드를 시도합니다.
따라서 템플릿 파일은 **'링크가 있는 사용자 누구나 보기'** 로 공유되어야 안정적으로 다운로드됩니다.
(권한이 막혀 있으면 HTML 로그인 페이지가 내려와 실패합니다.)
