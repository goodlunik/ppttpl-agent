# 설치 프로그램(EXE) 최종본을 "SDK 설치 없이" 받는 방법

이 프로젝트는 **GitHub Actions**가 Windows에서 자동으로 빌드/패키징해서
`PptTplAgent_Setup.exe`를 만들어 줍니다.  
즉, 로컬 PC에 .NET SDK를 설치할 필요가 없습니다.

## 1) GitHub에 새 저장소 만들기
- GitHub에서 새 repo 생성
- 이 ZIP의 내용을 그대로 업로드(커밋/푸시)

## 2) Actions에서 빌드 실행
- repo → **Actions** → "Build Windows Installer"
- "Run workflow" 실행 (또는 main 브랜치에 push하면 자동 실행)

## 3) 결과물 다운로드
- Actions 실행 결과 → Artifacts → **PptTplAgent_Setup**
- `PptTplAgent_Setup.exe` 다운로드

## 4) 사용자 PC에서 설치
- `PptTplAgent_Setup.exe` 실행
- `{localappdata}\PptTplAgent`에 설치 (관리자 권한 불필요)
- `ppttpl://` 프로토콜이 HKCU에 등록되어 웹UI 클릭이 곧바로 동작

## 필수 조건
- Windows + 데스크톱 PowerPoint 설치(웹버전/모바일은 불가)
- 구글 파일은 “링크가 있는 사용자 누구나 보기” 권한 권장
