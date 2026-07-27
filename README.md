# 기업서류 PDF 자동 수집기

기업명, 참여자명, 멘토명, 수집 차수와 월 구간을 기준으로 필요한 PDF를 찾아 참여자별 폴더에 복사하는 도구입니다. 원본 파일은 이동하거나 수정하지 않습니다.

## 주요 기능

- 선택한 폴더와 모든 하위 폴더에서 PDF 검색
- 기업명과 문서별 핵심 키워드로 파일 판별
- 참여자별 결과 폴더 생성
- 같은 조건의 파일이 여러 개면 최근 수정 파일 선택
- 2차 수집 시 출석부, 면담일지, 멘토수당 신청서는 파일명에 `2차`가 있는 파일만 선택
- `개인정보 동의서`와 `개인정보 활용에 관한 동의서` 모두 인정

## Windows에서 실행

PowerShell 버전은 `실행.bat`을 더블클릭합니다.

Python 버전:

```powershell
python app.py
```

Windows 실행 파일 생성:

```powershell
python -m pip install pyinstaller
python -m PyInstaller --clean --noconfirm --onefile --windowed --name "기업서류 PDF 수집기" app.py
```

완성된 파일은 `dist/기업서류 PDF 수집기.exe`입니다.

## macOS에서 실행

```bash
python3 app.py
```

macOS 앱 생성:

```bash
python3 -m pip install pyinstaller
python3 -m PyInstaller --clean --noconfirm --windowed --name "기업서류 PDF 수집기" app.py
```

완성된 앱은 `dist/기업서류 PDF 수집기.app`입니다. macOS 앱은 macOS에서 빌드해야 합니다.

## 사용 방법

1. 검색할 폴더와 결과 저장 폴더를 선택합니다.
2. 기업명, 수집 차수, 월 구간, 참여자명과 멘토명을 입력합니다.
3. `PDF 복사본 수집`을 누릅니다.
4. 완료 후 `결과 폴더 열기`로 수집 결과를 확인합니다.

참여자와 멘토가 여러 명이면 같은 순서로 쉼표로 구분합니다.

