# 🚀 앱 배포 가이드 (Streamlit Cloud)

이 가이드는 현재 작성된 코드를 **Github**에 올리고, **Streamlit Cloud**를 통해 웹사이트로 배포하는 과정을 설명합니다.

---

### **1단계: Github 저장소(Repository) 만들기**

1.  [Github 웹사이트](https://github.com/)에 로그인합니다.
2.  우측 상단의 **'+'** 아이콘을 누르고 **'New repository'**를 클릭합니다.
3.  **Repository name**에 `sat-analysis-app` (또는 원하는 이름)을 입력합니다.
4.  **Public** (공개) 또는 **Private** (비공개)를 선택합니다. (무료 계정은 Public 추천)
5.  다른 설정(Add a README 등)은 건드리지 말고 맨 아래 **'Create repository'** 버튼을 클릭합니다.
6.  생성 완료 후 나오는 화면에서 **HTTPS 주소** (예: `https://github.com/사용자명/sat-analysis-app.git`)를 복사해둡니다.

---

### **2단계: 코드 업로드 (터미널 명령어)**

현재 폴더에는 이미 Git 설정이 완료되어 있습니다. 아래 명령어만 순서대로 입력하세요.

1.  **원격 저장소 연결** (위에서 복사한 주소를 붙여넣으세요)
    ```powershell
    git remote add origin https://github.com/여러분의_아이디/sat-analysis-app.git
    ```
    *(만약 `error: remote origin already exists` 오류가 나면 `git remote set-url origin 주소` 를 입력하세요)*

2.  **코드 올리기 (Push)**
    ```powershell
    git branch -M main
    git push -u origin main
    ```
    *(이 과정에서 Github 로그인 창이 뜨면 로그인해주세요)*

---

### **3단계: Streamlit Cloud 배포**

1.  [Streamlit Cloud](https://share.streamlit.io/)에 접속하여 로그인합니다 (Github 계정으로 로그인).
2.  **'New app'** 버튼을 클릭합니다.
3.  **'Use existing repo'**를 선택하고, 방금 만든 `sat-analysis-app` 저장소를 선택합니다.
4.  설정은 다음과 같이 확인합니다:
    *   **Main file path**: `app.py`
5.  **'Deploy!'** 버튼을 클릭합니다.

---

### **⚠️ 4단계: API Key 설정 (가장 중요!)**

앱이 실행되려면 Google API Key가 필요합니다. 코드가 아닌 **Streamlit Secrets**에 안전하게 저장해야 합니다.

1.  배포 중인 앱 화면 우측 하단의 **'Manage app'** 버튼을 누릅니다. (또는 대시보드에서 앱 옆의 **'...'** -> **'Settings'**)
2.  **'Secrets'** 탭을 클릭합니다.
3.  아래 내용을 복사해서 붙여넣습니다:
    ```toml
    GOOGLE_API_KEY = "여기에_여러분의_실제_API_KEY를_넣으세요"
    ```
4.  **'Save'**를 누릅니다.

이제 앱이 자동으로 재시작되며 정상적으로 작동할 것입니다! 🎉
