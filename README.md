# WPF-Viewer-Score-Manager
WPF 기반 시청자 점수 관리 프로그램입니다. DataGrid를 활용하여 점수 입력, 자동 계산, JSON 저장/불러오기를 지원합니다.

- DataGrid 기반 점수 관리
- 총점 자동 계산 (DataColumn Expression)
- JSON 저장 / 불러오기
- 자동 저장 (App 종료 시)
- 이미지 표시 기능 (UPI / 뿌요)
- 동적 컬럼 추가 / 삭제 / 이름 변경

  ### 환경
- .NET 6+
- Windows (WPF)

### 실행
1. 프로젝트 클론
2. Visual Studio에서 열기
3. 실행 (F5)

### Screenshot
<img width="1471" height="765" alt="스크린샷(708)" src="https://github.com/user-attachments/assets/691123b1-704f-44c7-adbe-084eb4f6e913" />

## 🛠 Tech Stack

- C# / WPF
- DataTable
- JSON (System.Text.Json)

## 📚 What I Learned

- DataTable을 활용한 동적 테이블 구조 설계
- WPF DataGrid 커스터마이징
- JSON 기반 상태 저장 및 복원
- INotifyPropertyChanged를 통한 UI 바인딩

  ## 🔧 Future Improvements

- MVVM 구조로 리팩토링
- SQServer 연동
- 이미지 Base64 저장으로 포터블화

  ## 🤖 About AI Usage

This project was initially generated with the help of OpenAI Codex,
and then manually reviewed, modified, and extended.

I focused on understanding the code structure and improving:
- Data handling logic
- UI behavior
- Error handling
