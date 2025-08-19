# Sales Department Weekly Report System

## 📊 프로젝트 개요
영업부 주간 보고서 자동화 시스템으로, ERP 시스템에서 데이터를 수집하여 표준화된 보고서를 생성합니다.

## 🚀 주요 기능
- **자동 데이터 수집**: ERP 시스템에서 매출 및 매출채권 데이터 자동 추출
- **매출 분석**: 일별, 월별, 누적 매출 계산 및 분석
- **매출채권 분석**: 미수금 현황 및 회수 가능성 분석
- **자동 보고서 생성**: 표준화된 XML 포맷으로 보고서 자동 생성
- **GUI 인터페이스**: 사용자 친화적인 그래픽 인터페이스 제공

## 📁 프로젝트 구조
```
SalesReportDashboard/
├── applications/
│   ├── main.py              # 메인 실행 파일
│   └── gui.py               # GUI 애플리케이션
├── modules/
│   ├── core/                # 핵심 비즈니스 로직
│   │   ├── sales_calculator.py
│   │   ├── processed_receivables_analyzer.py
│   │   └── accounts_receivable_analyzer.py
│   ├── data/                # 데이터 처리 모듈
│   │   └── processors/
│   ├── gui/                 # GUI 컴포넌트
│   │   └── login_dialog.py
│   ├── reports/             # 보고서 생성
│   │   └── xml_safe_report_generator.py
│   └── utils/               # 유틸리티
│       └── config_manager.py
├── requirements.txt         # 의존성 패키지
└── README.md               # 프로젝트 문서
```

## ⚙️ 설치 및 실행

### 1. 필요 패키지 설치
```bash
pip install -r requirements.txt
```

### 2. 실행 방법
```bash
# GUI 모드로 실행
python applications/main.py

# 또는 직접 GUI 실행
python applications/gui.py
```

## 🔧 설정
- `modules/utils/config_manager.py`에서 ERP 연결 설정 및 기타 옵션 구성
- 첫 실행 시 ERP 로그인 정보 입력 필요

## 📋 시스템 요구사항
- Python 3.8+
- Windows 10/11 (ERP 시스템 호환성)
- 최소 4GB RAM 권장

## 🛠️ 개발 정보
- **언어**: Python
- **GUI 프레임워크**: tkinter
- **데이터 처리**: pandas, openpyxl
- **ERP 연동**: selenium (웹 기반 ERP)

## 📖 사용법
1. 애플리케이션 실행
2. ERP 로그인 정보 입력
3. 보고서 생성 기간 선택
4. 자동 데이터 수집 및 분석 실행
5. 생성된 보고서 확인 및 저장

## 🔍 문제해결
- ERP 로그인 실패 시: 계정 정보 및 네트워크 연결 확인
- 데이터 수집 오류 시: ERP 시스템 접근 권한 확인
- 보고서 생성 실패 시: 출력 디렉토리 권한 확인

## 📞 지원
프로젝트 관련 문의사항은 Issues 탭을 통해 등록해 주세요.
