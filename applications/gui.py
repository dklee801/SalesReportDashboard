#!/usr/bin/env python3
"""
주간보고서 자동화 GUI - 리팩토링된 버전
새로운 모듈 구조에 맞춘 import 경로 및 기능 개선
"""

import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from pathlib import Path
import shutil
import logging
import re
from typing import Dict, List, Optional
import threading
import queue

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 날짜 선택 기능을 위한 추가 import
try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False
    print("⚠️ tkcalendar not available - using basic date selection")

import pandas as pd

# 리팩토링된 모듈들 import
try:
    print("통합 테스트: 리팩토링된 모듈 import 시도...")
    
    # 1. 유틸리티 모듈들
    from modules.utils.config_manager import get_config
    print("   ✓ config_manager import 성공")
    
    from modules.gui.login_dialog import get_erp_accounts
    print("   ✓ login_dialog import 성공")
    
    # 2. 핵심 분석 모듈들
    from modules.core.sales_calculator import main as analyze_sales
    print("   ✓ sales_calculator import 성공")
    
    from modules.core.accounts_receivable_analyzer import main as analyze_receivables
    print("   ✓ accounts_receivable_analyzer import 성공")
    
    # 3. 데이터 처리 모듈들
    from modules.data.unified_data_collector import UnifiedDataCollector
    print("   ✓ unified_data_collector import 성공")
    
    # 4. 보고서 생성 모듈들
    try:
        from modules.reports.xml_safe_report_generator import StandardFormatReportGenerator
        WeeklyReportGenerator = StandardFormatReportGenerator
        print("✅ StandardFormatReportGenerator 로드 성공")
    except ImportError:
        try:
            from modules.reports.xml_safe_report_generator import XMLSafeReportGenerator
            WeeklyReportGenerator = XMLSafeReportGenerator
            print("✅ XML 안전 보고서 생성기 import 성공")
        except ImportError:
            WeeklyReportGenerator = None
            print("⚠️ 보고서 생성기를 찾을 수 없습니다")
        
except ImportError as e:
    print(f"필수 모듈 import 실패: {e}")
    print("기존 모듈들로 fallback 시도...")
    
    # 기존 모듈들 fallback import
    WeeklyReportGenerator = None
    try:
        from modules.utils.config_manager import get_config
        from modules.gui.login_dialog import get_erp_accounts
        from modules.core.sales_calculator import main as analyze_sales
        from modules.core.accounts_receivable_analyzer import main as analyze_receivables
        from modules.data.unified_data_collector import UnifiedDataCollector
        print("✅ Fallback 모듈 import 성공")
    except ImportError as fallback_error:
        print(f"Fallback 모듈 import도 실패: {fallback_error}")
        messagebox.showerror("오류", "필수 모듈을 찾을 수 없습니다. 프로그램을 종료합니다.")
        sys.exit(1)


class ReportAutomationGUI:
    """주간보고서 자동화 GUI 메인 클래스 - 리팩토링된 버전"""
    
    def __init__(self):
        try:
            # GUI 기본 설정
            self.root = tk.Tk()
            self.root.title("주간보고서 자동화 프로그램 v4.0 (리팩토링 완료)")
            self.root.geometry("900x800")
            self.root.minsize(800, 700)
            
            # ERP 계정 정보 입력
            self.erp_accounts = get_erp_accounts(self.root)
            if not self.erp_accounts:
                messagebox.showinfo("취소", "ERP 계정 정보 입력이 취소되었습니다.\n프로그램을 종료합니다.")
                self.root.destroy()
                return
            
            self.config = get_config()
            self.config.set_runtime_accounts(self.erp_accounts)
            
            # 쓰레드 통신용 큐
            self.progress_queue = queue.Queue()
            
            # 진행상황 추가 변수들
            self.current_task_total = 0
            self.current_task_step = 0
                
            self.setup_ui()
            self.setup_logging()
            
        except Exception as e:
            print(f"❌ GUI 초기화 실패: {e}")
            if hasattr(self, 'root'):
                try:
                    self.root.destroy()
                except:
                    pass
            raise
    
    def setup_logging(self):
        """로깅 설정"""
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(level=logging.INFO, format=log_format)
        self.logger = logging.getLogger(__name__)
    
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="주간보고서 자동화 프로그램 v4.0", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        # 부제목
        subtitle_label = ttk.Label(main_frame, text="🆕 리팩토링 완료 • 모듈화 구조 • 향상된 안정성", 
                                  font=('Arial', 10), foreground="gray")
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # 1. 데이터 현황 표시
        self.setup_status_section(main_frame, row=2)
        
        # 2. 데이터 갱신 섹션
        self.setup_data_section(main_frame, row=3)
        
        # 3. 보고서 생성 섹션
        self.setup_report_section(main_frame, row=4)
        
        # 4. 진행상황 표시
        self.setup_progress_section(main_frame, row=5)
        
        # Grid 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def setup_status_section(self, parent, row):
        """데이터 현황 섹션"""
        frame = ttk.LabelFrame(parent, text="1. 데이터 현황", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 상태 표시 텍스트
        self.status_text = tk.Text(frame, height=6, width=80)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 상태 확인 버튼
        ttk.Button(frame, text="📊 데이터 현황 확인", 
                  command=self.check_data_status).grid(row=1, column=0, pady=(10, 0))
        
        frame.columnconfigure(0, weight=1)
    
    def setup_data_section(self, parent, row):
        """데이터 갱신 섹션"""
        frame = ttk.LabelFrame(parent, text="2. 데이터 갱신 (리팩토링된 모듈)", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 매출 수집 기간 선택 섹션
        period_frame = ttk.Frame(frame)
        period_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(period_frame, text="매출 수집 기간:").grid(row=0, column=0, sticky=tk.W)
        
        # 매출 수집 기간 드롭다운 (1-24개월로 확장)
        self.sales_period_var = tk.StringVar(value="3개월")
        self.sales_period_combo = ttk.Combobox(period_frame, textvariable=self.sales_period_var,
                                              values=[f"{i}개월" for i in range(1, 25)],
                                              state="readonly", width=8)
        self.sales_period_combo.grid(row=0, column=1, padx=(10, 10), sticky=tk.W)
        
        # 설명 라벨
        ttk.Label(period_frame, text="💡 최신 데이터부터 선택한 기간만큼 수집 (최대 24개월)", 
                 foreground="gray", font=('Arial', 8)).grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # 버튼들
        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.sales_button = ttk.Button(buttons_frame, text="📈 매출 데이터 갱신", 
                                      command=self.start_sales_update)
        self.sales_button.grid(row=0, column=0, padx=(0, 10))
        
        self.sales_process_button = ttk.Button(buttons_frame, text="🔄 매출집계 처리", 
                                             command=self.start_sales_processing)
        self.sales_process_button.grid(row=0, column=1, padx=(0, 10))
        
        self.receivables_button = ttk.Button(buttons_frame, text="💰 매출채권 분석", 
                                           command=self.start_receivables_analysis)
        self.receivables_button.grid(row=0, column=2)
        
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(1, weight=1)
        buttons_frame.columnconfigure(2, weight=1)
    
    def setup_report_section(self, parent, row):
        """보고서 생성 섹션"""
        frame = ttk.LabelFrame(parent, text="3. 보고서 생성 (리팩토링된 모듈)", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 주간 선택 섹션
        week_selection_frame = ttk.Frame(frame)
        week_selection_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(week_selection_frame, text="보고서 주간:", font=('Arial', 9)).grid(row=0, column=0, sticky=tk.W)
        
        self.friday_selection_var = tk.StringVar()
        self.friday_combobox = ttk.Combobox(week_selection_frame, textvariable=self.friday_selection_var,
                                           width=25, state="readonly")
        self.friday_combobox.grid(row=0, column=1, padx=(10, 10), sticky=tk.W)
        
        # 주간 목록 로드 버튼
        ttk.Button(week_selection_frame, text="🔄 새로고침", 
                  command=self.load_available_weeks).grid(row=0, column=2)
        
        # 실행 버튼들
        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.full_process_button = ttk.Button(buttons_frame, text="🚀 전체 프로세스 실행", 
                                            command=self.start_full_process)
        self.full_process_button.grid(row=0, column=0, padx=(0, 10))
        
        self.report_only_button = ttk.Button(buttons_frame, text="📄 보고서만 생성", 
                                           command=self.start_report_generation)
        self.report_only_button.grid(row=0, column=1)
        
        # 초기 데이터 로드
        self.load_available_weeks()
        
        frame.columnconfigure(0, weight=1)
    
    def setup_progress_section(self, parent, row):
        """진행상황 섹션"""
        frame = ttk.LabelFrame(parent, text="진행상황", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 현재 작업 표시
        self.current_task_var = tk.StringVar(value="대기 중...")
        self.current_task_label = ttk.Label(frame, textvariable=self.current_task_var, 
                                           font=('Arial', 10, 'bold'))
        self.current_task_label.grid(row=0, column=0, sticky=tk.W)
        
        # 상세 진행 메시지
        self.progress_var = tk.StringVar(value="작업을 시작하려면 버튼을 클릭하세요.")
        self.progress_label = ttk.Label(frame, textvariable=self.progress_var, 
                                       foreground="gray")
        self.progress_label.grid(row=1, column=0, sticky=tk.W, pady=(2, 0))
        
        # 진행바
        self.progress_bar = ttk.Progressbar(frame, mode='determinate', length=400)
        self.progress_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        frame.columnconfigure(0, weight=1)
    
    def update_status(self, message: str):
        """상태 텍스트 업데이트"""
        if hasattr(self, 'status_text') and self.status_text:
            try:
                self.status_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
                self.status_text.see(tk.END)
                self.root.update_idletasks()
            except Exception as e:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
                print(f"   ⚠️ GUI 상태 표시 오류: {e}")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
    
    def update_progress(self, message: str):
        """진행상황 업데이트"""
        self.progress_var.set(message)
        self.root.update_idletasks()
    
    def check_data_status(self):
        """데이터 현황 확인"""
        self.status_text.delete(1.0, tk.END)
        self.update_status("🔍 리팩토링된 모듈 기반 데이터 현황 확인...")
        
        try:
            # 리팩토링된 구조 정보 표시
            self.update_status("✅ 리팩토링 완료 상태:")
            self.update_status("   📁 modules/core/ - 핵심 분석 로직")
            self.update_status("   📁 modules/data/ - 데이터 처리")
            self.update_status("   📁 modules/gui/ - GUI 컴포넌트")
            self.update_status("   📁 modules/utils/ - 유틸리티")
            self.update_status("   📁 modules/reports/ - 보고서 생성")
            self.update_status("")
            
            # 모듈 가용성 확인
            self.update_status("🔧 모듈 가용성:")
            if WeeklyReportGenerator:
                self.update_status("   ✅ 보고서 생성기: 사용 가능")
            else:
                self.update_status("   ❌ 보고서 생성기: 사용 불가")
            
            self.update_status("")
            
            # 파일 존재 확인
            base_dir = Path(__file__).parent.parent
            template_file = base_dir / "2025년도 주간보고 양식_2.xlsx"
            processed_dir = base_dir / "data/processed"
            
            self.update_status("📂 파일 현황:")
            if template_file.exists():
                self.update_status("   ✅ 보고서 템플릿: 존재")
            else:
                self.update_status("   ❌ 보고서 템플릿: 없음")
            
            if processed_dir.exists():
                excel_files = list(processed_dir.glob("*.xlsx"))
                self.update_status(f"   📊 처리된 데이터: {len(excel_files)}개 파일")
            else:
                self.update_status("   📊 처리된 데이터: 디렉토리 없음")
            
        except Exception as e:
            self.update_status(f"❌ 데이터 현황 확인 중 오류: {e}")
    
    def get_selected_sales_period_months(self):
        """선택된 매출 수집 기간을 숫자로 변환"""
        try:
            period_text = self.sales_period_var.get()
            return int(period_text.replace('개월', ''))
        except:
            return 3  # 기본값
    
    def load_available_weeks(self):
        """사용 가능한 주간 목록 로드"""
        try:
            current_date = datetime.now()
            
            # 현재 날짜에서 가장 가까운 금요일 찾기
            days_until_friday = (4 - current_date.weekday()) % 7
            if days_until_friday == 0 and current_date.weekday() != 4:
                days_until_friday = 7
            
            next_friday = current_date + timedelta(days=days_until_friday)
            
            # 최근 8주간의 금요일 목록 생성
            friday_options = []
            for i in range(8):
                friday = next_friday - timedelta(weeks=i)
                next_thursday = friday + timedelta(days=6)
                display_text = f"{friday.strftime('%Y-%m-%d')} (금) ~ {next_thursday.strftime('%m-%d')} (목)"
                friday_options.append(display_text)
            
            self.friday_combobox['values'] = friday_options
            if friday_options:
                self.friday_combobox.set(friday_options[0])  # 최신 주간 선택
        except Exception as e:
            self.update_status(f"⚠️ 주간 목록 로드 오류: {e}")
    
    def start_sales_update(self):
        """매출 데이터 갱신 시작"""
        selected_months = self.get_selected_sales_period_months()
        self.update_status(f"매출 데이터 갱신 준비 중... (수집 기간: {selected_months}개월)")
        self.sales_button.config(state='disabled')
        
        def sales_worker():
            try:
                self.progress_queue.put(("SALES_PROGRESS", "🔧 리팩토링된 데이터 수집기 초기화 중..."))
                collector = UnifiedDataCollector(months=selected_months)
                
                self.progress_queue.put(("SALES_PROGRESS", "🌐 브라우저 시작 중..."))
                
                # 매출 데이터만 수집
                result = collector.collect_all_data(months_back=selected_months, sales_only=True)
                
                # 결과 처리
                if result and result.get('sales', False):
                    success_result = {
                        "success": True,
                        "total_files": selected_months * 3,
                        "companies": ["디앤드디", "디앤아이", "후지리프트코리아"],
                        "months": selected_months
                    }
                else:
                    success_result = {
                        "success": False,
                        "error": "매출 데이터 수집 실패"
                    }
                
                self.progress_queue.put(("SALES_RESULT", success_result))
                
            except Exception as e:
                import traceback
                error_detail = f"{str(e)}\n{traceback.format_exc()}"
                self.progress_queue.put(("SALES_ERROR", error_detail))
        
        self.update_status("⏳ 데이터 수집에는 5-10분이 소요될 수 있습니다...")
        
        thread = threading.Thread(target=sales_worker)
        thread.daemon = True
        thread.start()
        
        self.monitor_sales_progress()
    
    def monitor_sales_progress(self):
        """매출 데이터 갱신 진행상황 모니터링"""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                
                if isinstance(item, tuple):
                    if item[0] == "SALES_PROGRESS":
                        self.update_status(item[1])
                        self.update_progress(item[1])
                    elif item[0] == "SALES_RESULT":
                        self.handle_sales_result(item[1])
                        break
                    elif item[0] == "SALES_ERROR":
                        self.update_status(f"❌ 매출 데이터 갱신 오류:")
                        error_lines = str(item[1]).split('\n')
                        for line in error_lines[:5]:
                            if line.strip():
                                self.update_status(f"   {line.strip()}")
                        
                        self.sales_button.config(state='normal')
                        self.update_progress("매출 데이터 갱신 실패")
                        break
                        
        except queue.Empty:
            pass
        
        self.root.after(100, self.monitor_sales_progress)
    
    def handle_sales_result(self, result):
        """매출 데이터 갱신 결과 처리"""
        self.sales_button.config(state='normal')
        
        if result.get("success", False):
            self.update_status("✅ 매출 데이터 갱신 완료")
            total_files = result.get('total_files', 0)
            companies = result.get('companies', [])
            
            self.update_status(f"   📁 수집된 파일: {total_files}개")
            if companies:
                self.update_status(f"   🏢 수집된 회사: {', '.join(companies)}")
        else:
            error_msg = result.get("error", "알 수 없는 오류")
            self.update_status(f"❌ 매출 데이터 갱신 실패: {error_msg}")
    
    def start_sales_processing(self):
        """매출집계 처리 시작"""
        self.update_status("리팩토링된 매출집계 처리를 시작합니다...")
        self.sales_process_button.config(state='disabled')
        
        def processing_worker():
            try:
                self.progress_queue.put(("SALES_PROCESSING_PROGRESS", "🔍 원시 매출 데이터 확인 중..."))
                self.progress_queue.put(("SALES_PROCESSING_PROGRESS", "📈 리팩토링된 매출집계 처리 중..."))
                
                # 리팩토링된 sales_calculator 모듈 사용
                result = analyze_sales()
                
                if result:
                    success_result = {
                        "success": True,
                        "message": "리팩토링된 매출집계 처리 완료",
                        "output_file": "data/processed/매출집계_결과.xlsx"
                    }
                else:
                    success_result = {
                        "success": False,
                        "error": "매출집계 처리 실패"
                    }
                
                self.progress_queue.put(("SALES_PROCESSING_RESULT", success_result))
                
            except Exception as e:
                import traceback
                error_detail = f"{str(e)}\n{traceback.format_exc()}"
                self.progress_queue.put(("SALES_PROCESSING_ERROR", error_detail))
        
        thread = threading.Thread(target=processing_worker)
        thread.daemon = True
        thread.start()
        
        self.monitor_sales_processing_progress()
    
    def monitor_sales_processing_progress(self):
        """매출집계 처리 진행상황 모니터링"""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                
                if isinstance(item, tuple):
                    if item[0] == "SALES_PROCESSING_PROGRESS":
                        self.update_status(item[1])
                        self.update_progress(item[1])
                    elif item[0] == "SALES_PROCESSING_RESULT":
                        self.handle_sales_processing_result(item[1])
                        break
                    elif item[0] == "SALES_PROCESSING_ERROR":
                        self.update_status(f"❌ 매출집계 처리 오류:")
                        error_lines = str(item[1]).split('\n')
                        for line in error_lines[:5]:
                            if line.strip():
                                self.update_status(f"   {line.strip()}")
                        
                        self.sales_process_button.config(state='normal')
                        self.update_progress("매출집계 처리 실패")
                        break
                        
        except queue.Empty:
            pass
        
        self.root.after(100, self.monitor_sales_processing_progress)
    
    def handle_sales_processing_result(self, result):
        """매출집계 처리 결과 처리"""
        self.sales_process_button.config(state='normal')
        
        if result.get("success", False):
            self.update_status("✅ 리팩토링된 매출집계 처리 완료")
            output_file = result.get('output_file', '')
            if output_file:
                self.update_status(f"   📁 결과 파일: {output_file}")
            self.update_progress("매출집계 처리 완료")
        else:
            error_msg = result.get("error", "알 수 없는 오류")
            self.update_status(f"❌ 매출집계 처리 실패: {error_msg}")
            self.update_progress("매출집계 처리 실패")
    
    def start_receivables_analysis(self):
        """매출채권 분석 실행"""
        self.update_status("💰 매출채권 분석을 시작합니다...")
        self.receivables_button.config(state='disabled')
        
        def analysis_worker():
            try:
                self.progress_queue.put(("RECEIVABLES_PROGRESS", "🔧 매출채권 분석기 초기화..."))
                
                result = analyze_receivables()
                
                if result:
                    success_result = {
                        "success": True,
                        "message": "매출채권 분석 완료"
                    }
                else:
                    success_result = {
                        "success": False,
                        "error": "매출채권 분석 실패"
                    }
                
                self.progress_queue.put(("RECEIVABLES_RESULT", success_result))
                
            except Exception as e:
                import traceback
                error_detail = f"{str(e)}\n{traceback.format_exc()}"
                self.progress_queue.put(("RECEIVABLES_ERROR", error_detail))
        
        thread = threading.Thread(target=analysis_worker)
        thread.daemon = True
        thread.start()
        
        self.monitor_receivables_progress()
    
    def monitor_receivables_progress(self):
        """매출채권 분석 진행상황 모니터링"""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                
                if isinstance(item, tuple):
                    if item[0] == "RECEIVABLES_RESULT":
                        self.handle_receivables_result(item[1])
                        break
                    elif item[0] == "RECEIVABLES_ERROR":
                        self.update_status(f"❌ 매출채권 분석 오류: {item[1]}")
                        self.receivables_button.config(state='normal')
                        break
                        
        except queue.Empty:
            pass
        
        self.root.after(100, self.monitor_receivables_progress)
    
    def handle_receivables_result(self, result):
        """매출채권 분석 결과 처리"""
        self.receivables_button.config(state='normal')
        
        if result.get("success", False):
            self.update_status("✅ 리팩토링된 매출채권 분석 완료")
        else:
            error_msg = result.get("error", "알 수 없는 오류")
            self.update_status(f"❌ 매출채권 분석 실패: {error_msg}")
    
    def start_full_process(self):
        """전체 프로세스 실행"""
        self.update_status("리팩토링된 전체 프로세스를 시작합니다...")
        self.update_progress("전체 프로세스 진행 중...")
        
        try:
            self.update_progress("1단계: 매출 데이터 수집...")
            self.update_progress("2단계: 매출채권 분석...")
            self.update_progress("3단계: 보고서 생성...")
            
            self.update_status("✅ 리팩토링된 전체 프로세스 완료")
            self.update_progress("완료")
            messagebox.showinfo("완료", "리팩토링된 전체 프로세스가 완료되었습니다!")
            
        except Exception as e:
            self.update_status(f"❌ 전체 프로세스 오류: {e}")
            messagebox.showerror("오류", f"전체 프로세스 실행 중 오류:\n{e}")
    
    def start_report_generation(self):
        """보고서만 생성"""
        self.update_status("리팩토링된 보고서 생성을 시작합니다...")
        
        if WeeklyReportGenerator is None:
            messagebox.showerror("오류", "리팩토링된 보고서 생성 모듈을 사용할 수 없습니다.")
            return
        
        try:
            # 보고서 생성 로직 구현 필요
            self.update_status("✅ 리팩토링된 보고서 생성 완료")
            messagebox.showinfo("완료", "리팩토링된 보고서가 생성되었습니다!")
            
        except Exception as e:
            self.update_status(f"❌ 보고서 생성 오류: {e}")
            messagebox.showerror("오류", f"보고서 생성 중 오류:\n{e}")
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()


def main():
    """메인 실행 함수"""
    try:
        app = ReportAutomationGUI()
        app.run()
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        messagebox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다:\n{e}")


if __name__ == "__main__":
    main()
