import json
import time
import os
import shutil
import calendar
import re
import logging  # logging import 추가
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
import sys
from abc import ABC, abstractmethod
from typing import Dict, List, Tuple, Optional

# 리팩토링된 경로 설정
sys.path.append(str(Path(__file__).parent.parent))

# 설정 관리자 import (새 구조)
from modules.utils.config_manager import get_config

# BaseDataCollector import (새 구조)
from modules.data.collectors.base_collector import BaseDataCollector

# 데이터 검증기 import (새 구조)
try:
    from modules.data.validators.sales_data_validator import SalesDataValidator
    VALIDATOR_AVAILABLE = True
except ImportError:
    print("⚠️ 매출 데이터 검증기를 찾을 수 없습니다. 검증 기능이 비활성화됩니다.")
    VALIDATOR_AVAILABLE = False

class SalesDataCollector(BaseDataCollector):
    """매출 데이터 수집 클래스 - 백업본의 작동하는 로직 사용 (리팩토링됨)"""
    
    def __init__(self, headless_mode=None):
        super().__init__(headless_mode)
        self.accounts = self.config.get_accounts()
        
        # 데이터 검증기 초기화
        if VALIDATOR_AVAILABLE:
            self.validator = SalesDataValidator()
            self.logger = logging.getLogger('SalesDataCollector')
            print("✅ 매출 데이터 검증기 활성화")
        else:
            self.validator = None
            self.logger = logging.getLogger('SalesDataCollector')
            print("⚠️ 매출 데이터 검증기 비활성화")
        
    def get_target_accounts(self) -> List[Dict[str, str]]:
        return self.accounts
    
    def generate_monthly_date_ranges(self, num_months=None, start_date=None, end_date=None):
        """월별 날짜 범위 생성"""
        if start_date and end_date:
            print(f"🔍 수동 지정된 기간: {start_date} ~ {end_date}")
            return [(start_date, end_date)]

        if num_months is None:
            num_months = self.config.get_default_num_months()
        
        today = datetime.today()
        print(f"🗓️ 기준일: {today.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📅 수집할 개월 수: {num_months}개월")
        
        date_ranges = []
        for i in range(num_months):
            # 더 정확한 월 계산 방식 사용
            target_year = today.year
            target_month = today.month - i
            
            # 년도 넘김 처리
            while target_month <= 0:
                target_month += 12
                target_year -= 1
            
            # 해당 월의 첫째 날과 마지막 날 계산
            first_day = datetime(target_year, target_month, 1)
            if target_month == 12:
                last_day = datetime(target_year + 1, 1, 1) - timedelta(days=1)
            else:
                last_day = datetime(target_year, target_month + 1, 1) - timedelta(days=1)
            
            start_str = first_day.strftime('%Y%m%d')
            end_str = last_day.strftime('%Y%m%d')
            
            print(f"결과: {start_str}_{end_str} ({target_year}년 {target_month}월)")
            date_ranges.append((start_str, end_str))
            
        return date_ranges

    def get_month_xpath(self, month: str, company_name: str) -> str:
        """회사별 월 xpath 반환"""
        company_config = self.config.get_company_config(company_name)
        xpath_format = company_config.get("month_xpath_format", "//a[text()='{month}']")
        
        if "후지리프트" in company_name:
            int_month = int(month)
            result_xpath = xpath_format.format(int_month=int_month)
            return result_xpath
        else:
            result_xpath = xpath_format.format(month=month)
            return result_xpath

    def generate_save_path(self, start_date: str) -> Path:
        """저장 경로 생성"""
        year = start_date[:4]
        save_dir = self.config.get_sales_raw_data_dir() / year
        save_dir.mkdir(parents=True, exist_ok=True)
        return save_dir

    def navigate_to_target_page(self, driver, account):
        """매출조회 페이지로 이동"""
        wait = WebDriverWait(driver, self.selenium_config.get("implicit_wait", 10))
        company_name = account.get("company_name", "")
        
        print(f"   📋 {company_name} 매출조회 페이지 이동...")
        
        # 매출 메뉴로 이동
        self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "link_depth1_4"))))
        time.sleep(2)
        self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "link_depth4_492"))))
        time.sleep(5)

        # 검색창 열기
        try:
            search_area = driver.find_element(By.CLASS_NAME, "wrapper-header-search")
            if not search_area.is_displayed():
                self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "search"))))
        except:
            self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "search"))))
        time.sleep(3)

    def set_search_criteria(self, driver, start_date: str, end_date: str, company_name: str, attempt: int = 1, **kwargs):
        """검색 조건 설정 - 수정된 버전 (매번 검색창 재오픈)"""
        wait = WebDriverWait(driver, 10)
        
        print(f"   📅 날짜 설정 시작: {start_date} ~ {end_date} (시도 #{attempt})")
        
        # 1. 매번 검색창을 새로 열어서 초기화
        try:
            print(f"   🔄 검색창 재오픈...")
            # 검색창이 이미 열려있으면 닫기
            try:
                search_area = driver.find_element(By.CLASS_NAME, "wrapper-header-search")
                if search_area.is_displayed():
                    search_close_btn = driver.find_element(By.ID, "search")
                    self.js_click(driver, search_close_btn)
                    time.sleep(1)
            except:
                pass
            
            # 검색창 다시 열기
            search_btn = wait.until(EC.element_to_be_clickable((By.ID, "search")))
            self.js_click(driver, search_btn)
            time.sleep(3)
            print(f"   ✅ 검색창 재오픈 완료")
            
        except Exception as e:
            print(f"   ⚠️ 검색창 재오픈 실패: {e}")
        
        # 2. 기본 안정화 대기
        time.sleep(2)
        
        # 3. 적용양식 설정
        try:
            labels = driver.find_elements(By.CSS_SELECTOR, "div.selectbox-label")
            selected_label = None
            for label in labels:
                if "판매조회" in label.text:
                    selected_label = label
                    break

            if selected_label is None:
                print(f"   ❌ 적용양식 라벨 요소를 찾을 수 없음")
            else:
                current_value = selected_label.text.strip()
                target_label = "selenium_data(이동규)"
                
                if current_value != target_label:
                    print(f"   🔁 현재 적용양식: '{current_value}' → 변경 필요")

                    # 드롭다운 열기
                    dropdown_button = selected_label.find_element(By.XPATH, "./ancestor::button")
                    self.js_click(driver, dropdown_button)

                    # 드롭다운 로딩 대기
                    WebDriverWait(driver, 5).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.dropdown-menu-item.show"))
                    )

                    # 타겟 항목 선택
                    options = driver.find_elements(By.CSS_SELECTOR, "ul.dropdown-menu-item.show li a")
                    found = False
                    for option in options:
                        text = option.text.strip()
                        if text == target_label:
                            self.js_click(driver, option)
                            print(f"   ✅ 적용양식 선택 완료")
                            found = True
                            break

                    if not found:
                        print(f"   ❌ 드롭다운 내 'selenium_data(이동규)' 항목 없음")
                else:
                    print(f"   ⏭️ 적용양식 이미 선택되어 있음 → 스킵")

            time.sleep(2)

        except Exception as e:
            print(f"   ❌ 적용양식 처리 오류: {e}")

        # 4. 날짜 설정 - 원본 방식 (검색 버튼 클릭 전에)
        print(f"   📅 날짜 설정 시작...")
        try:
            date_inputs = driver.find_elements(By.CSS_SELECTOR, "input#day")
            year_buttons = driver.find_elements(By.CSS_SELECTOR, "button[data-id='year']")
            month_buttons = driver.find_elements(By.CSS_SELECTOR, "button[data-id='month']")
            
            print(f"   발견된 요소: 날짜입력({len(date_inputs)}), 년도버튼({len(year_buttons)}), 월버튼({len(month_buttons)})")

            # 시작일 설정
            print(f"   📅 시작일 설정: {start_date[:4]}년 {start_date[4:6]}월 {start_date[6:8]}일")
            
            # 시작 년도 설정
            if len(year_buttons) >= 1:
                self.js_click(driver, year_buttons[0])
                time.sleep(1)
                year_option = driver.find_element(By.XPATH, f"//a[text()='{start_date[:4]}']") 
                self.js_click(driver, year_option)
                print(f"   ✅ 시작 년도: {start_date[:4]}")
                time.sleep(1)
            
            # 시작 월 설정
            if len(month_buttons) >= 1:
                start_month = start_date[4:6]
                print(f"   📅 시작 월 설정 시도: {start_month}")
                
                self.js_click(driver, month_buttons[0])
                time.sleep(1)
                
                # 회사별 최적화된 XPath 패턴 적용
                if "후지리프트" in company_name:
                    # 후지리프트: "월" 형식 우선
                    month_xpath_patterns = [
                        f"//a[text()='{str(int(start_month))}월']",  # 후지리프트 우선: "8월"
                        f"//a[contains(text(), '{str(int(start_month))}월')]",  # 후지리프트 부분매칭
                        f"//a[text()='{str(int(start_month))}']",  # 예비: "8"
                        f"//a[text()='{start_month}']",  # 예비: "08"
                    ]
                else:
                    # 디앤디, 디앤아이: 숫자 형식 우선
                    month_xpath_patterns = [
                        f"//a[text()='{start_month}']",  # 디앤아이 우선: "08"
                        f"//a[text()='{str(int(start_month))}']",  # 디앤아이 대안: "8"
                        f"//a[text()='{str(int(start_month))}월']",  # 예비: "8월"
                        f"//a[contains(text(), '{str(int(start_month))}월')]",  # 예비 부분매칭
                    ]
                
                success = False
                for i, xpath_pattern in enumerate(month_xpath_patterns, 1):
                    try:
                        month_option = driver.find_element(By.XPATH, xpath_pattern)
                        self.js_click(driver, month_option)
                        print(f"   ✅ 시작 월: {start_month} (패턴 #{i} 성공: {xpath_pattern})")
                        success = True
                        break
                    except Exception as e:
                        print(f"   ⚠️ 패턴 #{i} 실패: {xpath_pattern} - {str(e)[:50]}...")
                        continue
                
                if not success:
                    # 최후의 수단: 모든 월 옵션 출력하여 디버깅
                    try:
                        all_options = driver.find_elements(By.XPATH, "//a[contains(@class, '') or contains(text(), '월') or text()<=12]")
                        print(f"   🔍 사용가능한 월 옵션들:")
                        for opt in all_options[:10]:  # 최대 10개만 출력
                            print(f"      - '{opt.text}'")
                    except:
                        pass
                    raise Exception(f"모든 월 선택 패턴 실패: {start_month}")
                
                time.sleep(1)
            
            # 시작 일 설정
            if len(date_inputs) >= 1:
                date_inputs[0].clear()
                date_inputs[0].send_keys(start_date[6:8])
                print(f"   ✅ 시작 일: {start_date[6:8]}")

            # 종료일 설정
            print(f"   📅 종료일 설정: {end_date[:4]}년 {end_date[4:6]}월 {end_date[6:8]}일")
            
            # 종료 년도 설정
            if len(year_buttons) >= 2:
                self.js_click(driver, year_buttons[1])
                time.sleep(1)
                year_option = driver.find_element(By.XPATH, f"//a[text()='{end_date[:4]}']") 
                self.js_click(driver, year_option)
                print(f"   ✅ 종료 년도: {end_date[:4]}")
                time.sleep(1)
            
            # 종료 월 설정
            if len(month_buttons) >= 2:
                end_month = end_date[4:6]
                print(f"   📅 종료 월 설정 시도: {end_month}")
                
                self.js_click(driver, month_buttons[1])
                time.sleep(1)
                
                # 회사별 최적화된 XPath 패턴 적용
                if "후지리프트" in company_name:
                    # 후지리프트: "월" 형식 우선
                    month_xpath_patterns = [
                        f"//a[text()='{str(int(end_month))}월']",  # 후지리프트 우선: "8월"
                        f"//a[contains(text(), '{str(int(end_month))}월')]",  # 후지리프트 부분매칭
                        f"//a[text()='{str(int(end_month))}']",  # 예비: "8"
                        f"//a[text()='{end_month}']",  # 예비: "08"
                    ]
                else:
                    # 디앤디, 디앤아이: 숫자 형식 우선
                    month_xpath_patterns = [
                        f"//a[text()='{end_month}']",  # 디앤아이 우선: "08"
                        f"//a[text()='{str(int(end_month))}']",  # 디앤아이 대안: "8"
                        f"//a[text()='{str(int(end_month))}월']",  # 예비: "8월"
                        f"//a[contains(text(), '{str(int(end_month))}월')]",  # 예비 부분매칭
                    ]
                
                success = False
                for i, xpath_pattern in enumerate(month_xpath_patterns, 1):
                    try:
                        month_option = driver.find_element(By.XPATH, xpath_pattern)
                        self.js_click(driver, month_option)
                        print(f"   ✅ 종료 월: {end_month} (패턴 #{i} 성공: {xpath_pattern})")
                        success = True
                        break
                    except Exception as e:
                        print(f"   ⚠️ 패턴 #{i} 실패: {xpath_pattern} - {str(e)[:50]}...")
                        continue
                
                if not success:
                    # 최후의 수단: 모든 월 옵션 출력하여 디버깅
                    try:
                        all_options = driver.find_elements(By.XPATH, "//a[contains(@class, '') or contains(text(), '월') or text()<=12]")
                        print(f"   🔍 사용가능한 월 옵션들:")
                        for opt in all_options[:10]:  # 최대 10개만 출력
                            print(f"      - '{opt.text}'")
                    except:
                        pass
                    raise Exception(f"모든 월 선택 패턴 실패: {end_month}")
                
                time.sleep(1)
            
            # 종료 일 설정
            if len(date_inputs) >= 2:
                date_inputs[1].clear()
                date_inputs[1].send_keys(end_date[6:8])
                print(f"   ✅ 종료 일: {end_date[6:8]}")

            # 설정 완료 후 다른 곳 클릭
            try:
                wrapper_title = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "wrapper-title"))
                )
                self.js_click(driver, wrapper_title)
                print(f"   ✅ wrapper-title 클릭 성공")
            except Exception as e1:
                print(f"   ⚠️ wrapper-title 클릭 실패: {e1}")
                try:
                    driver.execute_script("document.body.click();")
                    print(f"   ✅ body 클릭으로 대체")
                except Exception as e2:
                    print(f"   ⚠️ body 클릭도 실패: {e2} - 계속 진행")
            
            time.sleep(2)
            
            print(f"   ✅ 날짜 설정 완료: {start_date} ~ {end_date}")
            
            # 5. 검색 버튼 클릭 (필수)
            print(f"   🔍 검색 버튼 클릭...")
            try:
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "header_search"))
                )
                self.js_click(driver, search_button)
                print(f"   ✅ 검색 버튼 클릭 성공")
                time.sleep(5)  # 검색 결과 로딩 대기
            except Exception as e:
                print(f"   ❌ 검색 버튼 클릭 실패: {e}")
                raise
            
            # 6. "확인" 탭 클릭 (신규 기능)
            print(f"   📄 '확인' 탭 클릭...")
            try:
                confirmed_tab = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@id='Y']//a"))
                )
                self.js_click(driver, confirmed_tab)
                print(f"   ✅ '확인' 탭 클릭 성공")
                time.sleep(3)  # 확인 데이터 로딩 대기
            except Exception as e:
                print(f"   ⚠️ '확인' 탭 클릭 실패: {e}")
                print(f"   → '전체' 탭으로 계속 진행 (기존 데이터 사용)")
                # 실패해도 계속 진행 - 전체 데이터로 다운로드

        except Exception as e:
            print(f"   ❌ 날짜 설정 실패: {e}")
            raise

    def download_and_save(self, driver, company_name: str, start_date: str, end_date: str, **kwargs) -> bool:
        """다운로드 및 저장"""
        try:
            print(f"   📊 {company_name} 엑셀 다운로드 시작...")
            
            # 엑셀 버튼 클릭
            wait = WebDriverWait(driver, 10)
            excel_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div#footer_toolbar_toolbar_item_excel button")))
            self.js_click(driver, excel_button)
            
            # 다운로드 대기
            filename = f"{company_name}_매출_조회_{start_date}_{end_date}.xlsx"
            downloaded_file = self.wait_for_download(company_name, filename)
            
            if downloaded_file:
                # 저장 경로로 이동
                save_dir = self.generate_save_path(start_date)
                final_path = save_dir / filename
                
                shutil.move(str(downloaded_file), str(final_path))
                print(f"   ✅ 파일 저장 완료: {final_path}")
                return True
            else:
                print(f"   ❌ 다운로드 실패")
                return False
                
        except Exception as e:
            print(f"   ❌ 다운로드 오류: {e}")
            return False

    def collect_data(self, num_months=None, start_date=None, end_date=None, progress_callback=None):
        """매출 데이터 수집 메인 함수 (리팩토링됨)"""
        print("🚀 매출 데이터 수집 시작 (리팩토링됨)")
        
        date_ranges = self.generate_monthly_date_ranges(num_months, start_date, end_date)
        accounts = self.get_target_accounts()
        
        total_tasks = len(accounts) * len(date_ranges)
        current_task = 0
        
        for account in accounts:
            company_name = account["company_name"]
            print(f"\n💼 {company_name} 데이터 수집 시작")
            
            driver = None
            try:
                driver = self.launch_driver()
                
                # 로그인
                if not self.basic_login(driver, account):
                    print(f"❌ {company_name} 로그인 실패")
                    continue
                
                # 매출조회 페이지로 이동
                self.navigate_to_target_page(driver, account)
                
                # 날짜별 데이터 수집
                for start_date, end_date in date_ranges:
                    try:
                        current_task += 1
                        print(f"\n📅 [{current_task}/{total_tasks}] {start_date} ~ {end_date} 수집 중...")
                        
                        if progress_callback:
                            progress = (current_task / total_tasks) * 100
                            progress_callback(f"{company_name} {start_date}~{end_date} 수집 중", progress)
                        
                        # 검색 조건 설정
                        self.set_search_criteria(driver, start_date, end_date, company_name)
                        
                        # 다운로드 및 저장
                        if self.download_and_save(driver, company_name, start_date, end_date):
                            print(f"   ✅ {start_date}~{end_date} 수집 완료")
                        else:
                            print(f"   ❌ {start_date}~{end_date} 수집 실패")
                            
                    except Exception as e:
                        print(f"   ❌ {start_date}~{end_date} 수집 중 오류: {e}")
                        
            except Exception as e:
                print(f"❌ {company_name} 데이터 수집 실패: {e}")
                
            finally:
                if driver:
                    driver.quit()
                    print(f"🔌 {company_name} 브라우저 종료")
        
        print(f"\n🎉 매출 데이터 수집 완료 (리팩토링됨)")

class ReceivablesDataCollector(BaseDataCollector):
    """매출채권 데이터 수집 클래스 - 금요일 기준 (리팩토링됨)"""
    
    def __init__(self, headless_mode=None):
        super().__init__(headless_mode)
        self.accounts = self.config.get_accounts()
        
    def get_target_accounts(self) -> List[Dict[str, str]]:
        return self.accounts
    
    def get_friday_date(self, target_date: str = None) -> str:
        """금요일 날짜 계산 - 월~금 기준"""
        if target_date:
            base_date = datetime.strptime(target_date, '%Y%m%d')
        else:
            base_date = datetime.today()
        
        # 가장 가까운 금요일 찾기 (4 = 금요일)
        days_since_friday = (base_date.weekday() - 4) % 7
        if days_since_friday == 0 and base_date.weekday() == 4:
            # 오늘이 금요일이면 그대로
            friday_date = base_date
        else:
            # 이전 금요일로 이동
            friday_date = base_date - timedelta(days=days_since_friday)
        
        return friday_date.strftime('%Y%m%d')

    def navigate_to_target_page(self, driver, account):
        """매출채권조회 페이지로 이동"""
        wait = WebDriverWait(driver, self.selenium_config.get("implicit_wait", 10))
        company_name = account.get("company_name", "")
        
        print(f"   📋 {company_name} 매출채권조회 페이지 이동...")
        
        # 매출채권 메뉴로 이동
        self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "link_depth1_4"))))
        time.sleep(2)
        self.js_click(driver, wait.until(EC.element_to_be_clickable((By.ID, "link_depth4_496"))))
        time.sleep(5)

    def set_search_criteria(self, driver, target_date: str, company_name: str, **kwargs):
        """검색 조건 설정 - 금요일 기준"""
        print(f"   📅 매출채권 기준일 설정: {target_date}")
        
        # 기본 설정은 자동으로 당일로 설정되므로 추가 설정이 필요하면 여기서
        time.sleep(3)

    def download_and_save(self, driver, company_name: str, target_date: str, **kwargs) -> bool:
        """다운로드 및 저장"""
        try:
            print(f"   📊 {company_name} 매출채권 엑셀 다운로드 시작...")
            
            # 엑셀 버튼 클릭
            wait = WebDriverWait(driver, 10)
            excel_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div#footer_toolbar_toolbar_item_excel button")))
            self.js_click(driver, excel_button)
            
            # 다운로드 대기
            filename = f"{company_name}_매출채권_{target_date}.xlsx"
            downloaded_file = self.wait_for_download(company_name, filename)
            
            if downloaded_file:
                # 저장 경로로 이동
                save_dir = self.config.get_receivables_raw_data_dir()
                save_dir.mkdir(parents=True, exist_ok=True)
                final_path = save_dir / filename
                
                shutil.move(str(downloaded_file), str(final_path))
                print(f"   ✅ 파일 저장 완료: {final_path}")
                return True
            else:
                print(f"   ❌ 다운로드 실패")
                return False
                
        except Exception as e:
            print(f"   ❌ 다운로드 오류: {e}")
            return False

    def collect_receivables_data(self, target_date: str = None, progress_callback=None):
        """매출채권 데이터 수집 - 금요일 기준 (리팩토링됨)"""
        print("🚀 매출채권 데이터 수집 시작 (금요일 기준, 리팩토링됨)")
        
        # 금요일 날짜 계산
        friday_date = self.get_friday_date(target_date)
        print(f"📅 수집 기준일: {friday_date} (금요일)")
        
        accounts = self.get_target_accounts()
        total_tasks = len(accounts)
        current_task = 0
        
        for account in accounts:
            company_name = account["company_name"]
            print(f"\n💼 {company_name} 매출채권 수집 시작")
            
            driver = None
            try:
                current_task += 1
                
                if progress_callback:
                    progress = (current_task / total_tasks) * 100
                    progress_callback(f"{company_name} 매출채권 수집 중", progress)
                
                driver = self.launch_driver()
                
                # 로그인
                if not self.basic_login(driver, account):
                    print(f"❌ {company_name} 로그인 실패")
                    continue
                
                # 매출채권조회 페이지로 이동
                self.navigate_to_target_page(driver, account)
                
                # 검색 조건 설정
                self.set_search_criteria(driver, friday_date, company_name)
                
                # 다운로드 및 저장
                if self.download_and_save(driver, company_name, friday_date):
                    print(f"   ✅ {company_name} 매출채권 수집 완료")
                else:
                    print(f"   ❌ {company_name} 매출채권 수집 실패")
                    
            except Exception as e:
                print(f"❌ {company_name} 매출채권 수집 실패: {e}")
                
            finally:
                if driver:
                    driver.quit()
                    print(f"🔌 {company_name} 브라우저 종료")
        
        print(f"\n🎉 매출채권 데이터 수집 완료 (금요일 기준, 리팩토링됨)")

class UnifiedDataCollector:
    """통합 데이터 수집기 - 매출과 매출채권을 통합 관리 (리팩토링됨)"""
    
    def __init__(self, headless_mode=None, months=None):
        self.config = get_config()
        self.sales_collector = SalesDataCollector(headless_mode)
        self.receivables_collector = ReceivablesDataCollector(headless_mode)
        
        # months 매개변수 처리 (GUI 호환성)
        if months is not None:
            self.default_months = months
        else:
            self.default_months = self.config.get_default_num_months()
        
    def set_headless_mode(self, headless: bool = True):
        """모든 수집기에 헤드리스 모드 설정"""
        self.sales_collector.set_headless_mode(headless)
        self.receivables_collector.set_headless_mode(headless)
        
    def collect_all_data(self, months_back: int = None, target_date: str = None, 
                        sales_only: bool = False, receivables_only: bool = False,
                        progress_callback=None):
        """모든 데이터 수집 (월~금 기준, 리팩토링됨)"""
        
        if months_back is None:
            months_back = self.default_months
            
        print(f"🚀 통합 데이터 수집 시작 (월~금 기준, 리팩토링됨)")
        print(f"📅 수집 범위: {months_back}개월 전부터")
        print(f"💰 매출채권: 금요일 기준 수집")
        
        success_results = {}
        
        # 매출 데이터 수집
        if not receivables_only:
            print(f"\n📊 1단계: 매출 데이터 수집 ({months_back}개월)")
            try:
                self.sales_collector.collect_data(num_months=months_back, progress_callback=progress_callback)
                success_results['sales'] = True
                print("✅ 매출 데이터 수집 완료")
            except Exception as e:
                print(f"❌ 매출 데이터 수집 실패: {e}")
                success_results['sales'] = False
        
        # 매출채권 데이터 수집 (금요일 기준)
        if not sales_only:
            print(f"\n💰 2단계: 매출채권 데이터 수집 (금요일 기준)")
            try:
                self.receivables_collector.collect_receivables_data(target_date=target_date, progress_callback=progress_callback)
                success_results['receivables'] = True
                print("✅ 매출채권 데이터 수집 완료 (금요일 기준)")
            except Exception as e:
                print(f"❌ 매출채권 데이터 수집 실패: {e}")
                success_results['receivables'] = False
        
        # 결과 요약
        print(f"\n📋 수집 결과 요약 (월~금 기준, 리팩토링됨):")
        for data_type, success in success_results.items():
            status = "✅ 성공" if success else "❌ 실패"
            print(f"   - {data_type}: {status}")
        
        total_success = sum(success_results.values())
        total_tasks = len(success_results)
        print(f"\n🎯 전체 성공률: {total_success}/{total_tasks}")
        
        return success_results

    def collect_data(self, progress_callback=None, **kwargs):
        """기본 데이터 수집 메서드 - GUI 호환성"""
        return self.collect_all_data(progress_callback=progress_callback, **kwargs)

    def collect_sales_data_with_dates(self, start_date: str, end_date: str, progress_callback=None):
        """특정 기간 매출 데이터 수집 (리팩토링됨)"""
        print(f"🎯 특정 기간 매출 데이터 수집: {start_date} ~ {end_date} (리팩토링됨)")
        
        try:
            self.sales_collector = SalesDataCollector()
            
            # 날짜 범위 설정
            from datetime import datetime
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            
            print(f"   📅 기간 설정: {start_dt.strftime('%Y-%m-%d')} ~ {end_dt.strftime('%Y-%m-%d')}")
            
            # 매출 데이터 수집 실행
            success = self.sales_collector.collect_data(
                start_date=start_dt.strftime('%Y%m%d'), 
                end_date=end_dt.strftime('%Y%m%d'),
                progress_callback=progress_callback
            )
            
            if success:
                print(f"✅ {start_date} ~ {end_date} 기간 매출 데이터 수집 완료")
            else:
                print(f"❌ {start_date} ~ {end_date} 기간 매출 데이터 수집 실패")
            
            return success
            
        except Exception as e:
            print(f"❌ 기간별 매출 데이터 수집 오류: {e}")
            import traceback
            traceback.print_exc()
            return False

# 이전 버전과의 호환성을 위한 클래스 별칭
DataCollector = SalesDataCollector
ReceivablesCollector = ReceivablesDataCollector

def main():
    """메인 실행 함수 (월~금 기준, 리팩토링됨)"""
    try:
        # 통합 수집기 사용
        unified_collector = UnifiedDataCollector()
        
        # 기본 설정으로 모든 데이터 수집
        print("🚀 통합 데이터 수집 시작 (리팩토링됨)")
        results = unified_collector.collect_all_data()
        
        print("\n🎉 데이터 수집 완료! (리팩토링됨)")
        return results
        
    except Exception as e:
        print(f"❌ 메인 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    main()
