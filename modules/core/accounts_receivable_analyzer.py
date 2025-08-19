import pandas as pd
import os
from pathlib import Path
from datetime import datetime, timedelta
import sys
import re
import logging

# 리팩토링된 경로 설정
sys.path.append(str(Path(__file__).parent.parent))

# 설정 관리자 import (새 구조)
from modules.utils.config_manager import get_config

class ReceivablesAnalyzer:
    """매출채권 분석 엔진 클래스 - 리팩토링된 버전"""
    
    def __init__(self):
        self.config = get_config()
        self.logger = logging.getLogger('ReceivablesAnalyzer')
    
    def find_best_file_for_week(self, files_with_dates, target_date):
        """주차에 맞는 최적 파일 찾기"""
        week_start = target_date - timedelta(days=target_date.weekday())
        week_end = week_start + timedelta(days=6)
        
        # 해당 주 내 파일들 찾기
        week_files = [
            (file_path, file_date) for file_path, file_date in files_with_dates
            if week_start <= file_date <= week_end
        ]
        
        if week_files:
            # 해당 주에서 가장 늦은 파일 선택
            return max(week_files, key=lambda x: x[1])
        
        # 해당 주에 파일이 없으면 가장 가까운 파일 찾기
        if not files_with_dates:
            return None
        
        closest_file = min(files_with_dates, key=lambda x: abs((x[1] - target_date).days))
        return closest_file


class AccountsReceivableAnalyzer:
    """매출채권 분석 클래스 - 월~금 기준 (리팩토링됨)"""
    
    def __init__(self):
        self.config = get_config()
        # 통합 로거 사용
        self.logger = logging.getLogger('AccountsReceivableAnalyzer')
        
    def read_data(self, file_path):
        """엑셀 파일에서 매출채권 데이터 읽기"""
        df_all = pd.DataFrame()
        
        try:
            xl = pd.ExcelFile(file_path)
            for sheet in xl.sheet_names:
                if "디앤드디" in sheet or "디앤아이" in sheet:
                    company = "디앤드디" if "디앤드디" in sheet else "디앤아이"
                    df = xl.parse(sheet)
                    df["회사"] = company
                    
                    # 합계 행 제외 (거래처명이 "합계"인 행)
                    if "거래처명" in df.columns:
                        before_count = len(df)
                        df = df[df["거래처명"] != "합계"]
                        after_count = len(df)
                        if before_count != after_count:
                            self.logger.debug(f"{company}: 합계 행 {before_count - after_count}개 제외됨")
                    
                    # 거래처코드가 비어있거나 숫자가 아닌 행도 제외 (추가 안전장치)
                    if "거래처코드" in df.columns:
                        before_count = len(df)
                        df = df[pd.to_numeric(df["거래처코드"], errors='coerce').notna()]
                        after_count = len(df)
                        if before_count != after_count:
                            self.logger.debug(f"{company}: 유효하지 않은 거래처코드 행 {before_count - after_count}개 제외됨")
                    
                    df_all = pd.concat([df_all, df], ignore_index=True)
                    
            print(f"  📄 데이터 로드: {len(df_all)}행")
            self.logger.info(f"데이터 로드 완료: {file_path}")
            self.logger.info(f"총 {len(df_all)}행 로드됨 (합계 행 제외)")
            return df_all
            
        except Exception as e:
            error_msg = f"파일 읽기 실패 {file_path}: {e}"
            print(f"  ❌ {error_msg}")
            self.logger.error(error_msg)
            return pd.DataFrame()

    def safe_round(self, value, decimals=4):
        """안전한 반올림 함수"""
        try:
            if pd.isna(value):
                return 0.0
            return round(float(value), decimals)
        except:
            return 0.0

    def safe_divide(self, numerator, denominator, decimals=4):
        """안전한 나눗셈 함수"""
        try:
            if denominator == 0 or pd.isna(denominator) or pd.isna(numerator):
                return 0.0
            result = float(numerator) / float(denominator)
            return self.safe_round(result, decimals)
        except:
            return 0.0

    def extract_date_from_filename(self, filename):
        """파일명에서 날짜 추출 (다양한 형식 지원)"""
        from datetime import datetime
        
        # 패턴 1: 매출채권계산결과YYYYMMDD.xlsx
        match1 = re.search(r'매출채권계산결과(\d{8})\.xlsx', str(filename))
        if match1:
            try:
                date_str = match1.group(1)
                return datetime.strptime(date_str, '%Y%m%d').date()
            except ValueError:
                pass
        
        # 패턴 2: 매출채권계산결과(YYYY-MM-DD).xlsx  
        match2 = re.search(r'매출채권계산결과\((\d{4}-\d{2}-\d{2})\)\.xlsx', str(filename))
        if match2:
            try:
                date_str = match2.group(1)
                return datetime.strptime(date_str, '%Y-%m-%d').date()
            except ValueError:
                pass
        
        # 패턴 3: 매출채권계산결과YYYY-MM-DD.xlsx (하이픈 형식)
        match3 = re.search(r'매출채권계산결과(\d{4}-\d{2}-\d{2})\.xlsx', str(filename))
        if match3:
            try:
                date_str = match3.group(1)
                return datetime.strptime(date_str, '%Y-%m-%d').date()
            except ValueError:
                pass
                
        return None

    def get_week_start_monday(self, date):
        """주어진 날짜가 속한 주의 월요일 반환 (월~금 기준으로 변경)"""
        from datetime import timedelta
        
        # Python에서 월요일=0, 화요일=1, ..., 일요일=6
        weekday = date.weekday()
        
        # 월~금을 한 주기로 보는 방식으로 변경
        if weekday <= 4:  # 월~금 (0,1,2,3,4)
            # 이번 주 월요일
            days_to_monday = weekday
            monday = date - timedelta(days=days_to_monday)
        else:  # 토~일 (5,6)
            # 다음 주 월요일
            days_to_next_monday = 7 - weekday
            monday = date + timedelta(days=days_to_next_monday)
        
        return monday

    def classify_week_by_date(self, extract_date, reference_date=None):
        """추출일을 기준으로 주차 분류 (월~금 기준으로 변경)"""
        if reference_date is None:
            reference_date = datetime.now().date()
        
        # 기준일이 속한 주의 월요일
        reference_monday = self.get_week_start_monday(reference_date)
        
        # 추출일이 속한 주의 월요일  
        extract_monday = self.get_week_start_monday(extract_date)
        
        # 주차 차이 계산
        week_diff = (reference_monday - extract_monday).days // 7
        
        if week_diff == 0:
            return "이번주"
        elif week_diff == 1:
            return "전주"
        else:
            return f"{week_diff}주전" if week_diff > 1 else f"{abs(week_diff)}주후"

    def find_latest_files_by_week(self, reference_date=None):
        """주간 기준으로 최신 파일들 찾기 (월~금 기준, 파일명 날짜 기준)"""
        receivable_dir = self.config.get_receivable_raw_data_dir()
        
        # 매출채권 파일 찾기
        receivable_files = list(receivable_dir.glob("매출채권계산결과*.xlsx"))
        if not receivable_files:
            error_msg = "매출채권 파일을 찾을 수 없습니다."
            print(f"  ❌ {error_msg}")
            self.logger.error(error_msg)
            return None, None
        
        self.logger.info(f"발견된 매출채권 파일: {len(receivable_files)}개")
        
        # 파일명에서 날짜를 추출하여 주차별로 분류
        files_by_week = {"이번주": [], "전주": [], "기타": []}
        
        for file_path in receivable_files:
            file_date = self.extract_date_from_filename(file_path.name)
            if file_date:
                week_category = self.classify_week_by_date(file_date, reference_date)
                
                if week_category in files_by_week:
                    files_by_week[week_category].append((file_path, file_date))
                else:
                    files_by_week["기타"].append((file_path, file_date))
            else:
                self.logger.warning(f"파일명에서 날짜 추출 실패: {file_path.name}")
                files_by_week["기타"].append((file_path, None))

        # 각 주차별로 가장 최신 파일 선택
        curr_file_path = None
        prev_file_path = None
        
        # 이번주 파일 중 최신 파일
        if files_by_week["이번주"]:
            curr_file_path, curr_date = max(files_by_week["이번주"], key=lambda x: x[1] if x[1] else datetime.min.date())
            self.logger.info(f"이번주 파일: {curr_file_path.name} (날짜: {curr_date})")
        
        # 전주 파일 중 최신 파일
        if files_by_week["전주"]:
            prev_file_path, prev_date = max(files_by_week["전주"], key=lambda x: x[1] if x[1] else datetime.min.date())
            self.logger.info(f"전주 파일: {prev_file_path.name} (날짜: {prev_date})")
        
        # 파일이 없는 경우 처리
        if curr_file_path is None:
            self.logger.warning("이번주 파일이 없습니다.")
            
            # 모든 유효한 파일 중 가장 최신 파일을 현재 주로 사용
            all_valid_files = []
            for week_files in files_by_week.values():
                all_valid_files.extend([f for f in week_files if f[1] is not None])
            
            if all_valid_files:
                curr_file_path, curr_date = max(all_valid_files, key=lambda x: x[1])
                curr_week_label = self.classify_week_by_date(curr_date) if curr_date else "알수없음"
                self.logger.info(f"대체 현재 파일: {curr_file_path.name} (날짜: {curr_date}, 실제주차: {curr_week_label})")
                
                # 전주 파일 재선택
                remaining_files = [f for f in all_valid_files if f[0] != curr_file_path]
                if remaining_files:
                    if prev_file_path is None or prev_file_path == curr_file_path:
                        prev_file_path, prev_date = max(remaining_files, key=lambda x: x[1])
                        prev_week_label = self.classify_week_by_date(prev_date) if prev_date else "알수없음"
                        self.logger.info(f"대체 전주 파일: {prev_file_path.name} (날짜: {prev_date}, 실제주차: {prev_week_label})")

        # 파일이 같은 경우 처리
        if curr_file_path and prev_file_path and curr_file_path == prev_file_path:
            self.logger.warning("현재 주와 전주 파일이 동일합니다. 전주 파일을 다시 선택합니다.")
            
            all_valid_files = []
            for week_files in files_by_week.values():
                all_valid_files.extend([f for f in week_files if f[1] is not None])
            
            remaining_files = [f for f in all_valid_files if f[0] != curr_file_path]
            if remaining_files:
                prev_file_path, prev_date = max(remaining_files, key=lambda x: x[1])
                self.logger.info(f"수정된 전주 파일: {prev_file_path.name} (날짜: {prev_date})")
            else:
                prev_file_path = None
                self.logger.info("전주 파일 없음 - 단일 파일로 분석")

        if prev_file_path is None and curr_file_path is not None:
            self.logger.warning("전주 파일이 없어 비교 분석을 건너뜁니다.")
        
        # 간단한 콘솔 출력
        if curr_file_path:
            print(f"  📄 현재 주 (월~금): {curr_file_path.name}")
        if prev_file_path:
            print(f"  📄 전주 (월~금): {prev_file_path.name}")
        
        return curr_file_path, prev_file_path

    def find_latest_files(self):
        """최신 파일들 자동 찾기 - 주간 기준 방식 사용 (월~금)"""
        return self.find_latest_files_by_week()

    def summarize_receivables(self, df):
        """매출채권 요약 분석"""
        if df.empty:
            return pd.DataFrame()
            
        self.logger.debug(f"입력 데이터 컬럼: {df.columns.tolist()}")
        
        # 숫자형 변환
        numeric_columns = ["총채권", "기간초과 매출채권", "90일초과 매출채권"]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                self.logger.warning(f"컬럼 '{col}'이 없습니다. 0으로 설정합니다.")
                df[col] = 0

        # 회사별 집계
        summary = df.groupby("회사").agg({
            "총채권": "sum",
            "90일초과 매출채권": "sum",
            "기간초과 매출채권": "sum"
        }).reset_index()

        # 전체 합계 추가
        total_row = pd.DataFrame({
            "회사": ["합계"],
            "총채권": [summary["총채권"].sum()],
            "90일초과 매출채권": [summary["90일초과 매출채권"].sum()],
            "기간초과 매출채권": [summary["기간초과 매출채권"].sum()]
        })

        final = pd.concat([summary, total_row], ignore_index=True)
        
        # 비율 계산을 안전하게 수행
        final["90일비율"] = 0.0
        final["기간초과비율"] = 0.0
        
        for idx, row in final.iterrows():
            final.at[idx, "90일비율"] = self.safe_divide(row["90일초과 매출채권"], row["총채권"])
            final.at[idx, "기간초과비율"] = self.safe_divide(row["기간초과 매출채권"], row["총채권"])
        
        return final

    def make_comparison(self, curr_summary, prev_summary):
        """전주 vs 금주 비교 분석 (계산 결과 시트용 - 원래 형태로 복원)"""
        if curr_summary.empty or prev_summary.empty:
            return curr_summary.copy() if not curr_summary.empty else pd.DataFrame()
            
        merged = curr_summary.merge(prev_summary, on="회사", suffixes=("_curr", "_prev"))
        
        result = pd.DataFrame()
        result["항목"] = merged["회사"]
        result["총채권(전주)"] = merged["총채권_prev"]
        result["총채권(금주)"] = merged["총채권_curr"]
        result["총채권(증감)"] = merged["총채권_curr"] - merged["총채권_prev"]
        result["장기미수채권90일(전주)"] = merged["90일초과 매출채권_prev"]
        result["장기미수채권90일(금주)"] = merged["90일초과 매출채권_curr"]
        result["장기미수채권90일(증감)"] = merged["90일초과 매출채권_curr"] - merged["90일초과 매출채권_prev"]
        result["90일비율(금주)"] = merged["90일비율_curr"]
        result["기간초과채권(금주)"] = merged["기간초과 매출채권_curr"]
        result["기간초과비율(금주)"] = merged["기간초과비율_curr"]
        
        return result

    def make_summary_pivot(self, curr_summary, prev_summary=None):
        """요약 피벗 테이블 생성 (파워포인트용, 새로운 컬럼 구조)"""
        if curr_summary.empty:
            return pd.DataFrame()
        
        # 1. 피벗 테이블 생성
        pivot_data = []
        
        for _, row in curr_summary.iterrows():
            company = row["회사"]
            if company == "디앤드디":
                display_name = "DND"
            elif company == "디앤아이":
                display_name = "DNI"
            else:
                display_name = company
            
            # 기본 데이터
            pivot_row = {
                "항목": display_name,
                "총채권": round(row["총채권"] / 1000000, 0),  # 백만원 단위
                "총채권 증감(%)": 0.0,  # 기본값
                "90일 채권 (100만)": round(row["90일초과 매출채권"] / 1000000, 1),
                "90일 채권 증감(%)": 0.0,  # 기본값
                "90일 총채권대비(%)": round(row["90일비율"] * 100, 1),
                "90일 증감(%p)": 0.0,  # 기본값
                "결제예정일 초과채권 (100만)": round(row["기간초과 매출채권"] / 1000000, 1),
                "결제예정일 초과채권 증감(%)": 0.0,  # 기본값
                "결제예정일 총채권대비(%)": round(row["기간초과비율"] * 100, 1),  # 결제예정일 초과 비율
                "결제예정일 초과증감(%p)": 0.0  # 기본값
            }
            
            # 전주 대비 증감률 계산
            if prev_summary is not None and not prev_summary.empty:
                prev_row = prev_summary[prev_summary["회사"] == company]
                if not prev_row.empty:
                    prev_data = prev_row.iloc[0]
                    
                    # 총채권 증감률 (%)
                    if prev_data["총채권"] != 0:
                        change_rate = (row["총채권"] - prev_data["총채권"]) / prev_data["총채권"] * 100
                        pivot_row["총채권 증감(%)"] = round(change_rate, 1)
                    
                    # 90일 채권 증감률 (%)
                    if prev_data["90일초과 매출채권"] != 0:
                        change_rate = (row["90일초과 매출채권"] - prev_data["90일초과 매출채권"]) / prev_data["90일초과 매출채권"] * 100
                        pivot_row["90일 채권 증감(%)"] = round(change_rate, 1)
                    
                    # 90일 초과 비율 증감 (%p)
                    prev_90_ratio = prev_data["90일비율"] * 100
                    curr_90_ratio = row["90일비율"] * 100
                    change_90 = curr_90_ratio - prev_90_ratio
                    pivot_row["90일 증감(%p)"] = round(change_90, 1)
                    
                    # 결제예정일 초과채권 증감률 (%)
                    if prev_data["기간초과 매출채권"] != 0:
                        change_rate = (row["기간초과 매출채권"] - prev_data["기간초과 매출채권"]) / prev_data["기간초과 매출채권"] * 100
                        pivot_row["결제예정일 초과채권 증감(%)"] = round(change_rate, 1)
                    
                    # 결제예정일 초과 비율 증감 (%p)
                    prev_overdue_ratio = prev_data["기간초과비율"] * 100
                    curr_overdue_ratio = row["기간초과비율"] * 100
                    change_overdue = curr_overdue_ratio - prev_overdue_ratio
                    pivot_row["결제예정일 초과증감(%p)"] = round(change_overdue, 1)
            
            pivot_data.append(pivot_row)
        
        pivot_df = pd.DataFrame(pivot_data)
        
        return pivot_df

    def make_top20_clients(self, curr_df, prev_df):
        """상위 20개 기간초과 채권 거래처 분석"""
        if curr_df.empty:
            return pd.DataFrame()
        
        # 거래처명 컬럼 찾기
        client_cols = [col for col in curr_df.columns if "거래처" in col and "명" in col]
        if not client_cols:
            self.logger.warning("거래처명 컬럼을 찾을 수 없습니다.")
            self.logger.debug(f"사용 가능한 컬럼: {curr_df.columns.tolist()}")
            return pd.DataFrame()
        
        client_col = client_cols[0]
        self.logger.debug(f"거래처 컬럼 사용: {client_col}")
            
        # 현재 주 거래처별 집계
        try:
            curr_agg = curr_df.groupby(client_col, as_index=False).agg({
                "총채권": "sum",
                "기간초과 매출채권": "sum"
            })
            curr_agg = curr_agg.rename(columns={
                "총채권": "총채권_금주", 
                "기간초과 매출채권": "기간초과_금주",
                client_col: "거래처명"
            })
        except Exception as e:
            self.logger.error(f"현재 주 집계 실패: {e}")
            return pd.DataFrame()

        # 전주 데이터 처리
        if not prev_df.empty and client_col in prev_df.columns:
            try:
                prev_agg = prev_df.groupby(client_col, as_index=False).agg({
                    "총채권": "sum",
                    "기간초과 매출채권": "sum"
                })
                prev_agg = prev_agg.rename(columns={
                    "총채권": "총채권_전주", 
                    "기간초과 매출채권": "기간초과_전주",
                    client_col: "거래처명"
                })
                
                merged = curr_agg.merge(prev_agg, on="거래처명", how="left")
                merged["총채권_전주"] = merged["총채권_전주"].fillna(0)
                merged["기간초과_전주"] = merged["기간초과_전주"].fillna(0)
            except Exception as e:
                self.logger.warning(f"전주 데이터 병합 실패: {e}")
                merged = curr_agg.copy()
                merged["총채권_전주"] = 0
                merged["기간초과_전주"] = 0
        else:
            merged = curr_agg.copy()
            merged["총채권_전주"] = 0
            merged["기간초과_전주"] = 0

        # 비율 및 증감율 계산을 안전하게 수행
        merged["결제예정일초과비율_금주"] = 0.0
        merged["결제예정일초과비율_전주"] = 0.0
        merged["전주대비채권증감율"] = 0.0
        merged["전주대비결제예정일초과증감율"] = 0.0
        
        for idx, row in merged.iterrows():
            # 결제예정일 초과비율 (현재주, 전주)
            merged.at[idx, "결제예정일초과비율_금주"] = self.safe_divide(row["기간초과_금주"], row["총채권_금주"]) * 100
            merged.at[idx, "결제예정일초과비율_전주"] = self.safe_divide(row["기간초과_전주"], row["총채권_전주"]) * 100
            
            # 전주 대비 채권 증감율 (%)
            merged.at[idx, "전주대비채권증감율"] = self.safe_divide(
                row["총채권_금주"] - row["총채권_전주"], 
                row["총채권_전주"]
            ) * 100
            
            # 전주 대비 결제예정일 초과비율 증감 (%p) - 비율의 차이
            merged.at[idx, "전주대비결제예정일초과증감율"] = (
                merged.at[idx, "결제예정일초과비율_금주"] - merged.at[idx, "결제예정일초과비율_전주"]
            )

        # 상위 20개 선택
        top20 = merged.sort_values(by="기간초과_금주", ascending=False).head(20)
        
        # 컬럼 선택 및 이름 변경 (백만원 단위로 변환)
        result = pd.DataFrame()
        result["거래처명"] = top20["거래처명"]
        result["총채권(백만)"] = round(top20["총채권_금주"] / 1000000, 1)
        result["결제예정일초과(백만)"] = round(top20["기간초과_금주"] / 1000000, 1)
        result["결제예정일초과비율(%)"] = round(top20["결제예정일초과비율_금주"], 1)
        result["전주대비채권증감율(%)"] = round(top20["전주대비채권증감율"], 1)
        result["전주대비결제예정일초과증감율(%p)"] = round(top20["전주대비결제예정일초과증감율"], 1)
        
        return result

    def create_file_info_sheet(self, curr_file_path, prev_file_path):
        """파일 정보 시트 생성 (월~금 기준)"""
        from datetime import datetime
        
        file_info_data = []
        
        # 현재 주 파일 정보
        if curr_file_path:
            curr_filename = Path(curr_file_path).name
            curr_date = self.extract_date_from_filename(curr_filename)
            curr_week_label = self.classify_week_by_date(curr_date) if curr_date else "알수없음"
            
            file_info_data.append({
                "구분": "현재 주 (월~금)",
                "파일명": curr_filename,
                "추출일": curr_date.strftime("%Y-%m-%d") if curr_date else "알수없음",
                "주차분류": curr_week_label,
                "파일경로": str(curr_file_path)
            })
        
        # 전주 파일 정보
        if prev_file_path:
            prev_filename = Path(prev_file_path).name
            prev_date = self.extract_date_from_filename(prev_filename)
            prev_week_label = self.classify_week_by_date(prev_date) if prev_date else "알수없음"
            
            file_info_data.append({
                "구분": "전주 (월~금)",
                "파일명": prev_filename,
                "추출일": prev_date.strftime("%Y-%m-%d") if prev_date else "알수없음",
                "주차분류": prev_week_label,
                "파일경로": str(prev_file_path)
            })
        
        # 분석 실행 정보
        file_info_data.append({
            "구분": "분석실행정보",
            "파일명": f"분석일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "추출일": f"기준주차: 월요일~금요일",      # 변경됨
            "주차분류": f"생성파일: 채권_분석_결과.xlsx",
            "파일경로": ""
        })
        
        return pd.DataFrame(file_info_data)

    def analyze_receivables(self, prev_file_path=None, curr_file_path=None, output_filename="채권_분석_결과.xlsx"):
        """매출채권 전체 분석 프로세스 (월~금 기준, 리팩토링됨)"""
        
        self.logger.info("=== 매출채권 분석 시작 (월~금 기준, 리팩토링됨) ===")
        
        # 파일 경로 결정
        if curr_file_path is None or prev_file_path is None:
            self.logger.info("파일 경로가 지정되지 않음. 주간 기준으로 최신 파일 찾는 중... (월~금)")
            auto_curr, auto_prev = self.find_latest_files_by_week()
            
            if curr_file_path is None:
                curr_file_path = auto_curr
            if prev_file_path is None:
                prev_file_path = auto_prev
        
        if curr_file_path:
            self.logger.info(f"현재 주 파일 (월~금): {curr_file_path}")
        if prev_file_path:
            self.logger.info(f"전주 파일 (월~금): {prev_file_path}")

        # 데이터 로드
        if curr_file_path is None:
            error_msg = "현재 주 파일을 찾을 수 없습니다."
            print(f"  ❌ {error_msg}")
            self.logger.error(error_msg)
            return None
            
        curr_df = self.read_data(curr_file_path)
        prev_df = self.read_data(prev_file_path) if prev_file_path else pd.DataFrame()

        if curr_df.empty:
            error_msg = "현재 주 데이터가 없습니다."
            print(f"  ❌ {error_msg}")
            self.logger.error(error_msg)
            return None

        # 분석 수행
        try:
            summary_curr = self.summarize_receivables(curr_df)
            summary_prev = self.summarize_receivables(prev_df) if not prev_df.empty else pd.DataFrame()
            
            # 비교 분석 (계산 결과 시트용)
            summary_combined = self.make_comparison(summary_curr, summary_prev)
            
            # 요약 피벗 테이블 (파워포인트용)
            pivot_summary = self.make_summary_pivot(summary_curr, summary_prev)
                
            # TOP20 분석
            top20 = self.make_top20_clients(curr_df, prev_df)
            
            # 파일 정보 시트 생성
            file_info_sheet = self.create_file_info_sheet(curr_file_path, prev_file_path)

            # 결과 저장
            output_path = self.config.get_processed_data_dir() / output_filename
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                # 시트들 저장
                file_info_sheet.to_excel(writer, sheet_name="파일정보", index=False)
                
                if not pivot_summary.empty:
                    pivot_summary.to_excel(writer, sheet_name="요약", index=False)
                
                if not summary_combined.empty:
                    summary_combined.to_excel(writer, sheet_name="계산 결과", index=False)
                
                if not top20.empty:
                    top20.to_excel(writer, sheet_name="TOP20_금주", index=False)
                
                curr_df.to_excel(writer, sheet_name="원본_금주", index=False)
                if not prev_df.empty:
                    prev_df.to_excel(writer, sheet_name="원본_전주", index=False)
                
            print(f"  💾 결과 저장: {output_filename}")
            self.logger.info(f"채권 분석 결과 저장 완료 (월~금 기준): {output_path}")
            
            # 간단한 요약 출력
            if not summary_combined.empty:
                self.logger.info("매출채권 분석 완료 (월~금 기준, 리팩토링됨)")
                self.logger.debug(f"분석 결과:\n{summary_combined.to_string(index=False)}")
                
                # KPI 체크
                kpi_target = self.config.get_kpi_target("장기미수채권_비율")
                total_rows = summary_combined[summary_combined["항목"] == "합계"]
                if not total_rows.empty and "90일비율(금주)" in summary_combined.columns:
                    current_ratio = total_rows["90일비율(금주)"].iloc[0] * 100
                    self.logger.info(f"KPI 체크: 현재 {current_ratio:.2f}% (목표: {kpi_target}%)")
                    if current_ratio > kpi_target:
                        self.logger.warning("KPI 기준 초과")
                    else:
                        self.logger.info("KPI 기준 달성")

            return {
                "file_info": file_info_sheet,
                "pivot_summary": pivot_summary,
                "calculation_result": summary_combined,
                "top20": top20,
                "curr_data": curr_df,
                "prev_data": prev_df
            }
            
        except Exception as e:
            error_msg = f"분석 중 오류 발생: {e}"
            print(f"  ❌ {error_msg}")
            self.logger.error(error_msg, exc_info=True)
            return None


def main(prev_file=None, curr_file=None):
    """메인 실행 함수 (월~금 기준, 리팩토링됨)"""
    try:
        analyzer = AccountsReceivableAnalyzer()
        
        # 파일 경로를 Path 객체로 변환 (문자열인 경우)
        if prev_file and isinstance(prev_file, str):
            prev_file = Path(prev_file) if prev_file != "None" else None
        if curr_file and isinstance(curr_file, str):
            curr_file = Path(curr_file) if curr_file != "None" else None
            
        results = analyzer.analyze_receivables(prev_file, curr_file)
        
        if results:
            print("🎉 매출채권 분석 완료! (월~금 기준, 리팩토링됨)")
            return results
        else:
            print("❌ 분석 실패")
            return None
            
    except Exception as e:
        print(f"❌ 전체 프로세스 오류: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    # 직접 실행시 파일 경로 지정 가능
    import sys
    
    prev_file = sys.argv[1] if len(sys.argv) > 1 else None
    curr_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    main(prev_file, curr_file)
