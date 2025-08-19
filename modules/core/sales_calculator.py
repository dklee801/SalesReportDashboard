import pandas as pd
import os
from glob import glob
from pathlib import Path
from datetime import datetime, timedelta
import sys
import logging

# Excel 복구를 위한 추가 import
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    logging.warning("openpyxl을 찾을 수 없습니다. Excel 복구 기능이 제한됩니다.")

# 리팩토링된 경로 설정
sys.path.append(str(Path(__file__).parent.parent))

# 설정 관리자 import (새 구조)
from modules.utils.config_manager import get_config

# 백업 관리자 import (새 구조)
try:
    from modules.utils.backup_manager import BackupManager
    BACKUP_AVAILABLE = True
except ImportError:
    BACKUP_AVAILABLE = False
    BackupManager = None

class SalesCalculator:
    """매출 데이터 정제 및 분석 클래스 - 금~목 기준 (리팩토링됨)"""
    
    def __init__(self):
        self.config = get_config()
        
        # 로깅 설정 - 먼저 로거 설정
        self.logger = logging.getLogger(__name__)
        if not self.logger.handlers:
            self.logger.setLevel(logging.INFO)
        
        # 백업 관리자 초기화
        if BACKUP_AVAILABLE:
            self.backup_manager = BackupManager(backup_retention_days=7)  # 수집 데이터는 7일간 보관
            self.logger.info("✅ 백업 관리자 활성화")
        else:
            self.backup_manager = None
            
        # 이제 staff_df와 exclude_codes, exclude_products 로드 (로거가 준비된 후)
        self.staff_df = self.load_staff_info()
        self.exclude_codes = [str(int(float(code))) for code in self.config.get_exclude_codes()]
        self.exclude_products = self.config.report_config.get("sales", {}).get("exclude_products", [])
        
    def load_staff_info(self):
        """담당자 정보 로드"""
        try:
            staff_file_path = self.config.get_staff_file_path()
            sheet_name = self.config.get_staff_sheet_name()
            
            if staff_file_path.exists():
                staff_df = pd.read_excel(staff_file_path, sheet_name=sheet_name)
                # 로거가 아직 없을 수 있으므로 조건부 로깅
                if hasattr(self, 'logger'):
                    self.logger.info(f"담당자 정보 로드 완료: {len(staff_df)}행")
                return staff_df
            else:
                if hasattr(self, 'logger'):
                    self.logger.warning(f"담당자 파일이 없습니다: {staff_file_path}")
                return pd.DataFrame()
        except Exception as e:
            if hasattr(self, 'logger'):
                self.logger.error(f"담당자 정보 로드 실패: {e}")
            return pd.DataFrame()

    def repair_excel_with_openpyxl(self, file_path):
        """openpyxl을 사용하여 손상된 Excel 파일 복구"""
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl이 설치되지 않았습니다")
            
        try:
            self.logger.info(f"📧 openpyxl로 Excel 복구 시도: {os.path.basename(file_path)}")
            
            # openpyxl로 파일 열기 (data_only=True로 수식 대신 값만 가져오기)
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            
            # 모든 데이터를 리스트로 변환
            data = []
            for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if row_idx == 0:  # skiprows=1 효과 (첫 번째 행 건너뛰기)
                    continue
                if any(cell is not None for cell in row):  # 빈 행이 아닌 경우만 추가
                    data.append(row)
            
            if not data:
                self.logger.warning("복구된 파일에 데이터가 없습니다")
                return pd.DataFrame()
            
            # 헤더 추출 (첫 번째 데이터 행이 헤더)
            headers = data[0] if data else []
            actual_data = data[1:] if len(data) > 1 else []
            
            # DataFrame 생성
            df = pd.DataFrame(actual_data, columns=headers)
            
            self.logger.info(f"✅ Excel 복구 성공: {len(df)}행, {len(df.columns)}열")
            return df
            
        except Exception as e:
            self.logger.error(f"openpyxl 복구 실패: {e}")
            raise

    def safe_excel_read(self, file_path, skiprows=1):
        """안전한 Excel 파일 읽기 (손상 파일 자동 복구)"""
        try:
            # 표준 pandas 시도
            return pd.read_excel(file_path, skiprows=skiprows)
        except Exception as e:
            error_msg = str(e).lower()
            if "stylesheet" in error_msg or "styles.xml" in error_msg:
                self.logger.warning(f"🔧 Stylesheet 오류 감지 - 복구 모드로 전환: {os.path.basename(file_path)}")
                if OPENPYXL_AVAILABLE:
                    return self.repair_excel_with_openpyxl(file_path)
                else:
                    self.logger.error("openpyxl이 없어 복구할 수 없습니다")
                    raise
            else:
                # 다른 오류는 그대로 전파
                raise

    def load_and_standardize(self, file_path, company_name, default_category=None):
        """파일 로드 및 표준화"""
        file_name = os.path.basename(file_path)
        self.logger.info(f"파일 처리 시작: {file_name}")
        
        try:
            df = self.safe_excel_read(file_path, skiprows=1)
        except Exception as e:
            self.logger.error(f"파일 로드 실패 {file_name}: {e}")
            return pd.DataFrame()
            
        self.logger.debug(f"Raw rows: {len(df)}")
        
        # 원본 금액 컬럼 총합 확인
        if "공급가액합계" in df.columns:
            original_total = df["공급가액합계"].sum()
            self.logger.debug(f"원본 공급가액합계 총합: {original_total:,}")
        
        # 컬럼명 정리
        df.columns = [str(c).strip() for c in df.columns]

        # 날짜 컬럼 찾기
        possible_date_cols = [col for col in df.columns if "일자" in col.replace(" ", "").replace("-", "").replace("_", "")]
        
        date_col = possible_date_cols[0] if possible_date_cols else None
        if date_col is None:
            self.logger.error(f"날짜 열을 찾을 수 없습니다: {file_name}")
            return pd.DataFrame()

        # 컬럼 표준화
        df = df.rename(columns={
            date_col: "일자원본",
            "거래처명": "client",
            "거래처코드": "client_code",
            "품목명": "product",
            "공급가액합계": "amount",
            "담당자코드": "manager"
        })

        df["company"] = company_name
        
        # 날짜 추출 및 변환
        date_extracted = df["일자원본"].astype(str).str.extract(r"^(\d{4}/\d{2}/\d{2})")
        df["date"] = pd.to_datetime(date_extracted[0], format="%Y/%m/%d", errors='coerce')
        
        # 금액 변환
        df["amount"] = pd.to_numeric(df["amount"], errors='coerce')
        
        # 결측치 제거
        before_dropna = len(df)
        df = df.dropna(subset=["date", "amount"])
        after_dropna = len(df)
        
        if before_dropna != after_dropna:
            self.logger.debug(f"결측치 제거: {before_dropna - after_dropna}행")
        
        # 거래처 코드 정리
        df["client_code"] = df["client_code"].apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
        
        # 기본 카테고리 설정
        if default_category:
            df["category"] = default_category
            
        # 반환할 컬럼 선택
        columns = ["company", "date", "client", "client_code", "product", "amount", "manager"]
        if default_category:
            columns.append("category")
        
        final_df = df[columns]
        self.logger.info(f"파일 처리 완료: {file_name} - {len(final_df):,}행 (금액: {final_df['amount'].sum():,.0f}원)")
        
        return final_df

    def categorize_and_filter(self, df, company):
        """카테고리 분류 및 필터링"""
        if company == "디앤드디":
            self.logger.debug(f"담당자 정보 병합: {company}")
            
            if not self.staff_df.empty:
                df = df.merge(self.staff_df, left_on="manager", right_on="사원번호", how="left")
                df["category"] = df["구분"]
                df.drop(columns=["사원번호", "구분"], inplace=True)

            # 카테고리 매핑 적용
            category_mappings = self.config.get_category_mappings()
            df["category"] = df["category"].replace(category_mappings)

            # 품목명 기반 제외 (새로 추가)
            if self.exclude_products:
                before_product_filter = len(df)
                # 품목명이 제외 목록에 포함된 항목 제거
                df = df[~df["product"].isin(self.exclude_products)]
                product_filtered_count = before_product_filter - len(df)
                
                if product_filtered_count > 0:
                    self.logger.info(f"품목명 필터 적용: {product_filtered_count}행 제거 (제외 품목: {self.exclude_products})")
                else:
                    self.logger.info("품목명 필터링: 제외될 항목 없음")
            else:
                self.logger.info("품목명 필터링 비활성화 (제외 품목 없음)")

            # 무역 필터링 적용 (현재 비활성화 - 제외 코드 목록이 비어있음)
            if "무역" in df["category"].values and self.exclude_codes:
                before_filter = len(df)
                df = df[~((df["category"] == "무역") & (df["client_code"].isin(self.exclude_codes)))]
                filtered_count = before_filter - len(df)
                
                if filtered_count > 0:
                    self.logger.info(f"무역 필터 적용: {filtered_count}행 제거")
            else:
                self.logger.info("무역 필터링 건너뜀 (제외 코드 없음)")

        return df

    def get_week_range(self, date):
        """주차 범위 계산 (금요일 기준)"""
        weekday = date.weekday()  # 월요일=0, 금요일=4
        
        # 금요일부터 목요일까지 (금~목 기준)
        if weekday >= 4:  # 금~일
            days_since_friday = weekday - 4
            week_start = date - timedelta(days=days_since_friday)
        else:  # 월~목
            days_to_last_friday = weekday + 3  # 지난 주 금요일까지 일수
            week_start = date - timedelta(days=days_to_last_friday)
        
        week_end = week_start + timedelta(days=6)  # 목요일 (금요일 + 6일)
        
        return week_start, week_end

    def enrich_with_time_columns(self, df):
        """시간 관련 컬럼 추가 (금~목 기준)"""
        df["week_start"], df["week_end"] = zip(*df["date"].map(self.get_week_range))
        df["기간"] = df["week_start"].dt.strftime("%Y-%m-%d") + " - " + df["week_end"].dt.strftime("%Y-%m-%d")
        df["year"] = df["date"].dt.year
        df["month"] = df["date"].dt.month
        df["week"] = df["week_start"].dt.strftime("%Y-%m-%d")
        
        # 주차 연도: 주차 끝날 기준 (목요일)
        df["week_year"] = df["week_end"].dt.year
        
        # 연월 컬럼 추가 (YYYY-MM 형식)
        df["year_month"] = df["date"].dt.strftime("%Y-%m")
        
        return df

    def summarize_monthly_data(self, df):
        """월별 데이터 집계"""
        grouped = df.groupby(["year", "month", "category"])["amount"].sum().reset_index()
        return grouped

    def summarize_weekly_data(self, df):
        """주차별 데이터 집계 (금~목 기준)"""
        grouped = df.groupby(["기간", "category"])["amount"].sum().reset_index()
        return grouped

    def summarize_client_monthly_data(self, df):
        """거래처별 월별 데이터 집계"""
        grouped = df.groupby(["client", "year", "month", "category"])["amount"].sum().reset_index()
        return grouped

    def validate_monthly_data(self, df):
        """월별 데이터 검증 및 로깅"""
        self.logger.info("🔍 월별 데이터 검증 시작")
        
        try:
            # 월별 데이터 개수 및 금액 집계
            monthly_summary = df.groupby(df['date'].dt.month).agg({
                'amount': ['count', 'sum']
            }).round(0)
            
            monthly_summary.columns = ['건수', '금액']
            
            self.logger.info("📈 월별 데이터 현황:")
            for month, row in monthly_summary.iterrows():
                count = int(row['건수'])
                amount = int(row['금액'])
                self.logger.info(f"   {month:2d}월: {count:>6,}건 | {amount:>15,}원")
            
            # 7월 데이터 특별 검사
            july_data = df[df['date'].dt.month == 7]
            if july_data.empty:
                self.logger.error("⚠️ 중요: 7월 데이터가 없습니다!")
            else:
                july_count = len(july_data)
                july_amount = july_data['amount'].sum()
                july_date_range = f"{july_data['date'].min().strftime('%Y-%m-%d')} ~ {july_data['date'].max().strftime('%Y-%m-%d')}"
                self.logger.info(f"✅ 7월 데이터 정상: {july_count:,}건, {july_amount:,.0f}원 ({july_date_range})")
            
            # 빈 월 찾기
            expected_months = set(range(1, 8))  # 1월~7월
            actual_months = set(monthly_summary.index)
            missing_months = expected_months - actual_months
            
            if missing_months:
                missing_list = sorted(missing_months)
                self.logger.warning(f"⚠️ 누락된 월: {missing_list}")
            else:
                self.logger.info("✅ 모든 월 데이터 존재 확인")
                
        except Exception as e:
            self.logger.error(f"월별 데이터 검증 오류: {e}")

    def save_pivot_to_excel(self, dataframes, output_path):
        """피벗 테이블을 엑셀로 저장"""
        category_order = self.config.get_category_order()
        
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 기존 파일이 있으면 백업
        if self.backup_manager and output_path.exists():
            backup_path = self.backup_manager.create_backup(output_path)
            if backup_path:
                self.logger.info(f"✅ 기존 분석 결과 백업: {backup_path}")
            else:
                self.logger.warning("⚠️ 백업 생성 실패 - 계속 진행")
        
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            sheet_written = False
            
            for sheet_name, df in dataframes.items():
                if df.empty:
                    self.logger.warning(f"빈 시트 건너뜀: {sheet_name}")
                    continue

                try:
                    if sheet_name == "주차별":
                        pivot = df.pivot_table(index="기간", columns="category", values="amount", fill_value=0)
                        pivot = pivot.reset_index()
                        
                    elif sheet_name == "거래처별_월별":
                        pivot = df.pivot_table(index=["client", "year", "month"], columns="category", values="amount", fill_value=0)
                        pivot = pivot.reset_index()
                        
                    elif sheet_name == "월별":
                        pivot = df.pivot_table(index=["year", "month"], columns="category", values="amount", fill_value=0)
                        pivot = pivot.reset_index()
                        
                    else:
                        first_col = df.columns[0]
                        pivot = df.pivot_table(index=first_col, columns="category", values="amount", fill_value=0)
                        pivot = pivot.reset_index()

                    # 카테고리 컬럼들만 정렬
                    if sheet_name == "월별":
                        base_columns = ["year", "month"]
                        category_columns = [col for col in pivot.columns if col not in base_columns]
                        available_categories = [cat for cat in category_order if cat in category_columns]
                        pivot = pivot[base_columns + available_categories]
                        
                    elif sheet_name == "거래처별_월별":
                        base_columns = ["client", "year", "month"]
                        category_columns = [col for col in pivot.columns if col not in base_columns]
                        available_categories = [cat for cat in category_order if cat in category_columns]
                        pivot = pivot[base_columns + available_categories]
                        
                    elif sheet_name == "주차별":
                        base_columns = ["기간"]
                        category_columns = [col for col in pivot.columns if col not in base_columns]
                        available_categories = [cat for cat in category_order if cat in category_columns]
                        pivot = pivot[base_columns + available_categories]

                    # 합계 행 추가
                    if sheet_name == "월별":
                        total_row = {"year": "합계", "month": ""}
                        for cat in available_categories:
                            total_row[cat] = pivot[cat].sum()
                        pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)
                        
                    elif sheet_name == "거래처별_월별":
                        total_row = {"client": "합계", "year": "", "month": ""}
                        for cat in available_categories:
                            total_row[cat] = pivot[cat].sum()
                        pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)
                        
                    elif sheet_name == "주차별":
                        total_row = {"기간": "합계"}
                        for cat in available_categories:
                            total_row[cat] = pivot[cat].sum()
                        pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)

                    # 인덱스 없이 저장
                    pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.logger.debug(f"시트 저장 완료: {sheet_name} ({len(pivot)}행, {len(pivot.columns)}열)")
                    sheet_written = True
                    
                except Exception as e:
                    self.logger.error(f"시트 저장 실패 {sheet_name}: {e}")

            if not sheet_written:
                pd.DataFrame({"메시지": ["저장할 데이터가 없습니다."]}).to_excel(writer, sheet_name="결과없음", index=False)
                self.logger.warning("저장할 유효한 데이터가 없습니다")

    def process_sales_data(self, output_filename="매출집계_결과.xlsx"):
        """매출 데이터 전체 처리 프로세스"""
        self.logger.info("매출 데이터 처리 시작 (금~목 기준) - 리팩토링됨")
        
        raw_dir = self.config.get_sales_raw_data_dir()
        files = list(raw_dir.glob("**/*판매조회*.xlsx"))
        self.logger.info(f"발견된 파일: {len(files)}개")

        if not files:
            self.logger.error("처리할 파일이 없습니다")
            raise FileNotFoundError(f"매출 데이터 파일을 찾을 수 없습니다: {raw_dir}")

        all_data = []
        failed_files = []

        for file_path in files:
            fname = file_path.name
            
            try:
                if "디앤드디" in fname:
                    df = self.load_and_standardize(file_path, "디앤드디")
                    if df.empty:
                        failed_files.append(fname)
                        continue
                    df = self.categorize_and_filter(df, "디앤드디")
                    
                elif "디앤아이" in fname:
                    company_config = self.config.get_company_config("디앤아이")
                    default_category = company_config.get("default_category")
                    df = self.load_and_standardize(file_path, "디앤아이", default_category=default_category)
                    if df.empty:
                        failed_files.append(fname)
                        continue
                        
                elif "후지리프트코리아" in fname:
                    company_config = self.config.get_company_config("후지리프트코리아")
                    default_category = company_config.get("default_category")
                    df = self.load_and_standardize(file_path, "후지리프트코리아", default_category=default_category)
                    if df.empty:
                        failed_files.append(fname)
                        continue
                else:
                    self.logger.warning(f"알 수 없는 파일 건너뜀: {fname}")
                    continue
                    
                df = self.enrich_with_time_columns(df)
                all_data.append(df)
                
            except Exception as e:
                self.logger.error(f"파일 처리 실패 {fname}: {e}")
                failed_files.append(fname)

        # 실패한 파일 처리
        if failed_files:
            self.logger.warning(f"처리되지 않은 파일들: {failed_files}")
            if len(failed_files) == len(files):
                self.logger.error("모든 매출 데이터 파일 처리 실패")
                raise RuntimeError("모든 매출 데이터 파일 처리 실패")

        if not all_data:
            self.logger.error("유효한 데이터가 없습니다")
            raise ValueError("매출 데이터 처리 후 유효한 데이터가 없습니다")

        # 전체 데이터 결합
        full_df = pd.concat(all_data, ignore_index=True)
        self.logger.info(f"총 데이터: {len(full_df):,}행 (금액: {full_df['amount'].sum():,.0f}원)")
        
        # 월별 데이터 검증 추가
        self.validate_monthly_data(full_df)

        # 결과 집계
        results = {
            "월별": self.summarize_monthly_data(full_df),
            "주차별": self.summarize_weekly_data(full_df),
            "거래처별_월별": self.summarize_client_monthly_data(full_df)
        }

        # 각 결과 데이터프레임 정보 출력
        for name, df in results.items():
            self.logger.info(f"{name}: {len(df):,}행")

        # 결과 저장
        output_path = self.config.get_processed_data_dir() / output_filename
        self.save_pivot_to_excel(results, output_path)
        self.logger.info(f"결과 저장 완료: {output_path}")
        
        return results, full_df


def main():
    """메인 실행 함수"""
    try:
        calculator = SalesCalculator()
        results, full_df = calculator.process_sales_data()
        logging.info("매출 데이터 처리 완료 - 리팩토링됨")
        return results, full_df
    except Exception as e:
        logging.error(f"전체 프로세스 오류: {e}")
        raise


if __name__ == "__main__":
    main()
