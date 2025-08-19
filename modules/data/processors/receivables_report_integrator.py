#!/usr/bin/env python3
"""
매출채권 보고서 통합기 (리팩토링됨)
매출채권 분석 결과를 주간보고서에 통합
"""
import pandas as pd
import numpy as np
from pathlib import Path
import logging
from datetime import datetime
import sys

# 리팩토링된 경로 설정
sys.path.append(str(Path(__file__).parent.parent.parent))

# 설정 관리자 import (새 구조)
from modules.utils.config_manager import get_config


class ReceivablesReportIntegrator:
    """매출채권 보고서 통합기 (리팩토링됨)"""
    
    def __init__(self):
        self.logger = logging.getLogger('ReceivablesReportIntegrator')
        self.config = get_config()
        
        # 리팩토리된 경로 설정
        self.processed_dir = self.config.get_processed_data_dir()
        self.receivables_file = self.processed_dir / "채권_분석_결과.xlsx"
        
    def find_receivables_result_file(self):
        """매출채권 분석 결과 파일 찾기"""
        try:
            # 기본 경로에서 찾기
            if self.receivables_file.exists():
                self.logger.info(f"매출채권 분석 결과 파일 발견: {self.receivables_file}")
                return self.receivables_file
            
            # 추가 경로에서 찾기
            possible_paths = [
                self.processed_dir / "receivables_analysis_result.xlsx",
                self.processed_dir / "매출채권분석결과.xlsx"
            ]
            
            for path in possible_paths:
                if path.exists():
                    self.logger.info(f"매출채권 분석 결과 파일 발견: {path}")
                    return path
            
            self.logger.error("매출채권 분석 결과 파일을 찾을 수 없습니다")
            return None
            
        except Exception as e:
            self.logger.error(f"파일 검색 중 오류: {e}")
            return None
    
    def read_receivables_result_file(self, file_path=None):
        """매출채권 분석 결과 파일 읽기"""
        try:
            if file_path is None:
                file_path = self.find_receivables_result_file()
            
            if file_path is None or not file_path.exists():
                self.logger.error(f"파일이 존재하지 않습니다: {file_path}")
                return None
            
            # 파일 형식 확인
            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names
                self.logger.info(f"파일 시트 목록: {sheet_names}")
                
                return self.read_standard_format(xls, sheet_names)
                        
        except Exception as e:
            self.logger.error(f"파일 읽기 실패: {e}")
            return None
    
    def read_standard_format(self, xls, sheet_names):
        """표준 형식 파일 읽기"""
        sheets_data = {}
        
        # 표준 시트 매핑
        target_sheets = {
            "파일정보": "파일정보",
            "요약": "요약",
            "계산 결과": "계산결과",
            "계산결과": "계산결과",  # 호환성
            "TOP20_금주": "TOP20_금주",
            "TOP20금주": "TOP20_금주"  # 호환성
        }
        
        for original_name, mapped_name in target_sheets.items():
            if original_name in sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=original_name)
                    sheets_data[mapped_name] = df
                    self.logger.info(f"시트 '{original_name}' → '{mapped_name}' 로드됨: {len(df)}행")
                except Exception as e:
                    self.logger.error(f"시트 '{original_name}' 읽기 실패: {e}")
                    sheets_data[mapped_name] = pd.DataFrame()
            else:
                self.logger.warning(f"시트 '{original_name}'이 없습니다")
                sheets_data[mapped_name] = pd.DataFrame()
        
        return sheets_data
    
    def format_summary_sheet(self, summary_df):
        """요약 시트 포맷팅"""
        if summary_df.empty:
            return pd.DataFrame()
        
        try:
            formatted_df = summary_df.copy()
            
            # 숫자 컬럼들 식별
            numeric_keywords = ['총채권', '90일', '결제예정일', '장기미수', '채권', '금액']
            
            for col in formatted_df.columns:
                # 숫자 데이터인지 확인
                if any(keyword in str(col) for keyword in numeric_keywords):
                    # 비율이나 퍼센트가 아닌 경우에만 백만원 단위로 변환
                    if '비율' not in str(col) and '%' not in str(col) and 'p)' not in str(col):
                        if formatted_df[col].dtype in ['int64', 'float64']:
                            # 값이 1000000 이상인 경우에만 백만원 단위로 변환
                            mask = formatted_df[col] >= 1000000
                            if mask.any():
                                formatted_df.loc[mask, col] = formatted_df.loc[mask, col] / 1000000
                                # 컬럼명에 단위 추가
                                if '백만' not in str(col) and 'M' not in str(col):
                                    formatted_df = formatted_df.rename(columns={col: f"{col}(백만원)"})
            
            return formatted_df
            
        except Exception as e:
            self.logger.error(f"요약 시트 포맷팅 실패: {e}")
            return summary_df
    
    def clean_data_for_excel(self, data):
        """엑셀 생성을 위한 데이터 정리"""
        try:
            if pd.isna(data) or data is None:
                return ""
            elif isinstance(data, (int, float)):
                if pd.isna(data) or np.isinf(data):
                    return 0
                return data
            else:
                # 문자열로 변환하고 특수문자 제거
                clean_str = str(data).replace('\\n', ' ').replace('\\r', ' ')
                # Excel에서 문제가 될 수 있는 문자들 제거
                clean_str = ''.join(char for char in clean_str if ord(char) < 65536)
                return clean_str
        except Exception:
            return ""
    
    def create_integrated_receivables_sheet(self, sheets_data):
        """통합된 매출채권 시트 생성"""
        try:
            integrated_data = []
            
            # 헤더 정보 추가
            integrated_data.append([f"=== 매출채권 분석 결과 ({datetime.now().strftime('%Y-%m-%d %H:%M')}) ==="])
            integrated_data.append([])
            
            # 1. 파일정보 섹션
            if "파일정보" in sheets_data and not sheets_data["파일정보"].empty:
                integrated_data.append(["📄 파일 정보"])
                integrated_data.append([])
                
                file_info_df = sheets_data["파일정보"]
                # 컬럼 헤더 추가
                if not file_info_df.empty:
                    integrated_data.append(file_info_df.columns.tolist())
                    # 데이터 추가
                    for _, row in file_info_df.iterrows():
                        integrated_data.append(row.tolist())
                
                integrated_data.append([])
                integrated_data.append([])
            
            # 2. 요약 섹션 (포맷팅 적용)
            if "요약" in sheets_data and not sheets_data["요약"].empty:
                integrated_data.append(["📊 요약 분석"])
                integrated_data.append([])
                
                summary_df = self.format_summary_sheet(sheets_data["요약"])
                if not summary_df.empty:
                    integrated_data.append(summary_df.columns.tolist())
                    for _, row in summary_df.iterrows():
                        integrated_data.append(row.tolist())
                
                integrated_data.append([])
                integrated_data.append([])
            
            # 3. 계산결과 섹션
            if "계산결과" in sheets_data and not sheets_data["계산결과"].empty:
                integrated_data.append(["🔢 전주 대비 계산 결과"])
                integrated_data.append([])
                
                calc_df = sheets_data["계산결과"]
                if not calc_df.empty:
                    integrated_data.append(calc_df.columns.tolist())
                    for _, row in calc_df.iterrows():
                        integrated_data.append(row.tolist())
                
                integrated_data.append([])
                integrated_data.append([])
            
            # 4. TOP20 섹션
            if "TOP20_금주" in sheets_data and not sheets_data["TOP20_금주"].empty:
                integrated_data.append(["🏆 TOP 20 기간초과 채권 거래처"])
                integrated_data.append([])
                
                top20_df = sheets_data["TOP20_금주"]
                if not top20_df.empty:
                    integrated_data.append(top20_df.columns.tolist())
                    # 최대 20행만 추가
                    for _, row in top20_df.head(20).iterrows():
                        integrated_data.append(row.tolist())
            
            # 빈 데이터 처리
            if not integrated_data:
                integrated_data = [["매출채권 데이터가 없습니다."]]
            
            # DataFrame으로 변환
            max_cols = max(len(row) for row in integrated_data) if integrated_data else 1
            
            # 모든 행을 동일한 컬럼 수로 맞춤
            normalized_data = []
            for row in integrated_data:
                if len(row) < max_cols:
                    row.extend([''] * (max_cols - len(row)))
                # 데이터 정리 적용
                cleaned_row = [self.clean_data_for_excel(cell) for cell in row[:max_cols]]
                normalized_data.append(cleaned_row)
            
            # 컬럼명 생성
            column_names = [f"컬럼{i+1}" for i in range(max_cols)]
            
            integrated_df = pd.DataFrame(normalized_data, columns=column_names)
            
            self.logger.info(f"통합 매출채권 시트 생성 완료: {len(integrated_df)}행 x {max_cols}열")
            return integrated_df
            
        except Exception as e:
            self.logger.error(f"통합 시트 생성 실패: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            
            # 실패 시 기본 데이터 반환
            return pd.DataFrame([["매출채권 데이터 통합 실패"]], columns=["오류"])
    
    def integrate_receivables_data_to_report(self, weekly_report_path, receivables_file_path=None):
        """주간보고서에 매출채권 데이터 통합"""
        try:
            self.logger.info("=== 매출채권 데이터 통합 시작 (리팩토링됨) ===")
            
            # 1. 매출채권 분석 결과 파일 읽기
            if receivables_file_path:
                file_path = Path(receivables_file_path)
            else:
                file_path = self.find_receivables_result_file()
            
            if not file_path or not file_path.exists():
                self.logger.error("매출채권 분석 결과 파일을 찾을 수 없습니다")
                return False
            
            sheets_data = self.read_receivables_result_file(file_path)
            if not sheets_data:
                self.logger.error("매출채권 데이터 읽기 실패")
                return False
            
            # 2. 통합 매출채권 시트 생성
            integrated_sheet = self.create_integrated_receivables_sheet(sheets_data)
            if integrated_sheet.empty or (len(integrated_sheet.columns) == 1 and "오류" in integrated_sheet.columns):
                self.logger.error("통합 시트 생성 실패")
                return False
            
            # 3. 주간보고서 파일 확인
            weekly_report = Path(weekly_report_path)
            if not weekly_report.exists():
                self.logger.error(f"주간보고서 파일이 없습니다: {weekly_report}")
                return False
            
            # 4. 기존 주간보고서 읽기
            try:
                with pd.ExcelFile(weekly_report) as xls:
                    existing_sheets = {}
                    for sheet_name in xls.sheet_names:
                        if sheet_name != "매출 채권":  # 기존 매출채권 시트는 덮어쓰기
                            try:
                                existing_sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                                self.logger.debug(f"기존 시트 로드: {sheet_name}")
                            except Exception as e:
                                self.logger.warning(f"시트 '{sheet_name}' 읽기 실패: {e}")
                                # 빈 시트로 처리
                                existing_sheets[sheet_name] = pd.DataFrame()
            except Exception as e:
                self.logger.error(f"주간보고서 읽기 실패: {e}")
                return False
            
            # 5. 업데이트된 파일 저장
            try:
                with pd.ExcelWriter(weekly_report, engine='xlsxwriter', options={'remove_timezone': True}) as writer:
                    # 기존 시트들 먼저 저장
                    for sheet_name, df in existing_sheets.items():
                        if not df.empty:
                            # 데이터 정리
                            cleaned_df = df.copy()
                            for col in cleaned_df.columns:
                                if cleaned_df[col].dtype == 'object':
                                    cleaned_df[col] = cleaned_df[col].astype(str).replace(['nan', 'None', 'inf', '-inf'], '')
                                elif cleaned_df[col].dtype in ['int64', 'float64']:
                                    cleaned_df[col] = cleaned_df[col].fillna(0).replace([np.inf, -np.inf], 0)
                            
                            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        else:
                            # 빈 시트 처리
                            pd.DataFrame([["데이터 없음"]]).to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 매출채권 시트 추가 (데이터 정리 적용)
                    final_integrated_sheet = integrated_sheet.copy()
                    
                    # 컬럼별 데이터 정리
                    for col in final_integrated_sheet.columns:
                        if final_integrated_sheet[col].dtype == 'object':
                            final_integrated_sheet[col] = final_integrated_sheet[col].astype(str).replace(['nan', 'None'], '')
                        elif final_integrated_sheet[col].dtype in ['int64', 'float64']:
                            final_integrated_sheet[col] = final_integrated_sheet[col].fillna(0).replace([np.inf, -np.inf], 0)
                    
                    final_integrated_sheet.to_excel(writer, sheet_name="매출 채권", index=False)
                    
                    self.logger.info("✅ 주간보고서에 매출채권 데이터 통합 완료")
                
                return True
                
            except Exception as e:
                self.logger.error(f"파일 저장 실패: {e}")
                return False
            
        except Exception as e:
            self.logger.error(f"주간보고서 통합 실패: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            return False
    
    def check_receivables_data_availability(self):
        """매출채권 데이터 가용성 확인"""
        try:
            file_path = self.find_receivables_result_file()
            if not file_path:
                return False, "매출채권 분석 결과 파일이 없습니다."
            
            sheets_data = self.read_receivables_result_file(file_path)
            if not sheets_data:
                return False, "매출채권 데이터 읽기 실패"
            
            # 최소 하나의 시트라도 데이터가 있는지 확인
            has_data = any(not df.empty for df in sheets_data.values())
            if not has_data:
                return False, "매출채권 데이터가 비어있습니다."
            
            return True, "매출채권 데이터 사용 가능"
            
        except Exception as e:
            return False, f"매출채권 데이터 확인 중 오류: {e}"
    
    def test_integration(self):
        """통합 기능 테스트"""
        try:
            self.logger.info("=== 매출채권 통합 기능 테스트 시작 (리팩토링됨) ===")
            
            # 1. 파일 찾기 테스트
            file_path = self.find_receivables_result_file()
            if file_path:
                print(f"✅ 매출채권 파일 발견: {file_path}")
                
                # 2. 파일 읽기 테스트
                sheets_data = self.read_receivables_result_file(file_path)
                if sheets_data:
                    print("✅ 파일 읽기 성공")
                    for sheet_name, df in sheets_data.items():
                        print(f"  - {sheet_name}: {len(df)}행")
                    
                    # 3. 통합 시트 생성 테스트
                    integrated_sheet = self.create_integrated_receivables_sheet(sheets_data)
                    if not integrated_sheet.empty:
                        print(f"✅ 통합 시트 생성 성공: {len(integrated_sheet)}행")
                        
                        # 테스트 파일로 저장
                        test_output = self.processed_dir / "테스트_통합_매출채권_결과.xlsx"
                        
                        try:
                            # 데이터 정리 후 저장
                            cleaned_sheet = integrated_sheet.copy()
                            for col in cleaned_sheet.columns:
                                if cleaned_sheet[col].dtype == 'object':
                                    cleaned_sheet[col] = cleaned_sheet[col].astype(str).replace(['nan', 'None'], '')
                                elif cleaned_sheet[col].dtype in ['int64', 'float64']:
                                    cleaned_sheet[col] = cleaned_sheet[col].fillna(0).replace([np.inf, -np.inf], 0)
                            
                            with pd.ExcelWriter(test_output, engine='xlsxwriter', options={'remove_timezone': True}) as writer:
                                cleaned_sheet.to_excel(writer, sheet_name="통합_매출채권", index=False)
                            print(f"✅ 테스트 파일 저장: {test_output}")
                        except Exception as e:
                            print(f"❌ 테스트 파일 저장 실패: {e}")
                        
                        return True
                    else:
                        print("❌ 통합 시트 생성 실패")
                        return False
                else:
                    print("❌ 파일 읽기 실패")
                    return False
            else:
                print("❌ 매출채권 파일을 찾을 수 없습니다")
                return False
                
        except Exception as e:
            print(f"❌ 테스트 실패: {e}")
            import traceback
            traceback.print_exc()
            return False


def main():
    """테스트용 메인 함수"""
    try:
        integrator = ReceivablesReportIntegrator()
        
        # 데이터 가용성 확인
        is_available, message = integrator.check_receivables_data_availability()
        print(f"매출채권 데이터 상태: {message}")
        
        if is_available:
            # 테스트 실행
            success = integrator.test_integration()
            if success:
                print("🎉 매출채권 통합 기능 테스트 성공!")
            else:
                print("💥 매출채권 통합 기능 테스트 실패!")
            return success
        else:
            print("매출채권 데이터를 먼저 생성해주세요.")
            return False
            
    except Exception as e:
        print(f"❌ 테스트 중 오류: {e}")
        return False


if __name__ == "__main__":
    main()
