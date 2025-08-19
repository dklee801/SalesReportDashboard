#!/usr/bin/env python3
"""
Auto Sales Report Generator v3.1
매출 보고서 자동 생성 시스템 메인 실행 파일 - 통합 매출채권 처리기 적용
"""

import sys
import argparse
import logging  # logging을 전역으로 이동
from pathlib import Path
from datetime import datetime

# 프로젝트 루트 디렉토리를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# modules 디렉토리를 Python 경로에 추가
modules_dir = project_root / "modules"
sys.path.insert(0, str(modules_dir))

# 필요한 모듈들 import (새 구조에 맞게 수정)
try:
    import pandas as pd
    from modules.utils.config_manager import get_config
    from modules.core.sales_calculator import main as analyze_sales
    
    # 매출채권 분석기 import
    try:
        from modules.core.accounts_receivable_analyzer import main as analyze_receivables
    except ImportError:
        logging.error("매출채권 분석 모듈을 찾을 수 없습니다.")
        analyze_receivables = None
    
    # 통합 데이터 수집기 import
    try:
        from modules.data.unified_data_collector import UnifiedDataCollector
    except ImportError:
        logging.warning("통합 데이터 수집기를 찾을 수 없습니다.")
        UnifiedDataCollector = None
    
    # 주간보고서 생성기 import
    try:
        from modules.reports.report_generator import WeeklyReportGenerator
    except ImportError:
        logging.warning("주간보고서 생성 모듈을 찾을 수 없습니다.")
        WeeklyReportGenerator = None
        
except ImportError as e:
    logging.error(f"모듈 import 실패: {e}")
    logging.error("modules 폴더의 파일들이 모두 있는지 확인해주세요.")
    sys.exit(1)


def setup_logging(quiet=False):
    """통일된 로깅 설정"""
    level = logging.WARNING if quiet else logging.INFO
    
    # 기본 로거 설정
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s' if not quiet else '%(message)s',
        datefmt='%H:%M:%S'
    )
    
    return logging.getLogger(__name__)


def setup_argument_parser():
    """단순화된 명령행 인수 파서 설정"""
    parser = argparse.ArgumentParser(
        description="매출 보고서 자동 생성 시스템 v3.1 (주간 비교 및 전주 대비 변화 분석)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python main.py                      # 전체 프로세스 (매출 3개월 + 매출채권 주간분석)
  python main.py --collect            # 데이터 수집만 (매출 3개월 + 매출채권 12주)
  python main.py --collect-sales      # 매출 데이터 수집만
  python main.py --collect-receivables # 매출채권 데이터 수집만
  python main.py --process            # 분석 처리만 실행 (수집 건너뜀)
  python main.py --report             # 보고서 생성만 실행
  python main.py --months 6           # 6개월 데이터 (매출 6개월 + 매출채권 24주)
  python main.py --quiet              # 최소 출력으로 실행
  python main.py --show-browser       # 브라우저 창 표시 (디버깅용)
        """
    )
    
    # 실행 모드 (상호 배타적)
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--collect", 
        action="store_true",
        help="데이터 수집만 실행 (매출 + 매출채권)"
    )
    mode_group.add_argument(
        "--collect-sales", 
        action="store_true",
        help="매출 데이터 수집만 실행"
    )
    mode_group.add_argument(
        "--collect-receivables", 
        action="store_true",
        help="매출채권 데이터 수집만 실행"
    )
    mode_group.add_argument(
        "--process", 
        action="store_true",
        help="분석 처리만 실행 (데이터 수집 건너뜀)"
    )
    mode_group.add_argument(
        "--report", 
        action="store_true",
        help="보고서 생성만 실행"
    )

    # 설정 옵션
    parser.add_argument(
        "--months", 
        type=int, 
        default=3,
        help="매출 데이터 수집 기간 (개월, 기본값: 3)"
    )
    
    parser.add_argument(
        "--weeks", 
        type=int, 
        default=None,
        help="매출채권 수집 주수 (기본값: months * 4)"
    )
    
    # 실행 제어 옵션
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="최소한의 출력만 표시"
    )
    
    parser.add_argument(
        "--show-browser",
        action="store_true",
        help="브라우저 창 표시 (기본: 백그라운드 실행)"
    )
    
    return parser


def analyze_sales_data(logger) -> bool:
    """매출 데이터 분석"""
    logger.info("매출 데이터 분석 중...")
    
    try:
        result = analyze_sales()
        if result:
            logger.info("매출 데이터 분석 완료")
            return True
        else:
            logger.error("매출 데이터 분석 실패 - 결과 없음")
            return False
    except Exception as e:
        logger.error(f"매출 데이터 분석 실패: {e}")
        return False


def analyze_receivables_data(logger) -> bool:
    """매출채권 데이터 분석"""
    logger.info("매출채권 데이터 분석 중...")
    
    try:
        if analyze_receivables is not None:
            result = analyze_receivables()
            if result:
                logger.info("매출채권 데이터 분석 완료")
                return True
            else:
                logger.error("매출채권 데이터 분석 실패 - 결과 없음")
                return False
        else:
            logger.error("매출채권 분석 모듈을 찾을 수 없습니다")
            return False
    except Exception as e:
        logger.error(f"매출채권 데이터 분석 실패: {e}")
        return False


def generate_report(logger) -> bool:
    """보고서 생성"""
    logger.info("보고서 생성 중...")
    
    try:
        if WeeklyReportGenerator is None:
            logger.error("주간보고서 생성 모듈을 찾을 수 없습니다")
            return False
        
        generator = WeeklyReportGenerator()
        success = generator.generate_report()
        
        if success:
            logger.info("보고서 생성 완료")
            logger.info(f"결과 파일: {generator.result_path}")
            return True
        else:
            logger.error("보고서 생성 실패")
            return False
        
    except Exception as e:
        logger.error(f"보고서 생성 실패: {e}")
        return False


def print_summary(successful_steps: list, total_steps: int, quiet: bool = False):
    """실행 결과 요약 출력"""
    if quiet:
        if len(successful_steps) == total_steps:
            print("✅ 모든 단계 완료")
        else:
            failed_count = total_steps - len(successful_steps)
            print(f"⚠️ {failed_count}개 단계 실패")
        return
    
    print("\n" + "="*60)
    print("🎉 자동화 프로세스 완료!")
    
    if len(successful_steps) == total_steps:
        print("✅ 모든 단계가 성공적으로 완료되었습니다!")
    else:
        print(f"✅ 성공한 단계: {', '.join(successful_steps)}")
        failed_count = total_steps - len(successful_steps)
        if failed_count > 0:
            print(f"❌ 실패한 단계: {failed_count}개")
    
    # 결과 파일 위치 출력
    try:
        config = get_config()
        processed_dir = config.get_processed_data_dir()
        report_dir = config.get_report_output_dir()
        
        print(f"\n📁 처리된 데이터 위치: {processed_dir}")
        print(f"📄 보고서 파일 위치: {report_dir}")
        
        # 생성된 파일들 확인
        result_files = []
        if (processed_dir / "매출집계_결과.xlsx").exists():
            result_files.append("매출집계_결과.xlsx")
        if (processed_dir / "채권_분석_결과.xlsx").exists():
            result_files.append("채권_분석_결과.xlsx")

        # 주간보고서 파일들 확인 (report 폴더에서)
        report_files = []
        if report_dir.exists():
            for file_path in report_dir.glob("주간보고서_*.xlsx"):
                report_files.append(file_path.name)
                
        if result_files:
            print("📊 분석 데이터 파일들:")
            for file in result_files:
                print(f"   - {file}")
        
        if report_files:
            print("📋 보고서 파일들:")
            for file in report_files:
                print(f"   - {file}")
        
        print(f"\n📅 보고서 기준: 월요일~금요일 (주말 제외)")
        print(f"💰 매출채권 수집: 매주 금요일 기준")
            
    except Exception as e:
        logging.warning(f"결과 파일 위치 확인 실패: {e}")


def main():
    """단순화된 메인 실행 함수"""
    parser = setup_argument_parser()
    args = parser.parse_args()
    
    # 로깅 설정
    logger = setup_logging(args.quiet)
    
    # weeks 기본값 설정
    if args.weeks is None:
        args.weeks = args.months * 4
    
    if not args.quiet:
        print("🚀 매출 보고서 자동 생성 시스템 v3.1")
        print(f"⏰ 실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if not args.show_browser:
            print("🔇 헤드리스 모드 (브라우저 창 숨김)")
    
    successful_steps = []
    
    try:
        # 실행 모드에 따른 분기
        if args.collect:
            # 전체 데이터 수집 (구현 필요)
            logger.info("전체 데이터 수집 모드 (구현 필요)")
            total_steps = 1
                
        elif args.collect_sales:
            # 매출 데이터만 수집 (구현 필요)
            logger.info("매출 데이터 수집 모드 (구현 필요)")
            total_steps = 1
                
        elif args.collect_receivables:
            # 매출채권 데이터만 수집 (구현 필요)
            logger.info("매출채권 데이터 수집 모드 (구현 필요)")
            total_steps = 1
                
        elif args.process:
            # 분석 처리만
            logger.info("분석 처리 모드")
            if analyze_sales_data(logger):
                successful_steps.append("매출 분석")
            if analyze_receivables_data(logger):
                successful_steps.append("매출채권 분석")
            total_steps = 2
                
        elif args.report:
            # 보고서 생성만
            logger.info("보고서 생성 모드")
            if generate_report(logger):
                successful_steps.append("보고서 생성")
            total_steps = 1
                
        else:
            # 전체 프로세스 (기본값)
            logger.info("전체 프로세스 모드")
            
            if analyze_sales_data(logger):
                successful_steps.append("매출 분석")
                
            if analyze_receivables_data(logger):
                successful_steps.append("매출채권 분석")
                
            if generate_report(logger):
                successful_steps.append("보고서 생성")
            
            total_steps = 3
        
        # 결과 출력
        print_summary(successful_steps, total_steps, args.quiet)
        return len(successful_steps) == total_steps
        
    except KeyboardInterrupt:
        print("\n⚠️ 사용자에 의해 중단되었습니다.")
        return False
    except Exception as e:
        logger.error(f"예상치 못한 오류 발생: {e}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
