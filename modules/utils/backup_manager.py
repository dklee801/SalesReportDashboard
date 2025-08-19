#!/usr/bin/env python3
"""
자동 백업 기능 모듈
파일 수정 전 자동으로 백업을 생성하고 관리
"""

import shutil
from pathlib import Path
from datetime import datetime, timedelta
import logging
from typing import Optional, List

class BackupManager:
    """파일 백업 관리 클래스"""
    
    def __init__(self, backup_retention_days: int = 30):
        """
        Args:
            backup_retention_days: 백업 파일 보관 기간 (일)
        """
        self.backup_retention_days = backup_retention_days
        self.logger = logging.getLogger(__name__)
        
    def create_backup(self, file_path: Path, backup_dir: Optional[Path] = None) -> Optional[Path]:
        """
        파일 백업 생성
        
        Args:
            file_path: 백업할 파일 경로
            backup_dir: 백업 저장 디렉토리 (기본값: 원본 파일 경로/backup)
            
        Returns:
            백업 파일 경로 (실패 시 None)
        """
        try:
            if not file_path.exists():
                self.logger.warning(f"백업할 파일이 존재하지 않습니다: {file_path}")
                return None
                
            # 백업 디렉토리 설정
            if backup_dir is None:
                backup_dir = file_path.parent / "backup"
            
            backup_dir.mkdir(parents=True, exist_ok=True)
            
            # 타임스탬프 포함 백업 파일명 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{file_path.stem}_{timestamp}{file_path.suffix}"
            backup_path = backup_dir / backup_filename
            
            # 파일 복사 (메타데이터 포함)
            shutil.copy2(file_path, backup_path)
            
            self.logger.info(f"백업 생성 완료: {backup_path}")
            
            # 오래된 백업 파일 정리
            self.cleanup_old_backups(backup_dir)
            
            return backup_path
            
        except Exception as e:
            self.logger.error(f"백업 생성 실패: {e}")
            return None
    
    def cleanup_old_backups(self, backup_dir: Path, custom_retention_days: Optional[int] = None):
        """
        오래된 백업 파일 정리
        
        Args:
            backup_dir: 백업 디렉토리
            custom_retention_days: 사용자 정의 보관 기간 (None이면 기본값 사용)
        """
        try:
            retention_days = custom_retention_days or self.backup_retention_days
            cutoff_time = datetime.now() - timedelta(days=retention_days)
            
            deleted_count = 0
            deleted_size = 0
            
            # 백업 디렉토리의 모든 파일 검사
            for backup_file in backup_dir.iterdir():
                if backup_file.is_file():
                    # 파일 수정 시간 확인
                    mtime = datetime.fromtimestamp(backup_file.stat().st_mtime)
                    
                    if mtime < cutoff_time:
                        file_size = backup_file.stat().st_size
                        backup_file.unlink()
                        deleted_count += 1
                        deleted_size += file_size
                        self.logger.debug(f"오래된 백업 삭제: {backup_file.name}")
            
            if deleted_count > 0:
                self.logger.info(
                    f"오래된 백업 정리 완료: {deleted_count}개 파일, "
                    f"{deleted_size / (1024*1024):.1f}MB 삭제"
                )
                
        except Exception as e:
            self.logger.error(f"백업 정리 중 오류: {e}")
    
    def restore_backup(self, backup_path: Path, target_path: Path) -> bool:
        """
        백업 파일 복원
        
        Args:
            backup_path: 백업 파일 경로
            target_path: 복원할 위치
            
        Returns:
            성공 여부
        """
        try:
            if not backup_path.exists():
                self.logger.error(f"백업 파일이 존재하지 않습니다: {backup_path}")
                return False
            
            # 기존 파일이 있으면 임시 백업
            temp_backup = None
            if target_path.exists():
                temp_backup = target_path.with_suffix(target_path.suffix + '.tmp')
                shutil.move(str(target_path), str(temp_backup))
            
            try:
                # 백업 파일 복원
                shutil.copy2(backup_path, target_path)
                self.logger.info(f"백업 복원 완료: {backup_path} → {target_path}")
                
                # 임시 백업 삭제
                if temp_backup and temp_backup.exists():
                    temp_backup.unlink()
                    
                return True
                
            except Exception as e:
                # 복원 실패 시 원래 파일 복구
                if temp_backup and temp_backup.exists():
                    shutil.move(str(temp_backup), str(target_path))
                raise e
                
        except Exception as e:
            self.logger.error(f"백업 복원 실패: {e}")
            return False
    
    def list_backups(self, original_file_path: Path, backup_dir: Optional[Path] = None) -> List[dict]:
        """
        특정 파일의 백업 목록 조회
        
        Args:
            original_file_path: 원본 파일 경로
            backup_dir: 백업 디렉토리
            
        Returns:
            백업 파일 정보 리스트
        """
        backups = []
        
        try:
            if backup_dir is None:
                backup_dir = original_file_path.parent / "backup"
            
            if not backup_dir.exists():
                return backups
            
            # 원본 파일명 기준으로 백업 파일 검색
            file_stem = original_file_path.stem
            file_suffix = original_file_path.suffix
            
            for backup_file in backup_dir.glob(f"{file_stem}_*{file_suffix}"):
                if backup_file.is_file():
                    stat = backup_file.stat()
                    
                    # 타임스탬프 추출 시도
                    timestamp_str = backup_file.stem.replace(f"{file_stem}_", "")
                    try:
                        backup_time = datetime.strptime(timestamp_str, "%Y%m%d_%H%M%S")
                    except:
                        backup_time = datetime.fromtimestamp(stat.st_mtime)
                    
                    backups.append({
                        'path': backup_file,
                        'name': backup_file.name,
                        'size_mb': stat.st_size / (1024 * 1024),
                        'created': backup_time,
                        'age_days': (datetime.now() - backup_time).days
                    })
            
            # 최신 순으로 정렬
            backups.sort(key=lambda x: x['created'], reverse=True)
            
        except Exception as e:
            self.logger.error(f"백업 목록 조회 실패: {e}")
        
        return backups


# 사용 예시를 위한 통합 함수
def integrate_backup_with_file_operation(file_path: Path, operation_func, *args, **kwargs):
    """
    파일 작업 전 자동 백업을 수행하는 데코레이터 함수
    
    Args:
        file_path: 작업할 파일 경로
        operation_func: 수행할 작업 함수
        *args, **kwargs: 작업 함수에 전달할 인자
        
    Returns:
        작업 함수의 반환값
    """
    backup_manager = BackupManager()
    
    # 백업 생성
    backup_path = backup_manager.create_backup(file_path)
    
    if backup_path:
        print(f"✅ 백업 생성됨: {backup_path}")
    else:
        print("⚠️ 백업 생성 실패 - 작업을 계속합니다")
    
    try:
        # 실제 작업 수행
        result = operation_func(file_path, *args, **kwargs)
        return result
        
    except Exception as e:
        print(f"❌ 작업 실패: {e}")
        
        # 백업이 있으면 복원 제안
        if backup_path and backup_path.exists():
            response = input("백업을 복원하시겠습니까? (y/n): ")
            if response.lower() == 'y':
                if backup_manager.restore_backup(backup_path, file_path):
                    print("✅ 백업 복원 완료")
                else:
                    print("❌ 백업 복원 실패")
        
        raise e


# 기존 모듈과의 통합 예시
if __name__ == "__main__":
    # 테스트 코드
    import tempfile
    
    # 임시 파일 생성
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xlsx', delete=False) as f:
        test_file = Path(f.name)
        f.write("테스트 데이터")
    
    # 백업 관리자 생성
    backup_mgr = BackupManager(backup_retention_days=7)
    
    # 백업 생성 테스트
    print("1. 백업 생성 테스트")
    backup_path = backup_mgr.create_backup(test_file)
    print(f"   백업 경로: {backup_path}")
    
    # 백업 목록 조회 테스트
    print("\n2. 백업 목록 조회 테스트")
    backups = backup_mgr.list_backups(test_file)
    for backup in backups:
        print(f"   - {backup['name']} ({backup['size_mb']:.2f}MB, {backup['age_days']}일 전)")
    
    # 백업 복원 테스트
    print("\n3. 백업 복원 테스트")
    if backups:
        # 원본 파일 수정
        test_file.write_text("수정된 데이터")
        print(f"   원본 수정 후: {test_file.read_text()}")
        
        # 백업 복원
        if backup_mgr.restore_backup(backups[0]['path'], test_file):
            print(f"   복원 후: {test_file.read_text()}")
    
    # 정리
    test_file.unlink()
    if backup_path:
        backup_path.unlink()
        backup_path.parent.rmdir()
    
    print("\n✅ 테스트 완료")
