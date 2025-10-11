"""
설정 관리 (config.json)
- 폴더 경로, 프린터 설정 등을 JSON 파일로 저장/로드
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, Optional, List


class ConfigManager:
    """프로그램 설정을 JSON 파일로 관리하는 클래스"""
    
    def __init__(self, config_path: str = "config.json"):
        """
        Args:
            config_path: 설정 파일 경로
        """
        self.config_path = Path(config_path)
        self.config = self._load_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """기본 설정값 반환"""
        return {
            "base_path_settings": {
                "base_path": "",  # 통합 기본 경로 (검색 + 저장)
                "last_used": "",  # 마지막 사용 시간
                "use_date_subfolder": False,  # 날짜별 하위폴더 사용
                "recent_paths": []  # 최근 사용한 경로들 (최대 5개)
            },
            "search_settings": {
                "base_folder": "",  # PDF 검색 기본 폴더 (호환성 유지)
                "order_pattern": r"\d{13}",  # 주문번호 정규식 (13자리 숫자)
                "recursive_search": True  # 하위 폴더까지 검색
            },
            "print_settings": {
                "printer_name": "",  # 기본 프린터명
                "copies": 1,  # 인쇄 매수
                "duplex": False,  # 양면 인쇄
                "sumatra_path": "SumatraPDF.exe"  # SumatraPDF 경로
            },
            "ui_settings": {
                "window_size": [1000, 800],
                "remember_last_search": True
            }
        }
    
    def _load_config(self) -> Dict[str, Any]:
        """설정 파일 로드"""
        if not self.config_path.exists():
            # 설정 파일이 없으면 기본 설정으로 생성
            default_config = self._get_default_config()
            self._save_config(default_config)
            return default_config
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 기본 설정과 병합 (새로운 키가 추가된 경우 대응)
            default_config = self._get_default_config()
            merged_config = self._merge_config(default_config, config)
            
            # 병합된 설정이 다르면 저장
            if merged_config != config:
                self._save_config(merged_config)
            
            return merged_config
            
        except (json.JSONDecodeError, FileNotFoundError) as e:
            print(f"설정 파일 로드 실패: {e}")
            default_config = self._get_default_config()
            self._save_config(default_config)
            return default_config
    
    def _merge_config(self, default: Dict[str, Any], user: Dict[str, Any]) -> Dict[str, Any]:
        """기본 설정과 사용자 설정을 병합"""
        result = default.copy()
        
        for key, value in user.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_config(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def _save_config(self, config: Dict[str, Any]):
        """설정을 파일에 저장"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"설정 파일 저장 실패: {e}")
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """
        중첩된 키로 설정값 가져오기
        
        Args:
            key_path: "category.key" 형식의 키 경로
            default: 기본값
            
        Returns:
            설정값
            
        Example:
            config.get("search_settings.base_folder")
        """
        keys = key_path.split('.')
        value = self.config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        
        return value
    
    def set(self, key_path: str, value: Any, save: bool = True):
        """
        중첩된 키로 설정값 저장
        
        Args:
            key_path: "category.key" 형식의 키 경로
            value: 저장할 값
            save: 즉시 파일에 저장할지 여부
        """
        keys = key_path.split('.')
        current = self.config
        
        # 마지막 키를 제외하고 중첩 딕셔너리 생성
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        
        # 마지막 키에 값 설정
        current[keys[-1]] = value
        
        if save:
            self._save_config(self.config)
    
    def get_base_folder(self) -> str:
        """PDF 검색 기본 폴더 경로 가져오기"""
        return self.get("search_settings.base_folder", "")
    
    def set_base_folder(self, folder_path: str):
        """PDF 검색 기본 폴더 경로 설정"""
        self.set("search_settings.base_folder", folder_path)
    
    def get_printer_name(self) -> str:
        """기본 프린터명 가져오기"""
        return self.get("print_settings.printer_name", "")
    
    def set_printer_name(self, printer_name: str):
        """기본 프린터명 설정"""
        self.set("print_settings.printer_name", printer_name)
    
    def get_sumatra_path(self) -> str:
        """SumatraPDF 경로 가져오기"""
        return self.get("print_settings.sumatra_path", "SumatraPDF.exe")
    
    def set_sumatra_path(self, path: str):
        """SumatraPDF 경로 설정"""
        self.set("print_settings.sumatra_path", path)
    
    def get_order_pattern(self) -> str:
        """주문번호 정규식 패턴 가져오기"""
        return self.get("search_settings.order_pattern", r"[A-Z]-\d{7}")
    
    def export_config(self, export_path: str):
        """설정을 다른 파일로 내보내기"""
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"설정 내보내기 실패: {e}")
            return False
    
    def import_config(self, import_path: str) -> bool:
        """다른 설정 파일에서 가져오기"""
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                imported_config = json.load(f)
            
            # 기본 설정과 병합
            self.config = self._merge_config(self._get_default_config(), imported_config)
            self._save_config(self.config)
            return True
        except Exception as e:
            print(f"설정 가져오기 실패: {e}")
            return False
    
    # ===== 통합 경로 관리 메서드들 =====
    
    def get_base_path(self) -> str:
        """통합 기본 경로 가져오기"""
        return self.get("base_path_settings.base_path", "")
    
    def set_base_path(self, path: str):
        """통합 기본 경로 설정"""
        from datetime import datetime
        
        # 경로 정규화
        normalized_path = str(Path(path).resolve()) if path else ""
        
        # 기본 경로 설정
        self.set("base_path_settings.base_path", normalized_path)
        self.set("base_path_settings.last_used", datetime.now().isoformat())
        
        # 최근 경로 목록 업데이트
        if normalized_path:
            self._update_recent_paths(normalized_path)
        
        # 호환성을 위해 기존 설정도 업데이트
        self.set("search_settings.base_folder", normalized_path, save=False)
    
    def _update_recent_paths(self, new_path: str):
        """최근 경로 목록 업데이트 (최대 5개)"""
        recent_paths = self.get("base_path_settings.recent_paths", [])
        
        # 기존에 있던 경로면 제거
        if new_path in recent_paths:
            recent_paths.remove(new_path)
        
        # 맨 앞에 추가
        recent_paths.insert(0, new_path)
        
        # 최대 5개까지만 유지
        recent_paths = recent_paths[:5]
        
        self.set("base_path_settings.recent_paths", recent_paths, save=False)
    
    def get_recent_paths(self) -> List[str]:
        """최근 사용한 경로 목록 가져오기"""
        recent_paths = self.get("base_path_settings.recent_paths", [])
        # 존재하는 경로만 필터링
        valid_paths = [path for path in recent_paths if os.path.exists(path)]
        
        # 유효한 경로만 남아있도록 업데이트
        if len(valid_paths) != len(recent_paths):
            self.set("base_path_settings.recent_paths", valid_paths)
        
        return valid_paths
    
    def get_use_date_subfolder(self) -> bool:
        """날짜별 하위폴더 사용 여부"""
        return self.get("base_path_settings.use_date_subfolder", False)
    
    def set_use_date_subfolder(self, use_date: bool):
        """날짜별 하위폴더 사용 설정"""
        self.set("base_path_settings.use_date_subfolder", use_date)
    
    def get_working_path(self) -> str:
        """
        실제 작업할 경로 가져오기
        날짜별 하위폴더 옵션이 활성화되어 있으면 YYYY-MM-DD 폴더 생성
        """
        base_path = self.get_base_path()
        if not base_path:
            return ""
        
        if self.get_use_date_subfolder():
            from datetime import datetime
            today = datetime.now().strftime('%Y-%m-%d')
            working_path = os.path.join(base_path, today)
            
            # 날짜 폴더가 없으면 생성
            try:
                os.makedirs(working_path, exist_ok=True)
                return working_path
            except Exception as e:
                print(f"날짜 폴더 생성 실패 ({working_path}): {e}")
                return base_path
        
        return base_path
    
    def validate_base_path(self, path: str = None) -> tuple[bool, str]:
        """
        기본 경로 유효성 검사
        
        Args:
            path: 검사할 경로 (None이면 현재 설정된 경로)
            
        Returns:
            (유효여부, 메시지)
        """
        if path is None:
            path = self.get_base_path()
        
        if not path:
            return False, "경로가 설정되지 않았습니다."
        
        if not os.path.exists(path):
            return False, f"폴더가 존재하지 않습니다: {path}"
        
        if not os.path.isdir(path):
            return False, f"유효한 폴더가 아닙니다: {path}"
        
        # 쓰기 권한 확인
        try:
            test_file = os.path.join(path, '.write_test')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            return True, "유효한 경로입니다."
        except Exception as e:
            return False, f"폴더에 쓰기 권한이 없습니다: {path}\n({str(e)})"
    
    def get_suggested_paths(self) -> List[str]:
        """추천 경로 목록 (데스크톱, 문서, 다운로드 등)"""
        suggestions = []
        
        try:
            import os
            # Windows 사용자 폴더들
            user_home = os.path.expanduser("~")
            
            common_folders = [
                os.path.join(user_home, "Desktop", "PDF_WORK"),
                os.path.join(user_home, "Documents", "PDF_WORK"),
                os.path.join(user_home, "Downloads", "PDF_WORK"),
                "C:\\PDF_WORK",
                "D:\\PDF_WORK"
            ]
            
            # 존재하지 않지만 생성 가능한 경로들도 포함
            for folder in common_folders:
                parent = os.path.dirname(folder)
                if os.path.exists(parent) and os.access(parent, os.W_OK):
                    suggestions.append(folder)
                    
        except Exception as e:
            print(f"추천 경로 생성 실패: {e}")
        
        return suggestions
    
    def create_base_path(self, path: str) -> tuple[bool, str]:
        """기본 경로 생성"""
        try:
            os.makedirs(path, exist_ok=True)
            return True, f"폴더가 생성되었습니다: {path}"
        except Exception as e:
            return False, f"폴더 생성 실패: {str(e)}"


# 전역 설정 관리자 인스턴스
config = ConfigManager()
