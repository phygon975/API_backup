"""
설정 파일

이 파일은 애플리케이션의 주요 설정값들을 포함합니다.
"""

from typing import Dict, Any

# =============================================================================
# 기본 설정값들
# =============================================================================

# 기본 재질
DEFAULT_MATERIAL = "CS"

# 기본 설치 인수
DEFAULT_INSTALL_FACTOR = 1.0

# 기본 CEPCI 설정
DEFAULT_TARGET_YEAR = 2024
DEFAULT_CEPCI_BASE_YEAR = 2017
DEFAULT_CEPCI_BASE_INDEX = 567.5

# =============================================================================
# 장비별 기본 타입
# =============================================================================

DEFAULT_EQUIPMENT_TYPES = {
    "pump": "centrifugal",
    "compressor": "centrifugal", 
    "turbine": "axial",
    "fan": "centrifugal_radial"
}

# =============================================================================
# Aspen Plus 관련 설정
# =============================================================================

# Aspen Plus 애플리케이션 이름
ASPEN_APPLICATION_NAME = "Aspen Plus"

# 기본 단위 세트 이름
DEFAULT_UNIT_SET = "SI"

# =============================================================================
# 비용 계산 관련 설정
# =============================================================================

# MCompr BM 인수 추가 비율 (일반 압축기 대비)
MCOMPR_BM_FACTOR_MULTIPLIER = 1.2

# 열교환기 기본 BM 인수
DEFAULT_HEAT_EXCHANGER_BM_FACTOR = 2.5

# Intercooler 기본 비용 (USD)
DEFAULT_INTERCOOLER_COST_PER_STAGE = 5000.0

# =============================================================================
# 출력 설정
# =============================================================================

# 출력 구분선 길이
OUTPUT_SEPARATOR_LENGTH = 60

# 소수점 자릿수
DECIMAL_PLACES = 2

# =============================================================================
# 디버그 설정
# =============================================================================

# 디버그 출력 활성화
ENABLE_DEBUG_OUTPUT = True

# 상세 계산 과정 출력 활성화
ENABLE_DETAILED_CALCULATION_OUTPUT = True

# =============================================================================
# 파일 경로 설정
# =============================================================================

# 백업 파일 접미사
BACKUP_FILE_SUFFIX = "_backup"

# 로그 파일 이름
LOG_FILE_NAME = "equipment_cost_calculation.log"

# =============================================================================
# 유효성 검사 설정
# =============================================================================

# 최소 압력 상승 (bar) - 팬 범위 판정 기준
FAN_MAX_PRESSURE_RISE = 0.16

# 최소 압력 하강 (bar) - 터빈 범위 판정 기준  
TURBINE_MIN_PRESSURE_DROP = 0.1

# =============================================================================
# 사용자 인터페이스 설정
# =============================================================================

# 프리뷰 모드에서 표시할 최대 항목 수
MAX_PREVIEW_ITEMS = 50

# 사용자 입력 타임아웃 (초)
USER_INPUT_TIMEOUT = 30

# =============================================================================
# 성능 최적화 설정
# =============================================================================

# 캐시 크기 제한
MAX_CACHE_SIZE = 1000

# COM 호출 간격 (초)
COM_CALL_INTERVAL = 0.1

# =============================================================================
# 오류 처리 설정
# =============================================================================

# 최대 재시도 횟수
MAX_RETRY_ATTEMPTS = 3

# 오류 로깅 활성화
ENABLE_ERROR_LOGGING = True

# =============================================================================
# 설정 검증 함수
# =============================================================================

def validate_config() -> bool:
    """설정값들의 유효성을 검증"""
    try:
        # 기본값들이 올바른지 확인
        assert DEFAULT_MATERIAL in ["CS", "SS", "TI", "NI"]
        assert DEFAULT_INSTALL_FACTOR > 0
        assert DEFAULT_TARGET_YEAR >= 2017
        assert DEFAULT_CEPCI_BASE_INDEX > 0
        assert MCOMPR_BM_FACTOR_MULTIPLIER > 1.0
        assert DEFAULT_HEAT_EXCHANGER_BM_FACTOR > 0
        assert FAN_MAX_PRESSURE_RISE > 0
        assert TURBINE_MIN_PRESSURE_DROP > 0
        
        return True
    except AssertionError:
        return False


def get_config_summary() -> Dict[str, Any]:
    """설정 요약 정보 반환"""
    return {
        "default_material": DEFAULT_MATERIAL,
        "default_target_year": DEFAULT_TARGET_YEAR,
        "default_cepi_base_index": DEFAULT_CEPCI_BASE_INDEX,
        "enable_debug_output": ENABLE_DEBUG_OUTPUT,
        "fan_max_pressure_rise": FAN_MAX_PRESSURE_RISE,
        "turbine_min_pressure_drop": TURBINE_MIN_PRESSURE_DROP,
        "mcompr_bm_multiplier": MCOMPR_BM_FACTOR_MULTIPLIER,
    }
