"""
Create on Sep 4, 2025

@author: Pyeong-Gon Jung
"""

import os
import win32com.client as win32
import numpy as np
import sys
import time
from threading import Thread
from typing import Optional
from equipment_costs import (
    register_default_correlations,
    CEPCIOptions,
    calculate_pressure_device_costs_auto,
    preview_pressure_devices_auto,
    print_preview_results,
    preview_heat_exchangers_auto,
    print_preview_hx_results,
    estimate_heat_exchanger_cost,
    get_hx_material_options,
    calculate_pressure_device_costs_with_data,
    clear_aspen_cache,
    get_cache_stats,
)
from aspen_data_extractor import get_unit_type_value
from unit_converter import get_cepi_index, convert_to_si_units
from config import (
    DEFAULT_MATERIAL, DEFAULT_TARGET_YEAR, DEFAULT_CEPCI_BASE_INDEX,
    OUTPUT_SEPARATOR_LENGTH, ENABLE_DEBUG_OUTPUT
)
import pickle
from dataclasses import dataclass, field
from typing import Dict, List, Any
import contextlib


#======================================================================
# Session model for preview/overrides persistence
#======================================================================
@dataclass
class PreviewSession:
    """프리뷰 결과와 사용자 오버라이드를 한 번에 저장/복원하는 세션 컨테이너"""
    aspen_file: str
    current_unit_set: Optional[str]
    block_info: Dict[str, str]  # {block_name: category}
    preview: List[Dict[str, Any]]  # preview_pressure_devices_auto 결과
    material_overrides: Dict[str, str] = field(default_factory=dict)
    type_overrides: Dict[str, str] = field(default_factory=dict)
    subtype_overrides: Dict[str, str] = field(default_factory=dict)

    def apply_overrides_to_preview(self) -> List[Dict[str, Any]]:
        updated_preview: List[Dict[str, Any]] = []
        for p in self.preview:
            up = p.copy()
            name = p.get('name')
            if name in self.material_overrides:
                up['material'] = self.material_overrides[name]
            if name in self.type_overrides:
                up['selected_type'] = self.type_overrides[name]
            if name in self.subtype_overrides:
                up['selected_subtype'] = self.subtype_overrides[name]
            updated_preview.append(up)
        return updated_preview

    def save(self, path: str) -> None:
        with open(path, 'wb') as f:
            pickle.dump(self, f)

    @staticmethod
    def load(path: str) -> "PreviewSession":
        with open(path, 'rb') as f:
            return pickle.load(f)

    # JSON 저장 기능은 혼동 방지를 위해 비활성화됨

#======================================================================
# Spinner
#======================================================================
class Spinner:
    """Simple CLI spinner to indicate progress during long-running tasks."""
    def __init__(self, message: str) -> None:
        self.message = message
        self._running = False
        self._thread = None

    def start(self) -> None:
        if self._running:
            return
        self._running = True
        self._thread = Thread(target=self._spin, daemon=True)
        self._thread.start()

    def _spin(self) -> None:
        frames = ['|', '/', '-', '\\']
        idx = 0
        while self._running:
            sys.stdout.write(f"\r{self.message} {frames[idx % len(frames)]}")
            sys.stdout.flush()
            time.sleep(0.1)
            idx += 1

    def stop(self, done_message: Optional[str] = None) -> None:
        if not self._running:
            return
        self._running = False
        if self._thread is not None:
            self._thread.join(timeout=0.5)
        sys.stdout.write('\r')
        if done_message:
            print(done_message)
        else:
            print('')

#======================================================================
# Aspen Plus Connection
#======================================================================

# 현재 폴더에서 .bkp 파일 자동 탐지 및 선택
current_dir = os.path.dirname(os.path.abspath(__file__))
# 최근 수정순으로 정렬된 .bkp 파일 목록
all_bkps = [f for f in os.listdir(current_dir) if f.lower().endswith('.bkp')]
bkp_files = sorted(all_bkps, key=lambda f: os.path.getmtime(os.path.join(current_dir, f)), reverse=True)

selected_bkp = None
if not bkp_files:
    # 기본 파일명으로 시도 (기존 동작 유지)
    default_file = 'Equipment_cost_estimation_aspen.bkp'
    print("경고: 현재 폴더에서 .bkp 파일을 찾지 못했습니다.")
    print(f"기본 파일명으로 시도합니다: {default_file}")
    selected_bkp = default_file
else:
    print("\n감지된 .bkp 파일 목록:")
    for i, fname in enumerate(bkp_files, 1):
        print(f"  {i}. {fname}")
    # 사용자에게 선택 받기 (검증 포함)
    while True:
        try:
            choice = input("사용할 .bkp 파일 번호를 선택하세요 (숫자): ").strip()
            idx = int(choice)
            if 1 <= idx <= len(bkp_files):
                selected_bkp = bkp_files[idx - 1]
                break
            else:
                print("잘못된 번호입니다. 다시 입력해주세요.")
        except ValueError:
            print("숫자를 입력해주세요.")

aspen_Path = os.path.join(current_dir, selected_bkp)
print(f"선택된 파일: {aspen_Path}")

try:
    # 4. Initiate Aspen Plus application
    print('\nConnecting to Aspen Plus... Please wait...')
    connect_spinner = Spinner('Connecting to Aspen Plus')
    connect_spinner.start()
    Application = win32.Dispatch('Apwn.Document') # Registered name of Aspen Plus
    connect_spinner.stop('Aspen Plus COM object created successfully!')
    
    # 5. Try to open the file
    print(f'Attempting to open file: {aspen_Path}')
    open_spinner = Spinner('Opening Aspen backup')
    open_spinner.start()
    Application.InitFromArchive2(aspen_Path)    
    open_spinner.stop('File opened successfully!')
    
    # 6. Make the files visible
    Application.visible = 1   
    print('Aspen Plus is now visible')

except Exception as e:
    print(f"ERROR connecting to Aspen Plus: {e}")
    print("\nPossible solutions:")
    print("1. Make sure Aspen Plus is installed on your computer")
    print("2. Make sure Aspen Plus is properly licensed")
    print("3. Try running Aspen Plus manually first to ensure it works")
    print("4. Check if the .bkp file is compatible with your Aspen Plus version")
    exit()

#======================================================================
# Block Classifier
#======================================================================
def get_block_names(Application):
    """
    Blocks 하위의 가장 상위 노드(블록 이름)들을 수집하는 함수
    """
    block_names = []
    
    try:
        # Blocks 노드 찾기
        blocks_node = Application.Tree.FindNode("\\Data\\Blocks")
        if blocks_node is None:
            print("Warning: Blocks node not found")
            return block_names
        
        # Blocks 하위의 직접적인 자식들만 수집 (가장 상위 노드)
        if hasattr(blocks_node, 'Elements') and blocks_node.Elements is not None:
            for element in blocks_node.Elements:
                try:
                    block_names.append(element.Name)
                except:
                    # 예외 발생 시 조용히 건너뛰기(에러메시지 출력 x)
                    pass
        
        return block_names
        
    except Exception as e:
        print(f"Error collecting block names: {str(e)}")
        return block_names

block_names = get_block_names(Application)
print(block_names)

#======================================================================
#Block Classifier
#======================================================================
def parse_bkp_file_for_blocks(file_path, block_names):
    """
    .bkp 파일을 텍스트로 읽어서 주어진 블록 이름들의 카테고리를 파싱하는 함수
    """
    block_info = {}
    
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        lines = content.split('\n')
        
        # 각 블록 이름에 대해 카테고리 찾기
        for block_name in block_names:
            category = "Unknown"
            
            # 블록 이름이 있는 줄 찾기
            for i, line in enumerate(lines):
                if line.strip() == block_name:
                    # 다음 4줄에서 카테고리 정보 찾기
                    for j in range(i + 1, min(i + 5, len(lines))):
                        next_line = lines[j].strip()
                        
                        # 카테고리 후보들
                        if next_line in ['Heater', 'Cooler', 'HeatX', 'Condenser']:
                            category = next_line
                            break
                        elif next_line in ['RadFrac', 'Distl', 'DWSTU']:
                            category = next_line
                            break
                        elif next_line in ['RStoic', 'RCSTR', 'RPlug', 'RBatch', 'REquil', 'RYield']:
                            category = next_line
                            break
                        elif next_line in ['Pump', 'Compr', 'MCompr', 'Vacuum', 'Flash', 'Sep', 'Mixer', 'FSplit', 'Valve']:
                            category = next_line
                            break
                        
                    
                    break  # 블록 이름을 찾았으므로 루프 종료
            
            block_info[block_name] = category
        
        return block_info
        
    except Exception as e:
        print(f"Error parsing BKP file: {str(e)}")
        return {}

def classify_blocks_from_bkp(file_path, block_names):
    """
    .bkp 파일에서 주어진 블록 이름들의 카테고리를 분류하는 함수
    """
    block_info = parse_bkp_file_for_blocks(file_path, block_names)
    
    block_categories = {
        'heat_exchangers': [],
        'distillation_columns': [],
        'reactors': [],
        'pumps and compressors': [],
        'vessels': [],
        'vacuum_systems': [],
        'Ignore': [],
        'other_blocks': []
    }
    
    for block_name, category in block_info.items():
        if category in ['Heater', 'Cooler', 'HeatX', 'Condenser']:
            block_categories['heat_exchangers'].append(block_name)
        elif category in ['RadFrac', 'Distl', 'DWSTU']:
            block_categories['distillation_columns'].append(block_name)
        elif category in ['RStoic', 'RCSTR', 'RPlug', 'RBatch', 'REquil', 'RYield']:
            block_categories['reactors'].append(block_name)
        elif category in ['Pump', 'Compr', 'MCompr']:
            block_categories['pumps and compressors'].append(block_name)
        elif category in ['Vacuum', 'Flash', 'Sep']:
            block_categories['vessels'].append(block_name)
        elif category in ['Mixer', 'FSplit', 'Valve']:
            block_categories['Ignore'].append(block_name)
        else:
            block_categories['other_blocks'].append(block_name)
    
    return block_categories, block_info

block_categories, block_info = classify_blocks_from_bkp(aspen_Path, block_names)

#======================================================================
# Device Loader Functions
#======================================================================

def get_heat_exchangers(block_categories):
    """
    열교환기 장치들만 반환하는 함수
    """
    return block_categories.get('heat_exchangers', [])

def get_distillation_columns(block_categories):
    """
    증류탑 장치들만 반환하는 함수
    """
    return block_categories.get('distillation_columns', [])

def get_reactors(block_categories):
    """
    반응기 장치들만 반환하는 함수
    """
    return block_categories.get('reactors', [])

def get_pumps_and_compressors(block_categories):
    """
    펌프와 압축기 장치들만 반환하는 함수
    """
    return block_categories.get('pumps and compressors', [])

def get_vessels(block_categories):
    """
    용기 장치들만 반환하는 함수
    """
    return block_categories.get('vessels', [])

def get_ignored_devices(block_categories):
    """
    무시할 장치들만 반환하는 함수
    """
    return block_categories.get('Ignore', [])

def get_other_devices(block_categories):
    """
    기타 장치들만 반환하는 함수
    """
    return block_categories.get('other_blocks', [])

#======================================================================
# Usage Examples
#======================================================================

print("\n" + "="*60)
print("DEVICE CATEGORIES")
print("="*60)

# 열교환기만 가져오기
heat_exchangers = get_heat_exchangers(block_categories)
print(f"\nHeat Exchangers ({len(heat_exchangers)} devices):")
for he in heat_exchangers:
    print(f"  - {he}")

# 증류탑만 가져오기
distillation_columns = get_distillation_columns(block_categories)
print(f"\nDistillation Columns ({len(distillation_columns)} devices):")
for dc in distillation_columns:
    print(f"  - {dc}")

# 반응기만 가져오기
reactors = get_reactors(block_categories)
print(f"\nReactors ({len(reactors)} devices):")
for reactor in reactors:
    print(f"  - {reactor}")

# 펌프와 압축기만 가져오기
pumps_compressors = get_pumps_and_compressors(block_categories)
print(f"\nPumps and Compressors ({len(pumps_compressors)} devices):")
for pc in pumps_compressors:
    print(f"  - {pc}")

# 용기만 가져오기
vessels = get_vessels(block_categories)
print(f"\nVessels ({len(vessels)} devices):")
for vessel in vessels:
    print(f"  - {vessel}")

# 무시할 장치들만 가져오기
ignored_devices = get_ignored_devices(block_categories)
print(f"\nIgnored Devices ({len(ignored_devices)} devices):")
for ignored in ignored_devices:
    print(f"  - {ignored}")

# 기타 장치들만 가져오기
other_devices = get_other_devices(block_categories)
print(f"\nOther Devices ({len(other_devices)} devices):")
for other in other_devices:
    print(f"  - {other}")

print(f"\n" + "="*60)
print("DEVICE LOADING COMPLETED")
print("="*60)
    

#======================================================================
# Unit detection
#======================================================================


def get_units_sets(Application):
    """
    Aspen Plus에서 사용된 단위 세트들을 가져오는 함수
    """
    units_sets = []
    
    try:
        # Units-Sets 노드 찾기
        units_sets_node = Application.Tree.FindNode("\\Data\\Setup\\Units-Sets")
        if units_sets_node is None:
            return units_sets
        
        # Units-Sets 하위의 직접적인 자식들 수집
        if hasattr(units_sets_node, 'Elements') and units_sets_node.Elements is not None:
            for element in units_sets_node.Elements:
                try:
                    # 'Current'는 제외하고 실제 unit set 이름들만 수집
                    if element.Name != 'Current':
                        units_sets.append(element.Name)
                except:
                    # 예외 발생 시 조용히 건너뛰기
                    pass
        
    except Exception as e:
        # 조용히 실패
        pass
    
    return units_sets

def get_current_unit_set(Application):
    """
    현재 사용 중인 Unit Set을 가져오는 함수
    
    Parameters:
    -----------
    Application : Aspen Plus COM object
        Aspen Plus 애플리케이션 객체
    
    Returns:
    --------
    str or None : 현재 사용 중인 Unit Set 이름
    """
    try:
        # OUTSET 노드에서 현재 사용 중인 Unit Set 가져오기
        outset_node = Application.Tree.FindNode("\\Data\\Setup\\Global\\Input\\OUTSET")
        
        if outset_node is None:
            print("Warning: OUTSET node not found")
            return None
        
        current_unit_set = outset_node.Value
        
        if current_unit_set:
            print(f"Current unit set: {current_unit_set}")
            return current_unit_set
        else:
            print("Warning: No current unit set found")
            return None
            
    except Exception as e:
        print(f"Error getting current unit set: {str(e)}")
        return None


# 사용하지 않는 함수 제거됨

def get_unit_set_details(Application, unit_set_name, unit_table):
    """
    특정 단위 세트의 상세 정보를 가져오고 하드코딩된 데이터와 연동하는 함수
    """
    # 필요한 unit_type들과 해당 인덱스 매핑
    required_unit_types = {
        'AREA': 1, 'COMPOSITION': 2, 'DENSITY': 3, 'ENERGY': 5, 'FLOW': 9,
        'MASS-FLOW': 10, 'MOLE-FLOW': 11, 'VOLUME-FLOW': 12, 'MASS': 18,
        'POWER': 19, 'PRESSURE': 20, 'TEMPERATURE': 22, 'TIME': 24,
        'VELOCITY': 25, 'VOLUME': 27, 'MOLE-DENSITY': 37, 'MASS-DENSITY': 38,
        'MOLE-VOLUME': 43, 'ELEC-POWER': 47, 'UA': 50, 'WORK': 52, 'HEAT': 53
    }
    
    unit_details = {
        'name': unit_set_name,
        'unit_types': {},
        'index_mapping': required_unit_types
    }
    
    try:
        # 각 unit_type에 대해 정보 가져오기
        for unit_type, aspen_index in required_unit_types.items():
            try:
                # Unit-Types 노드에서 해당 unit_type 찾기
                unit_type_node = Application.Tree.FindNode(f"\\Data\\Setup\\Units-Sets\\{unit_set_name}\\Unit-Types\\{unit_type}")
                if unit_type_node:
                    # 단위 값 가져오기
                    unit_value = unit_type_node.Value
                    
                    # 하드코딩된 데이터에서 해당 unit_type의 Physical Quantity 인덱스 찾기
                    physical_quantity_index = get_physical_quantity_by_unit_type(unit_table, unit_type)
                    
                    # 하드코딩된 데이터에서 해당 unit의 Unit of Measure 인덱스 찾기
                    unit_of_measure_index = None
                    if physical_quantity_index and physical_quantity_index in unit_table:
                        for unit_idx, hardcoded_unit in unit_table[physical_quantity_index]['units'].items():
                            if hardcoded_unit == unit_value:
                                unit_of_measure_index = unit_idx
                                break
                    
                    unit_details['unit_types'][unit_type] = {
                        'value': unit_value,
                        'aspen_index': aspen_index,
                        'csv_column_index': physical_quantity_index,
                        'unit_index_in_csv': unit_of_measure_index,
                        'data_available': physical_quantity_index is not None
                    }
                else:
                    # 노드를 찾을 수 없는 경우
                    physical_quantity_index = get_physical_quantity_by_unit_type(unit_table, unit_type)
                    unit_details['unit_types'][unit_type] = {
                        'value': 'Not Found in Aspen',
                        'aspen_index': aspen_index,
                        'csv_column_index': physical_quantity_index,
                        'unit_index_in_csv': None,
                        'data_available': physical_quantity_index is not None
                    }
            except Exception as e:
                # 예외 발생 시
                physical_quantity_index = get_physical_quantity_by_unit_type(unit_table, unit_type)
                unit_details['unit_types'][unit_type] = {
                    'value': f'Error: {str(e)}',
                    'aspen_index': aspen_index,
                    'csv_column_index': physical_quantity_index,
                    'unit_index_in_csv': None,
                    'data_available': physical_quantity_index is not None
                }
                
    except Exception as e:
        print(f"Warning: Could not get details for unit set '{unit_set_name}': {e}")
    
    return unit_details

def print_unit_set_details(unit_details):
    """
    단위 세트 상세 정보를 Physical Quantity와 Unit of Measure로 출력하는 함수
    """
    print(f"\nUnit Set: {unit_details['name']}")
    print("-" * 100)
    
    if unit_details['unit_types']:
        print(f"{'Unit Type':<20} {'Physical Quantity':<18} {'Value':<20} {'Unit of Measure':<15} {'Data Available':<15}")
        print("-" * 100)
        
        for unit_type, info in unit_details['unit_types'].items():
            csv_idx = info['unit_index_in_csv'] if info['unit_index_in_csv'] is not None else 'N/A'
            data_avail = 'Yes' if info['data_available'] else 'No'
            print(f"{unit_type:<20} {info['aspen_index']:<18} {info['value']:<20} {csv_idx:<15} {data_avail:<15}")
    else:
        print("  No unit types found")

def get_unit_by_indices(unit_table, physical_quantity_index, unit_of_measure_index):
    """
    Physical Quantity 인덱스와 Unit of Measure 인덱스로 unit 값을 가져오는 함수
    """
    return get_unit_by_index(unit_table, physical_quantity_index, unit_of_measure_index)

def get_available_units_for_type(unit_table, unit_type_name):
    """
    특정 unit_type의 모든 사용 가능한 unit들을 가져오는 함수
    """
    physical_quantity_index = get_physical_quantity_by_unit_type(unit_table, unit_type_name)
    if physical_quantity_index:
        return get_units_by_physical_quantity(unit_table, physical_quantity_index)
    return {}

def print_units_sets_summary(units_sets):
    """
    단위 세트 요약 정보를 출력하는 함수
    """
    print("\n" + "="*60)
    print("UNITS SETS SUMMARY")
    print("="*60)
    
    if not units_sets:
        print("No unit sets found")
        return
    
    print(f"Total unit sets found: {len(units_sets)}")
    print("\nUnit sets:")
    for i, unit_set in enumerate(units_sets, 1):
        print(f"  {i:2d}. {unit_set}")
    
    print("="*60)

# 사용하지 않는 함수 제거됨

# 사용하지 않는 함수 제거됨

# 사용하지 않는 함수 제거됨

#======================================================================
# Hardcoded Unit Data (for CSV-free operation)
#======================================================================

def get_hardcoded_unit_table():
    """
    CSV 파일 없이도 작동하도록 하드코딩된 단위 테이블을 반환하는 함수
    Unit_table.csv의 내용을 기반으로 함
    """
    # CSV 열 순서에 따른 unit_type 매핑 (1부터 시작)
    csv_column_to_unit_type = {
        1: 'AREA',           # sqm
        2: 'COMPOSITION',    # mol-fr
        3: 'DENSITY',        # kg/cum
        4: 'ENERGY',         # J
        5: 'FLOW',           # kg/sec
        6: 'MASS-FLOW',      # kg/sec
        7: 'MOLE-FLOW',      # kmol/sec
        8: 'VOLUME-FLOW',    # cum/sec
        9: 'MASS',           # kg
        10: 'POWER',         # Watt
        11: 'PRESSURE',      # N/sqm
        12: 'TEMPERATURE',   # K
        13: 'TIME',          # sec
        14: 'VELOCITY',      # m/sec
        15: 'VOLUME',        # cum
        16: 'MOLE-DENSITY',  # kmol/cum
        17: 'MASS-DENSITY',  # kg/cum
        18: 'MOLE-VOLUME',   # cum/kmol
        19: 'ELEC-POWER',    # Watt
        20: 'UA',            # J/sec-K
        21: 'WORK',          # J
        22: 'HEAT'           # J
    }
    
    # 하드코딩된 단위 데이터 (Unit_table.csv의 전체 내용)
    hardcoded_units = {
        1: {1: 'sqm', 2: 'sqft', 3: 'sqm', 4: 'sqcm', 5: 'sqin', 6: 'sqmile', 7: 'sqmm', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # AREA
        2: {1: 'mol-fr', 2: 'mol-fr', 3: 'mol-fr', 4: 'mass-fr', 5: '', 6: '', 7: '', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # COMPOSITION
        3: {1: 'kg/cum', 2: 'lb/cuft', 3: 'gm/cc', 4: 'lb/gal', 5: 'gm/cum', 6: 'gm/ml', 7: 'lb/bbl', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # DENSITY
        4: {1: 'J', 2: 'Btu', 3: 'cal', 4: 'kcal', 5: 'kWhr', 6: 'ft-lbf', 7: 'GJ', 8: 'kJ', 9: 'N-m', 10: 'MJ', 11: 'Mcal', 12: 'Gcal', 13: 'Mbtu', 14: 'MMBtu', 15: 'hp-hr', 16: 'MMkcal', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # ENERGY
        5: {1: 'kg/sec', 2: 'lb/hr', 3: 'kg/hr', 4: 'lb/sec', 5: 'Mlb/hr', 6: 'tons/day', 7: 'Mcfh', 8: 'tonne/hr', 9: 'lb/day', 10: 'kg/day', 11: 'tons/hr', 12: 'kg/min', 13: 'kg/year', 14: 'gm/min', 15: 'gm/hr', 16: 'gm/day', 17: 'Mgm/hr', 18: 'Ggm/hr', 19: 'Mgm/day', 20: 'Ggm/day', 21: 'lb/min', 22: 'MMlb/hr', 23: 'Mlb/day', 24: 'MMlb/day', 25: 'lb/year', 26: 'Mlb/year', 27: 'MMIb/year', 28: 'tons/min', 29: 'Mtons/year', 30: 'MMtons/year', 31: 'L-tons/min', 32: 'L-tons/hr', 33: 'L-tons/day', 34: 'ML-tons/year', 35: 'MML-tons/year', 36: 'ktonne/year', 37: 'kg/oper-year', 38: 'lb/oper-year', 39: 'Mlb/oper-year', 40: 'MIMIb/oper-year', 41: 'Mtons/oper-year', 42: 'MMtons/oper-year', 43: 'ML-tons/oper-year', 44: 'MML-tons/oper-year', 45: 'ktonne/oper-year', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # FLOW
        6: {1: 'kg/sec', 2: 'lb/hr', 3: 'kg/hr', 4: 'lb/sec', 5: 'Mlb/hr', 6: 'tons/day', 7: 'gm/sec', 8: 'tonne/hr', 9: 'lb/day', 10: 'kg/day', 11: 'tons/year', 12: 'tons/hr', 13: 'tonne/day', 14: 'tonne/year', 15: 'kg/min', 16: 'kg/year', 17: 'gm/min', 18: 'gm/hr', 19: 'gm/day', 20: 'Mgm/hr', 21: 'Ggm/hr', 22: 'Mgm/day', 23: 'Ggm/day', 24: 'lb/min', 25: 'MMlb/hr', 26: 'Mlb/day', 27: 'MMlb/day', 28: 'lb/year', 29: 'Mlb/year', 30: 'MMlb/year', 31: 'tons/min', 32: 'Mtons/year', 33: 'MMtons/year', 34: 'L-tons/min', 35: 'L-tons/hr', 36: 'L-tons/day', 37: 'ML-tons/year', 38: 'MML-tons/year', 39: 'ktonne/year', 40: 'tons/oper-year', 41: 'tonne/oper-year', 42: 'kg/oper-year', 43: 'lb/oper-year', 44: 'Mlb/oper-year', 45: 'MMlb/oper-year', 46: 'Mtons/oper-year', 47: 'MMtons/oper-year', 48: 'ML-tons/oper-year', 49: 'MML-tons/oper-year', 50: 'ktonne/oper-year', 51: ''},  # MASS-FLOW
        7: {1: 'kmol/sec', 2: 'lbmol/hr', 3: 'kmol/hr', 4: 'MMscfh', 5: 'MMscmh', 6: 'mol/sec', 7: 'lbmol/sec', 8: 'scmh', 9: 'bmol/day', 10: 'kmol/day', 11: 'MMscfd', 12: 'Mlscfd', 13: 'scfm', 14: 'mol/min', 15: 'kmol/khr', 16: 'kmol/Mhr', 17: 'mol/hr', 18: 'Mmol/hr', 19: 'Mlbmol/hr', 20: 'lbmol/Mhr', 21: 'lbmol/MMhr', 22: 'Mscfm', 23: 'scfh', 24: 'scfd', 25: 'ncmh', 26: 'ncmd', 27: 'ACFM', 28: 'kmol/min', 29: 'kmol/week', 30: 'kmol/month', 31: 'kmol/year', 32: 'kmol/oper-year', 33: 'lbmol/min', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # MOLE-FLOW
        8: {1: 'cum/sec', 2: 'cuft/hr', 3: 'l/min', 4: 'gal/min', 5: 'gal/hr', 6: 'bbl/day', 7: 'cum/hr', 8: 'cuft/min', 9: 'bbl/hr', 10: 'cuft/sec', 11: 'cum/day', 12: 'cum/year', 13: 'l/hr', 14: 'kbbl/day', 15: 'MMcuft/hr', 16: 'MMcuft/day', 17: 'Mcuft/day', 18: 'l/sec', 19: 'l/day', 20: 'cum/min', 21: 'kcum/sec', 22: 'kcum/hr', 23: 'kcum/day', 24: 'Mcum/sec', 25: 'Mcum/hr', 26: 'Mcum/day', 27: 'ACFM', 28: 'cuft/day', 29: 'Mcuft/min', 30: 'Mcuft/hr', 31: 'MMcuft/hr', 32: 'Mgal/min', 33: 'MMgal/min', 34: 'Mgal/hr', 35: 'MMgal/hr', 36: 'Mbbl/hr', 37: 'MMbbl/hr', 38: 'Mbbl/day', 39: 'MMbbl/day', 40: 'cum/oper-year', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # VOLUME-FLOW
        9: {1: 'kg', 2: 'lb', 3: 'kg', 4: 'gm', 5: 'ton', 6: 'Mlb', 7: 'tonne', 8: 'L-ton', 9: 'MMlb', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # MASS
        10: {1: 'Watt', 2: 'hp', 3: 'kW', 4: 'Btu/hr', 5: 'cal/sec', 6: 'ft-lbf/sec', 7: 'MIW', 8: 'GW', 9: 'MJ/hr', 10: 'kcal/hr', 11: 'Gcal/hr', 12: 'MMBtu/hr', 13: 'MBtu/hr', 14: 'Mhp', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # POWER
        11: {1: 'N/sqm', 2: 'PsIa', 3: 'atm', 4: 'lbf/sqft', 5: 'bar', 6: 'torr', 7: 'in-water', 8: 'kg/sqcm', 9: 'mmHg', 10: 'kPa', 11: 'mm-water', 12: 'mbar', 13: 'psig', 14: 'atmg', 15: 'barg', 16: 'kg/sqcmg', 17: 'lb/ft-sqsec', 18: 'kg/m-sqsec', 19: 'pa', 20: 'MiPa', 21: 'Pag', 22: 'kPag', 23: 'MPag', 24: 'mbarg', 25: 'in-Hg', 26: 'mmHg-vac', 27: 'in-Hg-vac', 28: 'in-water-60F', 29: 'in-water-vac', 30: 'in-water-60F-vac', 31: 'in-water-g', 32: 'in-water-60F-g', 33: 'mm-water-g', 34: 'mm-water-60F-g', 35: 'psi', 36: 'mm-water-60F', 37: 'bara', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # PRESSURE
        12: {1: 'K', 2: 'F', 3: 'K', 4: 'C', 5: 'R', 6: '', 7: '', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # TEMPERATURE
        13: {1: 'sec', 2: 'hr', 3: 'hr', 4: 'day', 5: 'min', 6: 'year', 7: 'month', 8: 'week', 9: 'nsec', 10: 'oper-year', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # TIME
        14: {1: 'm/sec', 2: 'ft/sec', 3: 'm/sec', 4: 'mile/hr', 5: 'km/hr', 6: 'ft/min', 7: 'mm/day', 8: 'mm/hr', 9: 'mm/day30', 10: 'in/day', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # VELOCITY
        15: {1: 'cum', 2: 'cuft', 3: 'l', 4: 'cuin', 5: 'gal', 6: 'bbl', 7: 'cc', 8: 'kcum', 9: 'Mcum', 10: 'Mcuft', 11: 'MMcuft', 12: 'ml', 13: 'kl', 14: 'MMl', 15: 'Mgal', 16: 'MMgal', 17: 'UKgal', 18: 'MUKgal', 19: 'MMUKgal', 20: 'Mbbl', 21: 'MMbbl', 22: 'kbbl', 23: 'cuyd', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # VOLUME
        16: {1: 'kmol/cum', 2: 'lbmol/cuft', 3: 'mol/cc', 4: 'lbmol/gal', 5: 'mol/l', 6: 'mmol/cc', 7: 'mmol/l', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # MOLE-DENSITY
        17: {1: 'kg/cum', 2: 'lb/cuft', 3: 'gm/cc', 4: 'lb/gal', 5: 'gm/cum', 6: 'gm/ml', 7: 'gm/l', 8: 'mg/l', 9: 'mg/cc', 10: 'mg/cum', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # MASS-DENSITY
        18: {1: 'cum/kmol', 2: 'cuft/lbmol', 3: 'cc/mol', 4: 'ml/mol', 5: 'bbl/mscf', 6: '', 7: '', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # MOLE-VOLUME
        19: {1: 'Watt', 2: 'kW', 3: 'kW', 4: 'MW', 5: 'GW', 6: '', 7: '', 8: '', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # ELEC-POWER
        20: {1: 'J/sec-K', 2: 'Btu/hr-R', 3: 'cal/sec-K', 4: 'kJ/sec-K', 5: 'kcal/sec-K', 6: 'kcal/hr-K', 7: 'Btu/hr-F', 8: 'kW/k', 9: '', 10: '', 11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # UA
        21: {1: 'J', 2: 'hp-hr', 3: 'kW-hr', 4: 'ft-lbf', 5: 'kJ', 6: 'N-m', 7: 'MJ', 8: 'Mbtu', 9: 'MMBtu', 10: 'Mcal', 11: 'Gcal', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''},  # WORK
        22: {1: 'J', 2: 'Btu', 3: 'cal', 4: 'kcal', 5: 'Mmkcal', 6: 'MMBtu', 7: 'Pcu', 8: 'MMPcu', 9: 'kJ', 10: 'GJ', 11: 'N-m', 12: 'MJ', 13: 'Mcal', 14: 'Gcal', 15: 'Mbtu', 16: 'kW-hr', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '', 28: '', 29: '', 30: '', 31: '', 32: '', 33: '', 34: '', 35: '', 36: '', 37: '', 38: '', 39: '', 40: '', 41: '', 42: '', 43: '', 44: '', 45: '', 46: '', 47: '', 48: '', 49: '', 50: '', 51: ''}  # HEAT
    }
    
    # unit_table 형태로 변환
    unit_table = {}
    for csv_col_idx, unit_type_name in csv_column_to_unit_type.items():
        if csv_col_idx in hardcoded_units:
            unit_table[csv_col_idx] = {
                'unit_type': unit_type_name,
                'units': {idx: unit for idx, unit in hardcoded_units[csv_col_idx].items() if unit.strip()}
            }
    
    return unit_table

#======================================================================
# SI Unit Conversion System
#======================================================================

# get_si_base_units 함수는 unit_converter.py로 이동됨

# Unit conversion factors moved to unit_converter.py

# 중복된 변환 함수들은 unit_converter.py로 이동됨

# convert_to_si_units 함수는 unit_converter.py의 convert_to_si_units로 대체됨

def convert_multiple_values_to_si(values_dict, units_dict, unit_types_dict):
    """
    여러 값들을 한 번에 SI 단위로 변환하는 함수
    
    Parameters:
    -----------
    values_dict : dict
        변환할 값들의 딕셔너리 {parameter_name: value}
    units_dict : dict
        각 값의 단위 딕셔너리 {parameter_name: unit}
    unit_types_dict : dict
        각 값의 물리량 타입 딕셔너리 {parameter_name: unit_type}
    
    Returns:
    --------
    dict : {parameter_name: (converted_value, si_unit)}
    """
    converted_results = {}
    
    for param_name, value in values_dict.items():
        if param_name in units_dict and param_name in unit_types_dict:
            try:
                from_unit = units_dict[param_name]
                unit_type = unit_types_dict[param_name]
                converted_value, si_unit = convert_to_si_units(value, from_unit, unit_type)
                converted_results[param_name] = (converted_value, si_unit)
            except Exception as e:
                print(f"Warning: Failed to convert {param_name}: {str(e)}")
                converted_results[param_name] = (value, units_dict[param_name])
        else:
            print(f"Warning: Missing unit information for {param_name}")
            converted_results[param_name] = (value, "unknown")
    
    return converted_results



#======================================================================
# Unit Table Functions
#======================================================================

def get_unit_by_index(unit_table, physical_quantity_index, unit_of_measure_index):
    """
    특정 Physical Quantity 인덱스와 Unit of Measure 인덱스로 unit 값을 가져오는 함수
    """
    if physical_quantity_index in unit_table and unit_of_measure_index in unit_table[physical_quantity_index]['units']:
        return unit_table[physical_quantity_index]['units'][unit_of_measure_index]
    return None

def get_units_by_physical_quantity(unit_table, physical_quantity_index):
    """
    특정 Physical Quantity 인덱스의 모든 unit들을 가져오는 함수
    """
    if physical_quantity_index in unit_table:
        return unit_table[physical_quantity_index]['units']
    return {}

def get_unit_type_by_physical_quantity(unit_table, physical_quantity_index):
    """
    특정 Physical Quantity 인덱스에 해당하는 unit_type 이름을 가져오는 함수
    """
    if physical_quantity_index in unit_table:
        return unit_table[physical_quantity_index]['unit_type']
    return None

def get_physical_quantity_by_unit_type(unit_table, unit_type_name):
    """
    특정 unit_type 이름에 해당하는 Physical Quantity 인덱스를 가져오는 함수
    """
    for physical_quantity_idx, data in unit_table.items():
        if data['unit_type'] == unit_type_name:
            return physical_quantity_idx
    return None

#======================================================================
# Main Execution
#======================================================================

# 하드코딩된 단위 테이블 사용
# 하드코딩된 단위 테이블 사용
    # unit_table은 이제 unit_converter.py에서 직접 사용됨

# Unit table loaded successfully

# 단위 테이블은 이제 unit_converter.py에서 관리됨

# Detecting unit sets from Aspen Plus...
units_spinner = Spinner('Detecting unit sets')
units_spinner.start()

# 모든 unit set들 감지
units_sets = get_units_sets(Application)
units_spinner.stop('Unit sets detected successfully!')

# 단위 세트 요약 출력
print_units_sets_summary(units_sets)

# 현재 사용 중인 Unit Set 감지
current_unit_set = get_current_unit_set(Application)

#======================================================================
# Pressure-driven equipment cost estimation wrapper
#======================================================================

def calculate_pressure_device_costs(material: str = 'CS', target_year: int = 2024, material_overrides: dict = None):
    # 2024년 CEPCI 인덱스 설정
    target_index = 800.0  # 2024년 추정값
    cepci = CEPCIOptions(target_index=target_index)
    
    print(f"CEPCI 설정: 기준년도 2017 (인덱스: 567.5) → 목표년도 {target_year} (인덱스: {target_index})")
    
    return calculate_pressure_device_costs_auto(
        Application,
        block_info,
        current_unit_set,
        material=material,
        cepci=cepci,
        material_overrides=material_overrides,
    )


#======================================================================
# Run cost calculation and print results
#======================================================================

try:
    # 출력 상세 수준 설정: 0=조용, 1=보통, 2=상세
    VERBOSITY = 1
    try:
        vv = input("출력 수준을 선택하세요 (0=조용, 1=보통, 2=상세, 기본=1): ").strip()
        if vv in ('0','1','2'):
            VERBOSITY = int(vv)
    except Exception:
        pass

    # 장치별 세부 출력 토글 기능 제거 (단순화)

    register_default_correlations()
    
    # 캐시 초기화
    clear_aspen_cache()
    
    # 1) Preview (세션 불러오기 옵션)
    session: Optional[PreviewSession] = None
    # 오버라이드 변수 초기화
    material_overrides = {}
    type_overrides = {}
    subtype_overrides = {}
    session = None
    
    load_choice = input("기존 세션을 불러오시겠습니까? (y/n): ").strip().lower()
    if load_choice == 'y':
        try:
            # 현재 폴더의 .pkl 세션을 최근 수정순으로 나열 후 선택
            session_files_all = [f for f in os.listdir(current_dir) if f.lower().endswith('.pkl')]
            session_files = sorted(session_files_all, key=lambda f: os.path.getmtime(os.path.join(current_dir, f)), reverse=True)
            if not session_files:
                print("불러올 .pkl 세션이 없습니다.")
                preview = preview_pressure_devices_auto(Application, block_info, current_unit_set)
            else:
                print("\n감지된 세션(.pkl) 목록:")
                for i, fname in enumerate(session_files, 1):
                    print(f"  {i}. {fname}")
                while True:
                    try:
                        choice = int(input("불러올 세션 번호를 선택하세요: ").strip())
                        if 1 <= choice <= len(session_files):
                            load_path = os.path.join(current_dir, session_files[choice - 1])
                            break
                        else:
                            print("잘못된 번호입니다. 다시 입력해주세요.")
                    except ValueError:
                        print("숫자를 입력해주세요.")
                session = PreviewSession.load(load_path)
                print(f"✅ 세션을 불러왔습니다: {session_files[choice - 1]}")
                
                if os.path.basename(session.aspen_file) != os.path.basename(aspen_Path):
                    print(f"경고: 세션 파일({session.aspen_file})이 현재 파일({aspen_Path})과 다릅니다. 세션 데이터를 그대로 사용합니다.")
                
                preview = session.apply_overrides_to_preview()
                material_overrides = dict(session.material_overrides)
                type_overrides = dict(session.type_overrides)
                subtype_overrides = dict(session.subtype_overrides)
                
                print(f"📋 세션에서 불러온 오버라이드:")
                print(f"   - 재질 오버라이드: {len(material_overrides)}개 - {material_overrides}")
                print(f"   - 타입 오버라이드: {len(type_overrides)}개 - {type_overrides}")
                print(f"   - 서브타입 오버라이드: {len(subtype_overrides)}개 - {subtype_overrides}")
        except Exception as e:
            print(f"세션 불러오기 실패: {e}")
            preview = preview_pressure_devices_auto(Application, block_info, current_unit_set)
            material_overrides = {}
            type_overrides = {}
            subtype_overrides = {}
    else:
        preview = preview_pressure_devices_auto(Application, block_info, current_unit_set)
    
    # 캐시 통계 출력
    cache_stats = get_cache_stats()
    # 캐시 통계 출력
    
    # 프리뷰 결과 출력 (모듈 함수 사용)
    power_unit = None
    pressure_unit = None
    flow_unit = None
    heat_unit = None
    temperature_unit = None
    if current_unit_set:
                power_unit = get_unit_type_value(Application, current_unit_set, 'POWER')
                pressure_unit = get_unit_type_value(Application, current_unit_set, 'PRESSURE')
                flow_unit = get_unit_type_value(Application, current_unit_set, 'VOLUME-FLOW')
                heat_unit = get_unit_type_value(Application, current_unit_set, 'HEAT')
                temperature_unit = get_unit_type_value(Application, current_unit_set, 'TEMPERATURE')
    
    # 세션을 불러왔다면 이미 preview가 설정되어 있음
    # 세션을 불러오지 않았다면 새로 생성
    if session is None:
        preview = preview_pressure_devices_auto(Application, block_info, current_unit_set)
    
    if VERBOSITY >= 1:
        # 세션 불러오기 상태 확인
        if session is not None:
            print(f"\n🔍 세션 불러오기 확인:")
            print(f"   - 세션 파일: {session.aspen_file}")
            print(f"   - 오버라이드 적용된 프리뷰 항목 수: {len(preview)}")
            # 첫 번째 장치의 재질 정보 확인
            if preview:
                first_device = preview[0]
                print(f"   - 첫 번째 장치 ({first_device.get('name', 'N/A')}): 재질={first_device.get('material', 'N/A')}, 타입={first_device.get('selected_type', 'N/A')}")
        
        print_preview_results(preview, Application, power_unit, pressure_unit)

    # 통합 장치 오버라이드 UI
    print("\n" + "="*60)
    print("EQUIPMENT DESIGN OVERRIDES")
    print("="*60)
    
    # 모든 장치를 하나의 리스트로 통합 (압력 장치만)
    all_devices = []
    for device in preview:
        device['device_type'] = 'pressure'
        all_devices.append(device)
    
    # 장치 이름으로 정렬
    all_devices.sort(key=lambda x: x['name'])
    
    # 통합 오버라이드 루프
    material_overrides = material_overrides if 'material_overrides' in locals() else {}
    type_overrides = type_overrides if 'type_overrides' in locals() else {}
    subtype_overrides = subtype_overrides if 'subtype_overrides' in locals() else {}
    
    while True:
        print("\n사용 가능한 장치 목록:")
        for i, device in enumerate(all_devices, 1):
            device_type = device.get('device_type', 'unknown')
            if device_type == 'pressure':
                cat = device.get('category', 'Unknown')
                print(f"  {i:2d}. {device['name']:20s} ({cat}) - 압력 장치")
            elif device_type == 'heat_exchanger':
                print(f"  {i:2d}. {device['name']:20s} (HeatExchanger) - 열교환기")
        
        ans = input("\n설계/재질을 변경할 장치 이름을 입력하세요 (없으면 엔터): ").strip()
        if not ans:
            break
        
        # 해당 장치 찾기
        device_info = None
        for device in all_devices:
            if device['name'] == ans:
                device_info = device
                break
        
        if not device_info:
            print(f"장치 '{ans}'를 찾을 수 없습니다.")
            continue
        
        device_type = device_info.get('device_type')
        
        if device_type == 'pressure':
            # 압력 장치 오버라이드 로직
            print(f"\n선택된 장치: {ans} ({device_info['category']}) - 압력 장치")
            print(f"현재 타입: {device_info.get('selected_type', 'N/A')}")
            print(f"현재 세부 타입: {device_info.get('selected_subtype', 'N/A')}")
            
            # 선택 가능한 타입과 세부 타입 표시 (압력 조건 고려)
            from equipment_costs import get_device_type_options
            type_options = get_device_type_options(device_info['category'])
            
            # 압력 조건에 따른 타입 필터링
            inlet_bar = device_info.get('inlet_bar')
            outlet_bar = device_info.get('outlet_bar')
            filtered_type_options = {}
            
            if inlet_bar is not None and outlet_bar is not None:
                pressure_rise = outlet_bar - inlet_bar
                
                for main_type, subtypes in type_options.items():
                    # 물리적 제약 조건 적용
                    if main_type == 'fan' and outlet_bar > 1.17325:  # 0.16 barg = 1.17325 bara
                        continue  # 팬은 출구 압력이 0.16 barg (1.17325 bara) 이하일 때만 가능
                    elif main_type == 'turbine' and pressure_rise >= 0:
                        continue  # 터빈은 압력 하강(음수)일 때만 가능
                    elif main_type == 'compressor' and pressure_rise <= 0:
                        continue  # 압축기는 압력 상승(양수)일 때만 가능
                    
                    filtered_type_options[main_type] = subtypes
            
            if filtered_type_options:
                print("\n물리적 조건을 고려한 사용 가능한 타입과 세부 타입:")
                for main_type, subtypes in filtered_type_options.items():
                    print(f"  {main_type}: {', '.join(subtypes)}")
                
                # 제한된 타입이 있는지 확인
                if not filtered_type_options:
                    print(f"\n경고: {ans}의 압력 조건 (입구: {inlet_bar} bar, 출구: {outlet_bar} bar)에 맞는 타입이 없습니다.")
                    print("기본 제안 타입을 사용합니다.")
                else:
                    # 타입 변경
                    type_input = input("\n타입을 변경하시겠습니까? (y/n): ").strip().lower()
                    if type_input == 'y':
                        print("\n사용 가능한 타입:")
                        main_types = list(filtered_type_options.keys())
                        for i, t in enumerate(main_types, 1):
                            print(f"  {i}. {t}")
                        
                        while True:
                            try:
                                type_choice = int(input("타입 번호를 선택하세요: ").strip())
                                if 1 <= type_choice <= len(main_types):
                                    selected_type = main_types[type_choice - 1]
                                    type_overrides[ans] = selected_type
                                    print(f"{ans}의 타입이 {selected_type}로 변경되었습니다.")
                                    
                                    # 세부 타입 선택
                                    available_subtypes = filtered_type_options[selected_type]
                                    print(f"\n사용 가능한 세부 타입:")
                                    for i, st in enumerate(available_subtypes, 1):
                                        print(f"  {i}. {st}")
                                    
                                    while True:
                                        try:
                                            subtype_choice = int(input("세부 타입 번호를 선택하세요: ").strip())
                                            if 1 <= subtype_choice <= len(available_subtypes):
                                                selected_subtype = available_subtypes[subtype_choice - 1]
                                                subtype_overrides[ans] = selected_subtype
                                                print(f"{ans}의 세부 타입이 {selected_subtype}로 변경되었습니다.")
                                                break
                                            else:
                                                print("잘못된 번호입니다. 다시 입력해주세요.")
                                        except ValueError:
                                            print("숫자를 입력해주세요.")
                                    break
                                else:
                                    print("잘못된 번호입니다. 다시 입력해주세요.")
                            except ValueError:
                                print("숫자를 입력해주세요.")
            else:
                print(f"\n경고: {ans}의 압력 조건에 맞는 타입이 없습니다. 기본 제안 타입을 사용합니다.")
            
            # 재질 변경
            valid_materials = ['CS', 'SS', 'Ni', 'Cu', 'Cl', 'Ti', 'Fiberglass']
            while True:
                mat = input("변경할 재질을 입력하세요 (예: CS, SS, Ni, Cl, Ti, Fiberglass, 없으면 엔터): ").strip()
                if not mat:  # 엔터만 입력한 경우
                    break
                elif mat in valid_materials:
                    material_overrides[ans] = mat
                    print(f"{ans}의 재질이 {mat}로 변경되었습니다.")
                    break
                else:
                    print(f"잘못된 재질입니다. 사용 가능한 재질: {', '.join(valid_materials)}")
                    print("다시 입력해주세요.")
        
        elif device_type == 'heat_exchanger':
            # 열교환기 오버라이드 로직
            print(f"\n선택된 장치: {ans} - 열교환기")
            
            # 타입 선택
            hx_types = [
                'fixed_tube','floating_head','bayonet','kettle_reboiler','double_pipe','multiple_pipe',
                'scraped_wall','air_cooler','teflon_tube','spiral_tube_shell','spiral_plate','flat_plate'
            ]
            print("사용 가능한 HX 타입:")
            for i,t in enumerate(hx_types,1):
                print(f"  {i}. {t}")
            
            while True:
                try:
                    sel = int(input("HX 타입 번호: ").strip())
                    if 1 <= sel <= len(hx_types):
                        hx_type = hx_types[sel-1]
                        break
                    else:
                        print("잘못된 번호입니다.")
                except ValueError:
                    print("숫자를 입력해주세요.")
            
            # 재질 옵션 표시
            from equipment_costs import get_hx_material_options
            guide = get_hx_material_options(hx_type)  # {shell:[], tube:[], notes:[]}
            notes = guide.get('notes', [])
            for m in notes:
                print("- "+m)
            shell_choice = None
            tube_choice = None
            if guide.get('shell') and guide['shell'][0] != '(선택 불가)' and guide['shell'][0] != '(내부 고정)':
                print("가능한 쉘 재질:", ', '.join(guide['shell']))
                shell_choice = input("쉘 재질 입력: ").strip()
            elif guide.get('shell') and guide['shell'][0] == '(내부 고정)':
                shell_choice = 'CS'
            if guide.get('tube') and guide['tube'][0] != 'Teflon (fixed)':
                print("가능한 튜브 재질:", ', '.join(guide['tube']))
                tube_choice = input("튜브 재질 입력: ").strip()
            
            # U/LMTD/Area 오버라이드 입력
            def _tryfloat(s):
                try:
                    return float(s)
                except ValueError:
                    return None
            
            q_w = device_info.get('q_w')
            u = device_info.get('u_W_m2K')
            lmtd = device_info.get('lmtd_K')
            area = device_info.get('area_m2')
            
            if input("U 값을 변경하시겠습니까? (y/n): ").strip().lower()=='y':
                u = _tryfloat(input("U [W/m2-K]: ").strip()) or u
            if input("LMTD 값을 변경하시겠습니까? (y/n): ").strip().lower()=='y':
                lmtd = _tryfloat(input("LMTD [K]: ").strip()) or lmtd
            if input("면적(A) 값을 직접 지정하시겠습니까? (y/n): ").strip().lower()=='y':
                area = _tryfloat(input("Area [m2]: ").strip()) or area
            
            # 비용 미리보기 계산 (현재는 미구현)
            # costs = estimate_heat_exchanger_cost(...)
        
        # 변경사항이 있으면 프리뷰 다시 표시
        if ans in material_overrides or ans in type_overrides or ans in subtype_overrides:
            print("\n" + "="*60)
            print("UPDATED PREVIEW: ALL EQUIPMENT")
            print("="*60)
            
            # 업데이트된 프리뷰 데이터 생성
            updated_preview = []
            for p in preview:
                updated_p = p.copy()
                device_name = p['name']
                
                # 모든 오버라이드 적용 (현재 장치와 이전에 변경한 장치들 모두)
                if device_name in material_overrides:
                    updated_p['material'] = material_overrides[device_name]
                if device_name in type_overrides:
                    updated_p['selected_type'] = type_overrides[device_name]
                if device_name in subtype_overrides:
                    updated_p['selected_subtype'] = subtype_overrides[device_name]
                    
                updated_preview.append(updated_p)
            
            # 업데이트된 프리뷰 출력
            print_preview_results(updated_preview, Application, power_unit, pressure_unit)
            
            # 세션 업데이트/생성
            if session is None:
                session = PreviewSession(
                    aspen_file=aspen_Path,
                    current_unit_set=current_unit_set,
                    block_info=block_info,
                    preview=preview,
                    material_overrides=material_overrides,
                    type_overrides=type_overrides,
                    subtype_overrides=subtype_overrides,
                )
            else:
                session.preview = preview
                session.material_overrides = material_overrides
                session.type_overrides = type_overrides
                session.subtype_overrides = subtype_overrides

    # 2) Build pre-extracted dict from block_info (freeze values)
    pre_extracted = {}
    
    # 압력 장치 카테고리 정의
    pressure_cats = {'Pump', 'Compr', 'MCompr'}
    
    # 압력 장치 데이터를 딕셔너리로 변환 (빠른 검색을 위해)
    pressure_devices_dict = {}
    for device in all_devices:
        pressure_devices_dict[device['name']] = device
    
    # block_info의 모든 장치에 대해 pre_extracted 생성
    for name, cat in block_info.items():
        if cat in pressure_cats:
            # 압력 장치 처리
            device_data = pressure_devices_dict.get(name)
            
            if device_data:
                # all_devices에서 데이터 추출
                pre_extracted[name] = {
                    'power_kilowatt': device_data.get('power_kilowatt'),
                    'volumetric_flow_m3_s': device_data.get('volumetric_flow_m3_s'),
                    'inlet_bar': device_data.get('inlet_bar'),
                    'outlet_bar': device_data.get('outlet_bar'),
                    'pressure_delta_bar': device_data.get('pressure_delta_bar'),
                    'stage_data': device_data.get('stage_data'),
                }
            else:
                # all_devices에 없는 경우 - 이는 오류 상황
                raise ValueError(f"압력 장치 '{name}' ({cat})가 all_devices에 없습니다. 프리뷰 함수에 문제가 있습니다.")
    # 저장 옵션
    save_choice = input("현재 세션을 저장하시겠습니까? (y/n): ").strip().lower()
    if save_choice == 'y':
        try:
            save_path = input("저장할 파일 경로(확장자 생략 가능, 기본 .pkl): ").strip()
            if not save_path.lower().endswith('.pkl'):
                save_path = save_path + '.pkl'
            if session is None:
                session = PreviewSession(
                    aspen_file=aspen_Path,
                    current_unit_set=current_unit_set,
                    block_info=block_info,
                    preview=preview,
                    material_overrides=material_overrides,
                    type_overrides=type_overrides,
                    subtype_overrides=subtype_overrides,
                )
            session.save(save_path)
            print(f"세션이 저장되었습니다: {save_path}")
        except Exception as e:
            print(f"세션 저장 실패: {e}")

    confirm = input("\n위 데이터/재질로 비용 계산을 진행할까요? (y/n): ").strip().lower()
    if confirm != 'y':
        print("사용자에 의해 계산이 취소되었습니다.")
        raise SystemExit(0)

    # 계산에 사용될 오버라이드 정보 출력
    print(f"\n🔧 계산에 적용될 오버라이드:")
    print(f"   - 재질 오버라이드: {material_overrides}")
    print(f"   - 타입 오버라이드: {type_overrides}")
    print(f"   - 서브타입 오버라이드: {subtype_overrides}")

    # 4) Run using pre-extracted data (no further COM reads)
    # 변수 초기화
    pressure_device_costs = []
    pressure_device_totals = {"purchased": 0.0, "purchased_adj": 0.0, "bare_module": 0.0, "installed": 0.0}
    
    # 상세 모드가 아니라면 내부 디버그 출력 억제
    if VERBOSITY >= 2:
        pressure_device_costs, pressure_device_totals = calculate_pressure_device_costs_with_data(
            pre_extracted=pre_extracted,
            block_info=block_info,
            material='CS',
            cepci=CEPCIOptions(target_index=800.0),  # 2024년 CEPCI 인덱스
            material_overrides=material_overrides,
            type_overrides=type_overrides,
            subtype_overrides=subtype_overrides,
        )
        
        # 열교환기 비용 계산
        from equipment_costs import calculate_heat_exchanger_costs_with_data
        heat_exchanger_costs, heat_exchanger_totals = calculate_heat_exchanger_costs_with_data(
            pre_extracted=pre_extracted,
            block_info=block_info,
            material='CS',
            cepci=CEPCIOptions(target_index=800.0),
            material_overrides=material_overrides,
            type_overrides=type_overrides,
            subtype_overrides=subtype_overrides,
        )
    else:
        with open(os.devnull, 'w') as _null, contextlib.redirect_stdout(_null):
            pressure_device_costs, pressure_device_totals = calculate_pressure_device_costs_with_data(
                pre_extracted=pre_extracted,
                block_info=block_info,
                material='CS',
                cepci=CEPCIOptions(target_index=800.0),  # 2024년 CEPCI 인덱스
                material_overrides=material_overrides,
                type_overrides=type_overrides,
                subtype_overrides=subtype_overrides,
            )

        # 카테고리별 세부 출력 기능 제거됨
    
    if pressure_device_costs:
        print("\n" + "="*60)
        print("CALCULATED PRESSURE DEVICE COSTS")
        print("="*60)
        for item in pressure_device_costs:
            name = item.get('name')
            dtype = item.get('type')
            bare = item.get('bare_module', 0.0)
            
            if dtype == 'error':
                error_msg = item.get('error', 'Unknown error')
                print(f"{name} (error): {error_msg}")
            else:
                print(f"{name} ({dtype}): Bare Module Cost = {bare:,.2f} USD")
        print(f"\nTotal Bare Module Cost for Pressure Devices: {pressure_device_totals.get('bare_module', 0.0):,.2f} USD")
        print("Note: Bare Module Cost includes installation costs")
        print("="*60)
    else:
        print("No pressure device costs calculated.")
    
    # 전체 총 비용 출력
    total_cost = pressure_device_totals.get('bare_module', 0.0)
    
    print("\n" + "="*60)
    print("TOTAL EQUIPMENT COSTS")
    print("="*60)
    print(f"Total Bare Module Cost: {total_cost:,.2f} USD")
    print("Note: Bare Module Cost includes installation costs")
    print("="*60)
    
except Exception as e:
    print(f"Error during pressure device cost calculation/printing: {e}")

    