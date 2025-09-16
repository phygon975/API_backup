"""
통합 단위 변환 시스템

이 모듈은 Aspen Plus와 호환되는 모든 단위 변환을 통합적으로 처리합니다.
TEA_machine.py의 하드코딩된 단위 데이터를 활용하여 포괄적인 변환을 제공합니다.
"""

from typing import Dict, Optional, Union
import math


class UnitConverter:
    """
    통합 단위 변환 클래스
    
    Aspen Plus의 모든 단위를 SI 단위로 변환하거나, SI 단위에서 목표 단위로 변환합니다.
    """
    
    def __init__(self):
        """단위 변환 시스템 초기화"""
        self._si_base_units = self._get_si_base_units()
        self._conversion_factors = self._get_unit_conversion_factors()
        self._unit_table = self._get_hardcoded_unit_table()
    
    def _get_si_base_units(self) -> Dict[str, str]:
        """각 물리량별 SI 기준 단위 정의"""
        return {
            'AREA': 'sqm',           # 제곱미터
            'COMPOSITION': 'mol-fr', # 몰분율 (무차원)
            'DENSITY': 'kg/cum',     # kg/m³
            'ENERGY': 'J',           # 줄
            'FLOW': 'kg/sec',        # kg/s
            'MASS-FLOW': 'kg/sec',   # kg/s
            'MOLE-FLOW': 'kmol/sec', # kmol/s
            'VOLUME-FLOW': 'cum/sec', # m³/s
            'MASS': 'kg',            # 킬로그램
            'POWER': 'Watt',         # 와트
            'PRESSURE': 'N/sqm',     # 파스칼 (N/m²)
            'TEMPERATURE': 'K',      # 켈빈
            'TIME': 'sec',           # 초
            'VELOCITY': 'm/sec',     # m/s
            'VOLUME': 'cum',         # m³
            'MOLE-DENSITY': 'kmol/cum', # kmol/m³
            'MASS-DENSITY': 'kg/cum',   # kg/m³
            'MOLE-VOLUME': 'cum/kmol',  # m³/kmol
            'ELEC-POWER': 'Watt',    # 와트
            'UA': 'J/sec-K',         # J/(s·K)
            'WORK': 'J',             # 줄
            'HEAT': 'J'              # 줄
        }
    
    def _get_hardcoded_unit_table(self) -> Dict[int, Dict[str, Union[str, Dict[int, str]]]]:
        """TEA_machine.py의 하드코딩된 단위 테이블"""
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
        
        # 하드코딩된 단위 데이터 (TEA_machine.py에서 가져옴)
        hardcoded_units = {
            1: {1: 'sqm', 2: 'sqft', 3: 'sqm', 4: 'sqcm', 5: 'sqin', 6: 'sqmile', 7: 'sqmm'},
            2: {1: 'mol-fr', 2: 'mol-fr', 3: 'mol-fr', 4: 'mass-fr'},
            3: {1: 'kg/cum', 2: 'lb/cuft', 3: 'gm/cc', 4: 'lb/gal', 5: 'gm/cum', 6: 'gm/ml', 7: 'lb/bbl'},
            4: {1: 'J', 2: 'Btu', 3: 'cal', 4: 'kcal', 5: 'kWhr', 6: 'ft-lbf', 7: 'GJ', 8: 'kJ', 9: 'N-m', 10: 'MJ', 11: 'Mcal', 12: 'Gcal', 13: 'Mbtu', 14: 'MMBtu', 15: 'hp-hr', 16: 'MMkcal'},
            5: {1: 'kg/sec', 2: 'lb/hr', 3: 'kg/hr', 4: 'lb/sec', 5: 'Mlb/hr', 6: 'tons/day', 7: 'Mcfh', 8: 'tonne/hr', 9: 'lb/day', 10: 'kg/day', 11: 'tons/hr', 12: 'kg/min', 13: 'kg/year', 14: 'gm/min', 15: 'gm/hr', 16: 'gm/day', 17: 'Mgm/hr', 18: 'Ggm/hr', 19: 'Mgm/day', 20: 'Ggm/day', 21: 'lb/min', 22: 'MMlb/hr', 23: 'Mlb/day', 24: 'MMlb/day', 25: 'lb/year', 26: 'Mlb/year', 27: 'MMIb/year', 28: 'tons/min', 29: 'Mtons/year', 30: 'MMtons/year', 31: 'L-tons/min', 32: 'L-tons/hr', 33: 'L-tons/day', 34: 'ML-tons/year', 35: 'MML-tons/year', 36: 'ktonne/year', 37: 'kg/oper-year', 38: 'lb/oper-year', 39: 'Mlb/oper-year', 40: 'MIMIb/oper-year', 41: 'Mtons/oper-year', 42: 'MMtons/oper-year', 43: 'ML-tons/oper-year', 44: 'MML-tons/oper-year', 45: 'ktonne/oper-year'},
            6: {1: 'kg/sec', 2: 'lb/hr', 3: 'kg/hr', 4: 'lb/sec', 5: 'Mlb/hr', 6: 'tons/day', 7: 'gm/sec', 8: 'tonne/hr', 9: 'lb/day', 10: 'kg/day', 11: 'tons/year', 12: 'tons/hr', 13: 'tonne/day', 14: 'tonne/year', 15: 'kg/min', 16: 'kg/year', 17: 'gm/min', 18: 'gm/hr', 19: 'gm/day', 20: 'Mgm/hr', 21: 'Ggm/hr', 22: 'Mgm/day', 23: 'Ggm/day', 24: 'lb/min', 25: 'MMlb/hr', 26: 'Mlb/day', 27: 'MMlb/day', 28: 'lb/year', 29: 'Mlb/year', 30: 'MMlb/year', 31: 'tons/min', 32: 'Mtons/year', 33: 'MMtons/year', 34: 'L-tons/min', 35: 'L-tons/hr', 36: 'L-tons/day', 37: 'ML-tons/year', 38: 'MML-tons/year', 39: 'ktonne/year', 40: 'tons/oper-year', 41: 'tonne/oper-year', 42: 'kg/oper-year', 43: 'lb/oper-year', 44: 'Mlb/oper-year', 45: 'MMlb/oper-year', 46: 'Mtons/oper-year', 47: 'MMtons/oper-year', 48: 'ML-tons/oper-year', 49: 'MML-tons/oper-year', 50: 'ktonne/oper-year'},
            7: {1: 'kmol/sec', 2: 'lbmol/hr', 3: 'kmol/hr', 4: 'MMscfh', 5: 'MMscmh', 6: 'mol/sec', 7: 'lbmol/sec', 8: 'scmh', 9: 'bmol/day', 10: 'kmol/day', 11: 'MMscfd', 12: 'Mlscfd', 13: 'scfm', 14: 'mol/min', 15: 'kmol/khr', 16: 'kmol/Mhr', 17: 'mol/hr', 18: 'Mmol/hr', 19: 'Mlbmol/hr', 20: 'lbmol/Mhr', 21: 'lbmol/MMhr', 22: 'Mscfm', 23: 'scfh', 24: 'scfd', 25: 'ncmh', 26: 'ncmd', 27: 'ACFM', 28: 'kmol/min', 29: 'kmol/week', 30: 'kmol/month', 31: 'kmol/year', 32: 'kmol/oper-year', 33: 'lbmol/min'},
            8: {1: 'cum/sec', 2: 'm3/s', 3: 'm^3/s', 4: 'cuft/hr', 5: 'l/min', 6: 'gal/min', 7: 'gal/hr', 8: 'bbl/day', 9: 'cum/hr', 10: 'm3/h', 11: 'm^3/h', 12: 'cuft/min', 13: 'bbl/hr', 14: 'cuft/sec', 15: 'cum/day', 16: 'cum/year', 17: 'l/hr', 18: 'kbbl/day', 19: 'MMcuft/hr', 20: 'MMcuft/day', 21: 'Mcuft/day', 22: 'l/sec', 23: 'l/day', 24: 'cum/min', 25: 'kcum/sec', 26: 'kcum/hr', 27: 'kcum/day', 28: 'Mcum/sec', 29: 'Mcum/hr', 30: 'Mcum/day', 31: 'ACFM', 32: 'cuft/day', 33: 'Mcuft/min', 34: 'Mcuft/hr', 35: 'MMcuft/hr', 36: 'Mgal/min', 37: 'MMgal/min', 38: 'Mgal/hr', 39: 'MMgal/hr', 40: 'Mbbl/hr', 41: 'MMbbl/hr', 42: 'Mbbl/day', 43: 'MMbbl/day', 44: 'cum/oper-year'},
            9: {1: 'kg', 2: 'lb', 3: 'kg', 4: 'gm', 5: 'ton', 6: 'Mlb', 7: 'tonne', 8: 'L-ton', 9: 'MMlb'},
            10: {1: 'Watt', 2: 'W', 3: 'hp', 4: 'kW', 5: 'Btu/hr', 6: 'cal/sec', 7: 'ft-lbf/sec', 8: 'MIW', 9: 'GW', 10: 'MJ/hr', 11: 'kcal/hr', 12: 'Gcal/hr', 13: 'MMBtu/hr', 14: 'MBtu/hr', 15: 'Mhp'},
            11: {1: 'N/sqm', 2: 'PsIa', 3: 'atm', 4: 'lbf/sqft', 5: 'bar', 6: 'torr', 7: 'in-water', 8: 'kg/sqcm', 9: 'mmHg', 10: 'kPa', 11: 'mm-water', 12: 'mbar', 13: 'psig', 14: 'atmg', 15: 'barg', 16: 'kg/sqcmg', 17: 'lb/ft-sqsec', 18: 'kg/m-sqsec', 19: 'pa', 20: 'MiPa', 21: 'Pag', 22: 'kPag', 23: 'MPag', 24: 'mbarg', 25: 'in-Hg', 26: 'mmHg-vac', 27: 'in-Hg-vac', 28: 'in-water-60F', 29: 'in-water-vac', 30: 'in-water-60F-vac', 31: 'in-water-g', 32: 'in-water-60F-g', 33: 'mm-water-g', 34: 'mm-water-60F-g', 35: 'psi', 36: 'mm-water-60F', 37: 'bara'},
            12: {1: 'K', 2: 'F', 3: 'K', 4: 'C', 5: 'R'},
            13: {1: 'sec', 2: 'hr', 3: 'hr', 4: 'day', 5: 'min', 6: 'year', 7: 'month', 8: 'week', 9: 'nsec', 10: 'oper-year'},
            14: {1: 'm/sec', 2: 'ft/sec', 3: 'm/sec', 4: 'mile/hr', 5: 'km/hr', 6: 'ft/min', 7: 'mm/day', 8: 'mm/hr', 9: 'mm/day30', 10: 'in/day'},
            15: {1: 'cum', 2: 'cuft', 3: 'l', 4: 'cuin', 5: 'gal', 6: 'bbl', 7: 'cc', 8: 'kcum', 9: 'Mcum', 10: 'Mcuft', 11: 'MMcuft', 12: 'ml', 13: 'kl', 14: 'MMl', 15: 'Mgal', 16: 'MMgal', 17: 'UKgal', 18: 'MUKgal', 19: 'MMUKgal', 20: 'Mbbl', 21: 'MMbbl', 22: 'kbbl', 23: 'cuyd'},
            16: {1: 'kmol/cum', 2: 'lbmol/cuft', 3: 'mol/cc', 4: 'lbmol/gal', 5: 'mol/l', 6: 'mmol/cc', 7: 'mmol/l'},
            17: {1: 'kg/cum', 2: 'lb/cuft', 3: 'gm/cc', 4: 'lb/gal', 5: 'gm/cum', 6: 'gm/ml', 7: 'gm/l', 8: 'mg/l', 9: 'mg/cc', 10: 'mg/cum'},
            18: {1: 'cum/kmol', 2: 'cuft/lbmol', 3: 'cc/mol', 4: 'ml/mol', 5: 'bbl/mscf'},
            19: {1: 'Watt', 2: 'kW', 3: 'kW', 4: 'MW', 5: 'GW'},
            20: {1: 'J/sec-K', 2: 'Btu/hr-R', 3: 'cal/sec-K', 4: 'kJ/sec-K', 5: 'kcal/sec-K', 6: 'kcal/hr-K', 7: 'Btu/hr-F', 8: 'kW/k'},
            21: {1: 'J', 2: 'hp-hr', 3: 'kW-hr', 4: 'ft-lbf', 5: 'kJ', 6: 'N-m', 7: 'MJ', 8: 'Mbtu', 9: 'MMBtu', 10: 'Mcal', 11: 'Gcal'},
            22: {1: 'J', 2: 'Btu', 3: 'cal', 4: 'kcal', 5: 'Mmkcal', 6: 'MMBtu', 7: 'Pcu', 8: 'MMPcu', 9: 'kJ', 10: 'GJ', 11: 'N-m', 12: 'MJ', 13: 'Mcal', 14: 'Gcal', 15: 'Mbtu', 16: 'kW-hr'}
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
    
    def _get_unit_conversion_factors(self) -> Dict[str, Union[float, str]]:
        """TEA_machine.py의 포괄적인 변환 계수"""
        return {
            # AREA (sqm 기준)
            'sqm': 1.0, 'sqft': 0.092903, 'sqcm': 0.0001, 'sqin': 0.00064516, 
            'sqmile': 2589988.11, 'sqmm': 0.000001,
            
            # MASS (kg 기준)
            'kg': 1.0, 'lb': 0.453592, 'gm': 0.001, 'ton': 1000.0, 'Mlb': 453592.0,
            'tonne': 1000.0, 'L-ton': 1016.05, 'MMlb': 453592000.0,
            
            # TIME (sec 기준)
            'sec': 1.0, 'hr': 3600.0, 'day': 86400.0, 'min': 60.0, 'year': 31536000.0,
            'month': 2628000.0, 'week': 604800.0, 'nsec': 1e-9, 'oper-year': 28382400.0,
            
            # TEMPERATURE (K 기준) - 특별 처리 필요
            'K': 1.0, 'C': 'C_to_K', 'F': 'F_to_K', 'R': 0.555556,
            
            # PRESSURE (N/sqm = Pa 기준)
            'N/sqm': 1.0, 'PsIa': 6894.76, 'atm': 101325.0, 'lbf/sqft': 47.8803,
            'bar': 100000.0, 'torr': 133.322, 'in-water': 249.089, 'kg/sqcm': 98066.5,
            'mmHg': 133.322, 'kPa': 1000.0, 'mm-water': 9.80665, 'mbar': 100.0,
            'psig': 'psig_to_Pa', 'atmg': 'atmg_to_Pa', 'barg': 'barg_to_Pa',
            'pa': 1.0, 'MiPa': 1000000.0, 'Pag': 'Pag_to_Pa', 'kPag': 'kPag_to_Pa',
            'MPag': 'MPag_to_Pa', 'mbarg': 'mbarg_to_Pa', 'psi': 6894.76, 'bara': 100000.0,
            
            # ENERGY (J 기준)
            'J': 1.0, 'Btu': 1055.06, 'cal': 4.184, 'kcal': 4184.0, 'kWhr': 3600000.0,
            'ft-lbf': 1.35582, 'GJ': 1000000000.0, 'kJ': 1000.0, 'N-m': 1.0, 'MJ': 1000000.0,
            'Mcal': 4184000.0, 'Gcal': 4184000000.0, 'Mbtu': 1055060000.0,
            'MMBtu': 1055060000000.0, 'hp-hr': 2684520.0, 'MMkcal': 4184000000000.0,
            'Mmkcal': 4184000000000000.0, 'Pcu': 1055.06, 'MMPcu': 1055060000000.0,
            'kW-hr': 3600000.0,
            
            # POWER (Watt 기준)
            'Watt': 1.0, 'W': 1.0, 'hp': 745.7, 'kW': 1000.0, 'Btu/hr': 0.293071, 'cal/sec': 4.184,
            'ft-lbf/sec': 1.35582, 'MIW': 1000000.0, 'GW': 1000000000.0, 'MJ/hr': 277.778,
            'kcal/hr': 1.16222, 'Gcal/hr': 1162220.0, 'MMBtu/hr': 293071.0, 'MBtu/hr': 293.071,
            'Mhp': 745700000.0,
            
            # FLOW (kg/sec 기준)
            'kg/sec': 1.0, 'lb/hr': 0.000125998, 'kg/hr': 0.000277778, 'lb/sec': 0.453592,
            'Mlb/hr': 125.998, 'tons/day': 0.0115741, 'Mcfh': 0.00786579, 'tonne/hr': 0.277778,
            'lb/day': 5.24991e-06, 'kg/day': 1.15741e-05, 'tons/hr': 0.277778, 'kg/min': 0.0166667,
            'kg/year': 3.17098e-08, 'gm/min': 1.66667e-05, 'gm/hr': 2.77778e-07, 'gm/day': 1.15741e-08,
            'Mgm/hr': 0.277778, 'Ggm/hr': 277.778, 'Mgm/day': 0.0115741, 'Ggm/day': 11.5741,
            'lb/min': 0.00755987, 'MMlb/hr': 125998.0, 'Mlb/day': 5.24991, 'MMlb/day': 5249.91,
            'lb/year': 1.43833e-08, 'Mlb/year': 1.43833e-05, 'MMIb/year': 0.0143833,
            'tons/min': 16.6667, 'Mtons/year': 31.7098, 'MMtons/year': 31709.8,
            'L-tons/min': 16.9333, 'L-tons/hr': 0.282222, 'L-tons/day': 0.0117593,
            'ML-tons/year': 32.1507, 'MML-tons/year': 32150.7, 'ktonne/year': 0.0317098,
            'kg/oper-year': 3.52775e-08, 'lb/oper-year': 1.59891e-08, 'Mlb/oper-year': 1.59891e-05,
            'MIMIb/oper-year': 0.0159891, 'Mtons/oper-year': 35.2775, 'MMtons/oper-year': 35277.5,
            'ML-tons/oper-year': 35.7230, 'MML-tons/oper-year': 35723.0, 'ktonne/oper-year': 0.0352775,
            'gm/sec': 0.001, 'tons/year': 0.0317098, 'tonne/day': 0.0115741, 'tonne/year': 0.0317098,
            'tons/oper-year': 0.0352775, 'tonne/oper-year': 0.0352775,
            
            # MOLE-FLOW (kmol/sec 기준)
            'kmol/sec': 1.0, 'lbmol/hr': 0.000125998, 'kmol/hr': 0.000277778, 'MMscfh': 0.000783986,
            'MMscmh': 0.000022414, 'mol/sec': 0.001, 'lbmol/sec': 0.453592, 'scmh': 0.000022414,
            'bmol/day': 1.15741e-05, 'kmol/day': 1.15741e-05, 'MMscfd': 0.00000907407,
            'Mlscfd': 0.00000907407, 'scfm': 0.000000471947, 'mol/min': 1.66667e-05,
            'kmol/khr': 0.000277778, 'kmol/Mhr': 0.277778, 'mol/hr': 2.77778e-07,
            'Mmol/hr': 0.277778, 'Mlbmol/hr': 0.125998, 'lbmol/Mhr': 0.125998,
            'lbmol/MMhr': 125.998, 'Mscfm': 0.000471947, 'scfh': 7.86579e-08, 'scfd': 3.27741e-09,
            'ncmh': 0.000022414, 'ncmd': 9.33917e-07, 'ACFM': 0.000000471947, 'kmol/min': 0.0166667,
            'kmol/week': 1.65344e-06, 'kmol/month': 3.80517e-07, 'kmol/year': 3.17098e-08,
            'kmol/oper-year': 3.52775e-08, 'lbmol/min': 0.00755987,
            
            # VOLUME-FLOW (cum/sec 기준)
            'cum/sec': 1.0, 'm3/s': 1.0, 'm^3/s': 1.0, 'cuft/hr': 7.86579e-06, 'l/min': 1.66667e-05, 'gal/min': 6.30902e-05,
            'gal/hr': 1.05150e-06, 'bbl/day': 1.84013e-06, 'cum/hr': 0.000277778, 'm3/h': 0.000277778, 'm^3/h': 0.000277778, 'cuft/min': 0.000471947,
            'bbl/hr': 4.41631e-05, 'cuft/sec': 0.0283168, 'cum/day': 1.15741e-05, 'cum/year': 3.17098e-08,
            'l/hr': 2.77778e-07, 'kbbl/day': 0.00184013, 'MMcuft/hr': 7.86579, 'MMcuft/day': 0.327741,
            'Mcuft/day': 0.000327741, 'l/sec': 0.001, 'l/day': 1.15741e-08, 'cum/min': 0.0166667,
            'kcum/sec': 1000.0, 'kcum/hr': 0.277778, 'kcum/day': 0.0115741, 'Mcum/sec': 1000000.0,
            'Mcum/hr': 277.778, 'Mcum/day': 11.5741, 'cuft/day': 3.27741e-07, 'Mcuft/min': 0.471947,
            'Mcuft/hr': 0.00786579, 'Mgal/min': 63.0902, 'MMgal/min': 63090.2, 'Mgal/hr': 1.05150,
            'MMgal/hr': 1051.50, 'Mbbl/hr': 44.1631, 'MMbbl/hr': 44163.1, 'Mbbl/day': 1.84013,
            'MMbbl/day': 1840.13, 'cum/oper-year': 3.52775e-08,
            
            # VOLUME (cum 기준)
            'cum': 1.0, 'cuft': 0.0283168, 'l': 0.001, 'cuin': 1.63871e-05, 'gal': 0.00378541,
            'bbl': 0.158987, 'cc': 0.000001, 'kcum': 1000.0, 'Mcum': 1000000.0, 'Mcuft': 28316.8,
            'MMcuft': 28316800.0, 'ml': 0.000001, 'kl': 1.0, 'MMl': 1000000.0, 'Mgal': 3785.41,
            'MMgal': 3785410.0, 'UKgal': 0.00454609, 'MUKgal': 4546.09, 'MMUKgal': 4546090.0,
            'Mbbl': 158987.0, 'MMbbl': 158987000.0, 'kbbl': 158.987, 'cuyd': 0.764555,
            
            # VELOCITY (m/sec 기준)
            'm/sec': 1.0, 'ft/sec': 0.3048, 'mile/hr': 0.44704, 'km/hr': 0.277778,
            'ft/min': 0.00508, 'mm/day': 1.15741e-08, 'mm/hr': 2.77778e-07, 'mm/day30': 1.15741e-08,
            'in/day': 2.93995e-07,
            
            # DENSITY (kg/cum 기준)
            'kg/cum': 1.0, 'lb/cuft': 16.0185, 'gm/cc': 1000.0, 'lb/gal': 119.826,
            'gm/cum': 0.001, 'gm/ml': 1000.0, 'lb/bbl': 2.85301, 'gm/l': 1.0,
            'mg/l': 0.001, 'mg/cc': 1.0, 'mg/cum': 0.000001,
            
            # MOLE-DENSITY (kmol/cum 기준)
            'kmol/cum': 1.0, 'lbmol/cuft': 16.0185, 'mol/cc': 1000.0, 'lbmol/gal': 119.826,
            'mol/l': 1.0, 'mmol/cc': 1.0, 'mmol/l': 0.001,
            
            # MASS-DENSITY (kg/cum 기준) - DENSITY와 동일
            # MOLE-VOLUME (cum/kmol 기준)
            'cum/kmol': 1.0, 'cuft/lbmol': 0.0624280, 'cc/mol': 0.001, 'ml/mol': 0.001,
            'bbl/mscf': 0.158987,
            
            # ELEC-POWER (Watt 기준) - POWER와 동일
            'MW': 1000000.0,
            
            # UA (J/sec-K 기준)
            'J/sec-K': 1.0, 'Btu/hr-R': 0.527527, 'cal/sec-K': 4.184, 'kJ/sec-K': 1000.0,
            'kcal/sec-K': 4184.0, 'kcal/hr-K': 1.16222, 'Btu/hr-F': 0.527527, 'kW/k': 1000.0,
            
            # WORK (J 기준) - ENERGY와 동일
            # HEAT (J 기준) - ENERGY와 동일
            
            # COMPOSITION (mol-fr 기준) - 무차원이므로 변환 불필요
            'mol-fr': 1.0, 'mass-fr': 1.0
        }
    
    def convert_to_si(self, value: float, from_unit: str, unit_type: str) -> tuple[float, str]:
        """
        임의의 단위를 SI 기준 단위로 변환
        
        Args:
            value: 변환할 값
            from_unit: 원래 단위
            unit_type: 물리량 타입 (예: 'PRESSURE', 'POWER', 'TEMPERATURE' 등)
            
        Returns:
            tuple: (변환된 값, SI 단위)
        """
        try:
            # SI 기준 단위 가져오기
            si_unit = self._si_base_units.get(unit_type)
            if not si_unit:
                raise ValueError(f"Unsupported unit type: {unit_type}")
            
            # 이미 SI 단위인 경우
            if from_unit == si_unit:
                return value, si_unit
            
            # 특별 변환이 필요한 경우들
            if unit_type == 'TEMPERATURE':
                converted_value = self._convert_temperature_to_kelvin(value, from_unit)
                return converted_value, si_unit
            
            elif unit_type == 'PRESSURE':
                # 압력의 경우 게이지 압력 처리
                if from_unit in ['psig', 'atmg', 'barg', 'Pag', 'kPag', 'MPag', 'mbarg']:
                    # 게이지 압력을 절대 압력으로 변환
                    abs_value = self._convert_pressure_gauge_to_absolute(value, from_unit)
                    # 절대 압력 단위로 변환
                    if from_unit == 'psig':
                        from_unit = 'PsIa'
                    elif from_unit == 'atmg':
                        from_unit = 'atm'
                    elif from_unit == 'barg':
                        from_unit = 'bar'
                    elif from_unit == 'Pag':
                        from_unit = 'pa'
                    elif from_unit == 'kPag':
                        from_unit = 'kPa'
                    elif from_unit == 'MPag':
                        from_unit = 'MiPa'
                    elif from_unit == 'mbarg':
                        from_unit = 'mbar'
                    value = abs_value
                
                # 환산 계수 확인
                if from_unit not in self._conversion_factors:
                    raise ValueError(f"Unsupported pressure unit: {from_unit}")
                
                factor = self._conversion_factors[from_unit]
                if isinstance(factor, str):
                    raise ValueError(f"Special conversion required for {from_unit}, but not implemented")
                
                converted_value = value * factor
                return converted_value, si_unit
            
            else:
                # 일반적인 단위 변환
                if from_unit not in self._conversion_factors:
                    raise ValueError(f"Unsupported unit: {from_unit}")
                
                factor = self._conversion_factors[from_unit]
                if isinstance(factor, str):
                    raise ValueError(f"Special conversion required for {from_unit}, but not implemented")
                
                converted_value = value * factor
                return converted_value, si_unit
                
        except Exception as e:
            raise ValueError(f"Unit conversion error for {from_unit} ({unit_type}): {str(e)}")
    
    def convert_from_si(self, value_si: float, to_unit: str, unit_type: str) -> float:
        """
        SI 단위에서 목표 단위로 변환
        
        Args:
            value_si: SI 단위의 값
            to_unit: 목표 단위
            unit_type: 물리량 타입
            
        Returns:
            변환된 값
        """
        try:
            # SI 기준 단위 가져오기
            si_unit = self._si_base_units.get(unit_type)
            if not si_unit:
                raise ValueError(f"Unsupported unit type: {unit_type}")
            
            # 이미 SI 단위인 경우
            if to_unit == si_unit:
                return value_si
            
            # 특별 변환이 필요한 경우들
            if unit_type == 'TEMPERATURE':
                return self._convert_temperature_from_kelvin(value_si, to_unit)
            
            else:
                # 일반적인 단위 변환
                if to_unit not in self._conversion_factors:
                    raise ValueError(f"Unsupported unit: {to_unit}")
                
                factor = self._conversion_factors[to_unit]
                if isinstance(factor, str):
                    raise ValueError(f"Special conversion required for {to_unit}, but not implemented")
                
                return value_si / factor
                
        except Exception as e:
            raise ValueError(f"Unit conversion error: {str(e)}")
    
    def _convert_temperature_to_kelvin(self, value: float, from_unit: str) -> float:
        """온도를 켈빈으로 변환"""
        if from_unit == 'K':
            return value
        elif from_unit == 'C':
            return value + 273.15
        elif from_unit == 'F':
            return (value - 32) * 5/9 + 273.15
        elif from_unit == 'R':
            return value * 5/9
        else:
            raise ValueError(f"Unsupported temperature unit: {from_unit}")
    
    def _convert_temperature_from_kelvin(self, value_k: float, to_unit: str) -> float:
        """켈빈에서 다른 온도 단위로 변환"""
        if to_unit == 'K':
            return value_k
        elif to_unit == 'C':
            return value_k - 273.15
        elif to_unit == 'F':
            return (value_k - 273.15) * 9/5 + 32
        elif to_unit == 'R':
            return value_k * 9/5
        else:
            raise ValueError(f"Unsupported temperature unit: {to_unit}")
    
    def _convert_pressure_gauge_to_absolute(self, value: float, from_unit: str) -> float:
        """게이지 압력을 절대 압력으로 변환"""
        if from_unit == 'psig':
            return value + 14.696
        elif from_unit == 'atmg':
            return value + 1.0
        elif from_unit == 'barg':
            return value + 1.01325
        elif from_unit == 'Pag':
            return value + 101325.0
        elif from_unit == 'kPag':
            return value + 101.325
        elif from_unit == 'MPag':
            return value + 0.101325
        elif from_unit == 'mbarg':
            return value + 1013.25
        else:
            return value  # 이미 절대 압력인 경우
    
    # 편의 메서드들
    def convert_power_to_kw(self, value: float, from_unit: Optional[str]) -> Optional[float]:
        """전력을 kW로 변환"""
        if from_unit is None:
            return None
        
        # 이미 kW 단위인 경우 변환하지 않음
        if from_unit in ['kW', 'kw']:
            return value
        
        try:
            converted_value, _ = self.convert_to_si(value, from_unit, 'POWER')
            # Watt를 kW로 변환
            return converted_value / 1000.0
        except Exception as e:
            print(f"POWER CONVERSION ERROR: {e}")
            return None
    
    def convert_pressure_to_bar(self, value: float, from_unit: Optional[str]) -> Optional[float]:
        """압력을 bar로 변환"""
        if from_unit is None:
            return None
        
        # 이미 bar 단위인 경우 변환하지 않음
        if from_unit in ['bar', 'bara']:
            return value
        
        try:
            converted_value, _ = self.convert_to_si(value, from_unit, 'PRESSURE')
            # Pa를 bar로 변환
            return converted_value / 100000.0
        except Exception as e:
            print(f"PRESSURE CONVERSION ERROR: {e}")
            return None
    
    def convert_flow_to_m3_s(self, value: float, from_unit: Optional[str]) -> Optional[float]:
        """유량을 m³/s로 변환"""
        if from_unit is None:
            return None
        
        # 이미 m³/s 단위인 경우 변환하지 않음
        if from_unit in ['m3/s', 'm^3/s', 'cum/sec']:
            return value
        
        try:
            converted_value, _ = self.convert_to_si(value, from_unit, 'VOLUME-FLOW')
            return converted_value
        except Exception as e:
            print(f"FLOW CONVERSION ERROR: {e}")
            return None


# 전역 인스턴스 생성
_unit_converter = UnitConverter()


# 편의 함수들 (기존 인터페이스 호환성 유지)
def convert_power_to_target_unit(value_kw: float, target_unit: str) -> float:
    """전력을 kW에서 목표 단위로 변환"""
    return _unit_converter.convert_from_si(value_kw * 1000.0, target_unit, 'POWER') / 1000.0


def convert_flow_to_target_unit(value_m3_s: float, target_unit: str) -> float:
    """유량을 m³/s에서 목표 단위로 변환"""
    return _unit_converter.convert_from_si(value_m3_s, target_unit, 'VOLUME-FLOW')


def convert_power_to_kw(value: float, unit: Optional[str]) -> Optional[float]:
    """전력을 kW로 변환"""
    return _unit_converter.convert_power_to_kw(value, unit)


def convert_pressure_to_bar(value: float, unit: Optional[str]) -> Optional[float]:
    """압력을 bar로 변환"""
    return _unit_converter.convert_pressure_to_bar(value, unit)


def convert_flow_to_m3_s(value: float, unit: Optional[str]) -> Optional[float]:
    """유량을 m³/s로 변환"""
    return _unit_converter.convert_flow_to_m3_s(value, unit)


def convert_to_si_units(value: float, from_unit: str, unit_type: str) -> tuple[float, str]:
    """통합 SI 변환 함수"""
    return _unit_converter.convert_to_si(value, from_unit, unit_type)


# =============================================================================
# 장비별 최소 크기 제한
# =============================================================================

MIN_SIZE_LIMITS = {
    "pump": {
        "centrifugal": 1.0,      # kW
        "reciprocating": 0.1,    # kW
    },
    "compressor": {
        "centrifugal": 450.0,    # kW
        "axial": 450.0,          # kW
        "reciprocating": 450.0,  # kW
    },
    "turbine": {
        "axial": 1.0,           # kW (100 kW에서 1 kW로 낮춤)
        "radial": 1.0,          # kW (100 kW에서 1 kW로 낮춤)
    },
    "fan": {
        "centrifugal_radial": 1.0,    # m³/s
        "centrifugal_backward": 1.0,  # m³/s
        "centrifugal_forward": 1.0,   # m³/s
        "axial": 1.0,                 # m³/s
    }
}


def check_minimum_size_limit(equipment_type: str, subtype: str, size_value: float, size_unit: str) -> tuple[bool, str]:
    """
    장치 크기가 최소 제한을 만족하는지 확인
    Returns: (is_valid, error_message)
    """
    min_limit = MIN_SIZE_LIMITS.get(equipment_type, {}).get(subtype)
    
    if min_limit is None:
        return True, ""  # 제한이 정의되지 않은 경우 통과
    
    if size_value < min_limit:
        return False, f"under limit (min: {min_limit} {size_unit})"
    
    return True, ""


# =============================================================================
# 장비별 최대 크기 제한 (분할 기준)
# =============================================================================

MAX_SIZE_LIMITS = {
    "pump": {
        "centrifugal": 1000.0,   # kW
        "reciprocating": 1000.0, # kW
    },
    "compressor": {
        "centrifugal": 10000.0,  # kW
        "axial": 10000.0,        # kW
        "reciprocating": 10000.0,# kW
    },
    "turbine": {
        "axial": 10000.0,        # kW
        "radial": 10000.0,       # kW
    },
    "fan": {
        "centrifugal_radial": 100.0,    # m³/s
        "centrifugal_backward": 100.0,  # m³/s
        "centrifugal_forward": 100.0,   # m³/s
        "axial": 100.0,                 # m³/s
    }
}


def get_max_size_limit(equipment_type: str, subtype: str) -> Optional[float]:
    """장비별 최대 크기 제한 반환"""
    return MAX_SIZE_LIMITS.get(equipment_type, {}).get(subtype)


# =============================================================================
# CEPCI 인덱스 데이터
# =============================================================================

CEPCI_BY_YEAR = {
    2017: 567.5,  # Turton 기준년도
    2018: 603.1,
    2019: 607.5,
    2020: 596.2,
    2021: 708.0,
    2022: 778.8,
    2023: 789.6,
    2024: 800.0,  # 추정값
    2025: 810.0,  # 추정값
}


def get_cepi_index(year: int) -> float:
    """연도별 CEPCI 인덱스 반환"""
    return CEPCI_BY_YEAR.get(year, 800.0)  # 기본값 2024년


# =============================================================================
# 레거시 호환성 함수들 (기존 코드와의 호환성 유지)
# =============================================================================

def is_gauge_pressure_unit(unit: Optional[str]) -> bool:
    """게이지 압력 단위인지 확인"""
    if unit is None:
        return False
    gauge_units = {'barg', 'psig', 'kpag', 'mpag', 'mbarg'}
    return unit.lower() in gauge_units


# =============================================================================
# 전역 인스턴스 접근 함수
# =============================================================================

def get_unit_converter() -> UnitConverter:
    """전역 UnitConverter 인스턴스 반환"""
    return _unit_converter
