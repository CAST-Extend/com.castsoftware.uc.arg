
"""
    Open Source Safty Statistical Information
"""

from distutils.log import INFO
from logging import DEBUG, info, warn, error

from pandas import DataFrame, concat
from logger import Logger

import math
import util

__author__ = "Nevin Kaplan"
__copyright__ = "Copyright 2022, CAST Software"
__email__ = "n.kaplan@castsoftware.com"



class AIPStats():
    def __init__(self,day_rate,logger_level=INFO):
        self._log = Logger('AIPStats',logger_level)

        self._day_rate=day_rate
        self._effort=0
        self._violations=0
        self._data = DataFrame()

    @property
    def effort(self):
        return self._effort
    @effort.setter
    def effort(self,value):
        self._effort = int(value)
    def add_effort(self,value):
        self._effort = self._effort + int(value)
    def sub_effort(self,value):
        self._effort = self._effort - int(value)


    @property
    def cost(self):
        return self.effort * self._day_rate /1000

    @property
    def violations(self):
        return self._violations
    @violations.setter
    def violations(self,value):
        self._violations = int(value)
    def add_violations(self,value):
        self._violations = self._violations + int(value)
    def sub_violations(self,value):
        self._violations = self._violations - int(value)

    @property
    def data(self) -> DataFrame:
        return self._data
    @data.setter
    def data(self,value):
        self._data = value
    def add_data(self,data):
        self.data = concat([self.data,data],ignore_index=True)

    def business_criteria(self,filtered):
        list = []
        if not filtered.empty:
            for business in filtered['Business Criteria']:
                items = business.split(',')
                for t in items:
                    if t.strip() not in list:
                        list.append(t.strip())
        return list

    def list_violations(self,filtered):
        first = True
        text = ""
        try:
            for criteria in filtered['Technical Criteria'].unique():
                df = filtered[filtered['Technical Criteria']==criteria]
                total = df['No. of Actions'].sum()
                
                cases = 'for'
                if first:
                    cases = 'cases of'
                    first = False
                
                rule = criteria[criteria.find('-')+1:].strip().lower()
                if len(rule) == 0:
                    rule = criteria
                text = f'{text}{total} {cases} {rule}, '
            return util.rreplace(text[:-2],', ',' and ')
        except (KeyError):
            return ""

    def replace_text(self,ppt,app_no,priority):
        bus_txt = util.list_to_text(self.business_criteria(self.data)) + ' '
        vio_txt = self.list_violations(self.data)
        base_txt = f'app{app_no}_aip_{priority}'
        msg = f'Replacing {base_txt} (_eff, _cost, _vio_cnt'

        ppt.replace_text(f'{{{base_txt}_eff}}',self.effort)
        ppt.replace_text(f'{{{base_txt}_cost}}',self.cost)
        ppt.replace_text(f'{{{base_txt}_vio_cnt}}',self.violations)
        if len(bus_txt) > 0:
            msg = f'{msg}, bus_txt'
            ppt.replace_text(f'{{{base_txt}_bus_txt}}',bus_txt)
        if len(vio_txt) > 0:
            msg = f'{msg}, vio_txt'
            ppt.replace_text(f'{{{base_txt}_vio_txt}}',vio_txt)
        msg = f'{msg})'
        self._log.info(msg)
        

class LicenseStats():
    def __init__(self,logger_level=INFO):
        self._log = Logger('LicenseStats',logger_level)

        self._high=0
        self._medium=0
        self._low=0

    @property
    def high(self):
        return self._high
    @high.setter
    def high(self,value):
        self._high = int(value)
    def add_high(self,value):
        self._high = self._high + int(value)

    @property
    def medium(self):
        return self._medium
    @medium.setter
    def medium(self,value):
        self._medium = int(value)
    def add_medium(self,value):
        self._medium = self._medium + int(value)

    @property
    def low(self):
        return self._low
    @low.setter
    def low(self,value):
        self._low = int(value)
    def add_low(self,value):
        self._low = self._low + int(value)

    def replace_text(self,ppt,app_no):
        base_txt = f'app{app_no}_hl_lic'
        self._log.info(f'Replacing {base_txt} tags (_high, _med, _low)')
        ppt.replace_text(f'{{{base_txt}_high}}',self.high)
        ppt.replace_text(f'{{{base_txt}_med}}',self.medium)
        ppt.replace_text(f'{{{base_txt}_low}}',self.low)



class OssStats():
    __priorities = ['crit','high','med']

    def __init__(self, app_id=None, day_rate=-1, data=None, priority=None,logger_level=INFO):
        # super().__init__("CVE")
        self._log = Logger('OssStats',logger_level)

        if app_id is None:
            raise KeyError('invalid parameters: Highlight Application Id is required')
        if day_rate <= 0:
            raise KeyError('invalid parameters: day rate must be greater than zero')
        
        if priority is None:
            priority = ''
        else:
            if priority not in self.__priorities:
                txt = ', '.join(self.__priorities)
                raise KeyError(f'invalid parameters: Priority must be {txt}')

        self._priority = priority
        self._day_rate = day_rate

        if data is None:
            self._violations = 0
            self._components = 0
            self._effort = 0
        else:
            self._violations = data.get_data(app_id)[f'cve_{self._priority}_tot']
            self._components = data.get_data(app_id)[f'cve_{self._priority}_comp_tot']

    @property
    def violations(self):
        return int(self._violations)
    @violations.setter
    def violations(self,value):
        self._violations = int(value)
    def add_violations(self,value):
        self._violations = self._violations + int(value)

    @property
    def components(self):
        return self._components
    @components.setter
    def components(self,value):
        self._components = int(value)
    def add_components(self,value):
        self._components = self._components + int(value)

    @property
    def effort(self):
        return math.ceil(self._components/2)

    @property
    def cost(self):
        return self.effort * self._day_rate /1000

    def replace_text(self,ppt,app_no):
        base_txt = f'app{app_no}_hl_{self._priority}'
        self._log.info(f'Replacing {base_txt} tags (_eff, _comp_tot, _cost, _vio_cnt)')
        ppt.replace_text(f'{{{base_txt}_eff}}',self.effort)
        ppt.replace_text(f'{{{base_txt}_comp_tot}}',self.components)
        ppt.replace_text(f'{{{base_txt}_cost}}',self.cost)
        ppt.replace_text(f'{{{base_txt}_vio_cnt}}',self.violations)


