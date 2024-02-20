from cast_arg.pages.mri_report import MRIPage
from cast_arg.powerpoint import PowerPoint
from cast_common.util import format_table

from pandas import ExcelWriter
from os.path import abspath
from site import getsitepackages
from json import load

import numpy as np 

class StrengthImprovment(MRIPage):
    description = 'Strength and Improvment Table'

    def report(self,app_name:str,app_no:int) -> bool:
        """
            Populate the strengths and improvement page
            The necessary data is found in the loc_tbl
        """
        imp_df = self.tqi_compliance(app_name)
        imp_df.drop(columns=['Weight','Total','Succeeded','Compliance'],inplace=True)
        imp_df.sort_values(by=['Score','Rule'], inplace=True, ascending = False)

        file_name = f'{self._config.output}/health-{self._config.title_list[app_no-1]}.xlsx'
        writer = ExcelWriter(file_name, engine='xlsxwriter')
        format_table(writer,imp_df,'Health Data')
        writer.close()

        imp_df.drop(columns=['Detail'],inplace=True)
        imp_df['RGB'] = np.where(imp_df.Score >= 3,'194,236,213',\
            np.where(imp_df.Score < 2,'255,210,210','255,240,194'))
        imp_df.Score = imp_df.Score.map('{:.2f}'.format)

        #cause_name = abspath(f'{dirname(__file__)}/cause.json')

        cause_name = abspath(f'{getsitepackages()[-1]}/cast_arg/cause.json')
        imp_df['Cause']=''
        with open(cause_name) as json_file:
            tech_data = load(json_file)
        imp_df['Cause']=imp_df['Key'].map(tech_data)

        imp_df = imp_df[['Rule','Score','Cause','Failed','RGB']]
        PowerPoint.ppt.update_table(f'app{app_no}_imp_table',imp_df,app_name,include_index=False,background='RGB')


    pass