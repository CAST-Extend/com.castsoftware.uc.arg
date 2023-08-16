from IPython.display import display
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import format_table
from pandas import DataFrame,Series,json_normalize,ExcelWriter
from os.path import abspath
from sys import exc_info

class GreenIt(Highlight):

    def report(self,app:str,app_no:int,prs:PowerPoint,output:str) -> bool:
        status = True
        try:
            index = self.get_green_indexes(app)
            for idx,val in index.items():
                if idx == 'greenOccurrences':
                    val = int(val)
                tag = f'{{app{app_no}_{idx}}}'
                prs.replace_text(tag,val)

            detail = self.get_green_detail(app)
            detail = detail[detail['Occurrences'] != 0]
            detail=detail.drop(columns=['Contribution'])

            agr = detail[['Technology','Occurrences']].groupby('Technology').aggregate('sum').reindex()
            prs.update_chart(f'app{app_no}_GreenTechPieChart',agr)
            agr.sort_values('Occurrences',ascending=False,inplace=True)
            prs.replace_text(f'{{app{app_no}_green_top_lang}}',agr.index[0])
            prs.replace_text(f'{{app{app_no}_green_top_lang_count}}',int(agr.iloc[0,0]))

            agr = detail[['Name','Occurrences']].groupby('Name').aggregate('sum').reindex()
            agr.sort_values('Occurrences',ascending=False,inplace=True)
            prs.replace_text(f'{{app{app_no}_green_top_rule1}}',agr.index[0])
            prs.replace_text(f'{{app{app_no}_green_top_rule2}}',agr.index[1])

            detail.sort_values(by=['Occurrences'],ascending=False,inplace=True)
            #detail['Contribution'] = detail['Contribution'].apply(lambda x: '{0:.2f}%'.format(x).rjust(10))
            detail['Occurrences'] = detail['Occurrences'].apply(lambda x: '{0:,.0f}'.format(x).rjust(10))
            self.create_excel(app,detail,output)

            #detail = detail.astype({'Contribution':'string'})
            detail['Technology'] = detail['Technology'].str.ljust(20)
            prs.update_table(f'app{app_no}_GreenDetailTable',detail,include_index=False,max_rows=8)


            pass      
        except Exception as ex:
            ex_type, ex_value, ex_traceback = exc_info()
            self.log.error(f'{ex_type.__name__}: {ex_value}')
            status = False

        return status

    def get_green_indexes(self,app_name:str) -> Series:
        self.log.info(f'Retrieving green it index data for: {app_name}')
        try:
            df = json_normalize(self._get_metrics(app_name)['greenDetail'],['greenIndexDetails'],meta=['technology','greenIndexScan']).dropna(axis='columns')
            df = df[['greenIndexScan','greenOccurrences','greenEffort']]
            df = df.aggregate(['sum','average'])

            if self.log.is_debug:
                self.log.debug('aggragation')
                display(df)

            rslt = Series()
            rslt.loc['greenEffort']=round(df.loc['sum','greenEffort']/60/8,1)
            rslt.loc['greenOccurrences']=round(df.loc['sum','greenOccurrences'],0)
            rslt.loc['greenIndexScan']=round(df.loc['average','greenIndexScan']*100,1)

            if self.log.is_debug:
                self.log.debug('final')
                display(rslt)

            return rslt
        except KeyError as ke:
            self.warning(f'{app_name} has no Green IT Data')
            return None

    def get_green_detail(self,app_name:str)->DataFrame:
        """Highlight green it data

        Args:
            app_name (str): name of the application

        Returns:
            DataFrame: flattened version of the Highlight green it data
        """
        self.log.info(f'Retrieving green it detail data for: {app_name}')
        try:
            df = json_normalize(self._get_metrics(app_name)['greenDetail'],['greenIndexDetails'],meta=['technology','greenIndexScan']).dropna(axis='columns')
            df.drop(columns=['greenRequirement.id','greenRequirement.hrefDoc','triggered','greenRequirement.ruleType','greenIndexScan'],inplace=True)
            df.rename(columns={'contributionScore':'Contribution',
                               'greenOccurrences':'Occurrences',
                               'greenRequirement.display':'Name',
                               'greenEffort':'Effort',
                               'technology':'Technology'},
                               inplace=True)
            df['Effort']=df['Effort'].div(60).div(8).round(2)
            df['Contribution']=df['Contribution'].round(2)
            return df[['Name','Technology','Contribution','Occurrences','Effort']]
        except KeyError as ke:
            self.warning(f'{app_name} has no Green IT Data')
            return None

    def create_excel(self,app_name:str,data:DataFrame,output:str):
        file_name = abspath(f'{output}/greenIt-Reporting-{app_name}.xlsx')
        writer = ExcelWriter(file_name, engine='xlsxwriter')
        format_table(writer,data,'Detail',width=[75,25,15,15,15],total_line=True)
        writer.close()




