from IPython.display import display
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import format_table
from cast_arg.pages.hl_report import HLPage

from pandas import DataFrame,Series,json_normalize,ExcelWriter
from os.path import abspath
from sys import exc_info

class GreenIt(HLPage):


    # def __init__(self,output:str,  
    #              hl_user:str=None, hl_pswd:str=None,hl_basic_auth=None, hl_instance:int=0,
    #              hl_apps:str=[],hl_tags:str=[], 
    #              hl_base_url:str=None, 
    #              log_level=INFO, timer_on=False):
    #     super().__init__(hl_user, hl_pswd,hl_basic_auth, hl_instance,hl_apps,hl_tags, hl_base_url, log_level, timer_on)
    #     self._output = output

    _prefix = None
    @property
    def prefix(self) -> str:
        return self._prefix

    def report(self,app:str,app_no:int) -> bool:
        status = True
        try:
            self._prefix = f'app{app_no}'
            self.slide = self.ppt.get_slide(self.ppt.get_shape_by_name(f'{self.prefix}_GreenTechPieChart'))

            # add the green impact score
            self.ppt.replace_text(f'{{{self.prefix}_hl_greenIndex}}',round(self._get_metrics(app)['greenIndex']*100,2),slide=self.slide)

            data = self.get_data(app)   # retrieve and format the data from the Highlight REST api
            self.get_green_totals(data)

            #sort the data and export it to excel 
            data.sort_values(by=['NB Occurrences'],ascending=False,inplace=True)
            self.create_excel(app,data,self.output) 
# 

            # update the techology chart
            agr = data[['Technology','NB Occurrences']].groupby('Technology').aggregate('sum').reindex()
            self.ppt.update_chart(f'{self.prefix}_GreenTechPieChart',agr)
            
            #format the data for the deck
            data['NB Occurrences'] = data['NB Occurrences'].apply(lambda x: '{0:,.0f}'.format(x).center(10))
            data['Technology'] = data['Technology'].str.ljust(20)
            self.ppt.update_table(f'{self.prefix}_GreenDetailTable',data,app,include_index=False)            

            pass      
        except Exception as ex:
            ex_type, ex_value, ex_traceback = exc_info()
            self.log.error(f'{ex_type.__name__}: {ex_value} while in {__class__}')
            status = False

        return status

    def get_data(self,app:str):
        df = json_normalize(self._get_metrics(app)['greenDetail'],['greenIndexDetails'],meta=['technology','greenIndexScan']).dropna(axis='columns')
        df = df.rename(columns={'greenOccurrences':'NB Occurrences',
                                'greenEffort':'Effort',
                                'greenRequirement.display':'Deficiencies',
                                'technology':'Technology'
                            })
        df = df[['Deficiencies','Technology','NB Occurrences','Effort']]
        df=df[df['NB Occurrences']>0]
        df['Effort'] = df['Effort']/60/8
        df['Effort'] = df['Effort'].round(1)

        return df

    def get_green_totals(self,data:DataFrame) -> None:
        total = data.sum(numeric_only=True)

        total_effort=round(total['Effort'],1)
        total_occurences=round(total['NB Occurrences'],0)

        self.ppt.replace_text(f'{{{self.prefix}_hl_green_total_effort}}',total_effort,slide=self.slide)
        self.ppt.replace_text(f'{{{self.prefix}_hl_green_total_occurences}}',int(total_occurences),slide=self.slide)


        #get the top tech and total occurences for it
        agr = data[['Technology','NB Occurrences']]. \
            groupby('Technology').aggregate('sum'). \
            sort_values('NB Occurrences',ascending=False)

        agr.iloc[:1] # get the first row (top language)
        df = agr.iloc[:1]

        self.ppt.replace_text(f'{{{self.prefix}_hl_green_top_tech}}',agr.index[0],slide=self.slide)
        self.ppt.replace_text(f'{{{self.prefix}_hl_green_top_tech_total}}',int(agr['NB Occurrences'].iloc[0]),slide=self.slide)


        pass




    def create_excel(self,app_name:str,data:DataFrame,output:str):
        file_name = abspath(f'{output}/greenIt-Reporting-{app_name}.xlsx')
        writer = ExcelWriter(file_name, engine='xlsxwriter')
        format_table(writer,data,'Detail',width=[75,25,15,15,15],total_line=True)
        writer.close()




