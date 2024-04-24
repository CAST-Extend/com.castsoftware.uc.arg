from IPython.display import display
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import format_table
from cast_arg.pages.hl_report import HLPage

from pandas import DataFrame,Series,json_normalize,ExcelWriter
from os.path import abspath
from sys import exc_info

class CloudMaturity(HLPage):

    def report(self,app:str,app_no:int) -> bool:
        status = True
        try:
            self.prefix = f'app{app_no}'
            self.slide = self.ppt.get_slide(self.ppt.get_shape_by_name(f'{self.prefix}_CloudTechPieChart'))
            metrics = self._get_metrics(app)

            self.ppt.replace_text(f'{{{self.prefix}_hl_CloudIndex}}',round(self._get_metrics(app)['cloudReady']*100,2),slide=self.slide)

            # get the cloud data from the Highlight REST API
            data = self.get_data(app)
            #export it to excel
            self.create_excel(app,data,self._output)

            self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_total_roadblocks}}',int(metrics['roadblocks']),slide=self.slide)
            self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_total_blockers}}',round(metrics['blockers']*100,1),slide=self.slide)
            self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_total_boosters}}',round(metrics['boosters']*100,1),slide=self.slide)
            self.get_cloud_totals(data)
                
            # update the techology chart
            agr = data[['Technology','NB Roadblocks']].groupby('Technology').aggregate('sum').reindex()
            self.ppt.update_chart(f'{self.prefix}_CloudTechPieChart',agr)

            data = data.drop(columns=['Files'])
            self.ppt.update_table(f'{self.prefix}_CloudDetailTable',data,app,include_index=False)  
            pass      
        except Exception as ex:
            ex_type, ex_value, ex_traceback = exc_info()
            self.log.error(f'{ex_type.__name__}: {ex_value} while in {__class__}')
            status = False

        return status

    def get_data(self,app:str):
        df = self.get_cloud_detail(app)
        df = df.rename(columns={'technology':'Technology',
                                    'cloudRequirement.display':'Requirements',
                                    'cloudRequirement.ruleType':'Rule Types',
                                    'roadblocks':'NB Roadblocks',
                                    'cloudEffort':'Effort',
                                    'cloudRequirement.criticality':'Criticality',
                                    'files':'Files'
                                    })

        df=df[(df['NB Roadblocks']>0) | (df['Rule Types'] == 'BOOSTER')]
        df = df.fillna(0)
        df['Effort'] = df['Effort']/60/8
        df['Effort'] = df['Effort'].round(1)

        df = df[['Requirements','Technology','NB Roadblocks','Effort','Criticality','Rule Types','Files']]

        return df

    def get_cloud_totals(self,data:DataFrame) -> None:
        total = data.sum(numeric_only=True)

        total_effort=round(total['Effort'],1)
        total_blockers=round(total['NB Roadblocks'],0)

        self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_total_effort}}',total_effort,slide=self.slide)

        #get the top tech and total occurences for it
        agr = data[['Technology','NB Roadblocks']]. \
            groupby('Technology').aggregate('sum'). \
            sort_values('NB Roadblocks',ascending=False)

        agr.iloc[:1] # get the first row (top language)
        df = agr.iloc[:1]

        self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_top_tech}}',agr.index[0],slide=self.slide)
        self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_top_tech_total}}',int(agr['NB Roadblocks'].iloc[0]),slide=self.slide)

        #len(data[data['']])
        self.ppt.replace_text(f'{{{self.prefix}_hl_cloud_top_tech_total}}',int(agr['NB Roadblocks'].iloc[0]),slide=self.slide)


    def create_excel(self,app_name:str,data:DataFrame,output:str):
        file_name = abspath(f'{output}/Cloud-Maturity-{app_name}.xlsx')
        writer = ExcelWriter(file_name, engine='xlsxwriter')
        col_widths=[50,10,10,10,10,10,10]
        cloud_tab = format_table(writer,data,'Cloud Data',col_widths)
        writer.close()



        # 
        # writer = ExcelWriter(file_name, engine='xlsxwriter')
        # format_table(writer,data,'Detail',width=[75,25,15,15,15],total_line=True)
        # writer.close()
