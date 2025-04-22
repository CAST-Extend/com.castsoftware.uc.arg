from IPython.display import display
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import format_table
from cast_arg.pages.hl_report import HLPage

from pandas import DataFrame,Series,json_normalize,ExcelWriter,concat
from os.path import abspath
from sys import exc_info

class CloudContainer(HLPage):

    def report(self,app:str,app_no:int) -> bool:
        status = True
        try:

            self.get_cloud_report(app)

            self.tag_prefix = f'app{app_no}'
            self.slide = self.ppt.get_slide(self.ppt.get_shape_by_name(f'{self.tag_prefix}_CloudContainerSlide'))
            if self.slide is None: 
                self.log.warning(f'Cloud Container Slide not found for {app}')
                return False

            container_df=self.get_cloud_container(app)
            container_df['Effort'] = container_df['Effort']/60/8
            container_df['Effort'] = container_df['Effort'].round(1)

            self.replace_text('CloudIndex',round(self._get_metrics(app)['cloudReady']*100,1),shape=True,slide=self.slide)
            self.replace_text('CloudContainerBlockers',len(container_df),shape=True,slide=self.slide)
            self.replace_text('CloudContainerOccurrences',int(container_df['roadblocks'].sum()),shape=True,slide=self.slide)

            df_high = container_df[container_df['criticality']=='High']
            df_high = df_high.sort_values(['technology','roadblocks'], ascending=False)

            df_medium = container_df[container_df['criticality'] == 'Medium']
            df_medium = df_medium.sort_values(['technology','roadblocks'], ascending=False)
            
            df_low = container_df[container_df['criticality'] == 'Low']
            df_low = df_low.sort_values(['technology','roadblocks'], ascending=False)

            rslt = concat([df_high, df_medium, df_low])            

            self.ppt.update_table(f'{self.tag_prefix}_CloudContainerTable',rslt,app,include_index=False,)  

            self.ppt.save()
            pass

        except Exception as ex:
            ex_type, ex_value, ex_traceback = exc_info()
            self.log.warning(f'{ex_type.__name__}: {ex_value} while in {__class__}')
            status = False

        return status


    def get_cloud_report(self, app:str) -> DataFrame:
#         #explore/companies/lastResults/450070/alerts/export?companySwitch=31773&activePagination=&pageOffset=&maxEntryPerPage=&maxPage=
#         #export/configuration?companySwitch=31773&reportType=Application
#         ret = 0
#         app_id = self.get_app_id(app)
#         url = f'/export/configuration?companySwitch={self._instance_id}&reportType=Application'
# #            url = f'/domains/{self._instance_id}/applications/{app_id}/export/configuration?companySwitch={self._instance_id}&reportType=Application'
        
#         try:
#             #https://rpa.casthighlight.com/WS/explore/companies/lastResults/450070/alerts/export?companySwitch=31773&activePagination=&pageOffset=&maxEntryPerPage=&maxPage=
#             url = f'/explore/companies/lastResults/{app_id}/alerts/export?companySwitch={Highlight._instance_id}&activePagination=&pageOffset=&maxEntryPerPage=&maxPage='
#             status = self.get(url)

            
#             pass
#         except KeyError as ex:
#             self.log.warning(f'Cloud Report not found for {app}')
        pass

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
