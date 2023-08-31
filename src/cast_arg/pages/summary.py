
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import list_to_text,convert_LOC

from pandas import json_normalize

class HighlightSummary(Highlight):

    def __init__(self,day_rate:int):
        super().__init__()
        self._day_rate = day_rate

    def report(self,app_name:str,app_no:int,prs:PowerPoint,output:str) -> bool:
        app_tag = f'app{app_no}'

        #create list of technolgies sorted by LOC in decending order
        tech_list = list_to_text(self.get_technology(app_name)['technology'].to_list())
        self.replace_text(prs,app_tag,'tech',tech_list)

        loc=self.get_total_lines_of_code(app_name)
        (total_loc,unit) = convert_LOC(int(loc))
        self.replace_text(prs,app_tag,'total_loc',f'{total_loc} {unit}')
        
        comp_total = self.get_component_total(app_name)
        self.replace_text(prs,app_tag,'component_total',f'{comp_total:,}')

        critical = len(self.get_cve_critical(app_name)['cve'].unique())
        self.replace_text(prs,app_tag,'cve_critical_total',f'{critical:,}')

        high = len(self.get_license_high(app_name)['component'].unique())
        self.replace_text(prs,app_tag,'license_high_total',f'{high:,}')

        cost = ((critical * self._day_rate)/2)/1000
        self.replace_text(prs,app_tag,'oss_cost',f'{cost:,.1f}')

        self.replace_text(prs,app_tag,'software_health_hml',self.get_software_health_hml(app_name).capitalize())
        self.replace_text(prs,app_tag,'elegance_score',f'{self.get_software_elegance_score(app_name):,.1f}')
        self.replace_text(prs,app_tag,'agility_score',f'{self.get_software_agility_score(app_name):,.1f}')
        self.replace_text(prs,app_tag,'resiliency_score',f'{self.get_software_resiliency_score(app_name):,.1f}')

        cloud_ready = self.get_cloud_detail(app_name)
        blockers = len(cloud_ready[cloud_ready['cloudRequirement.ruleType']=='BLOCKER'])
        self.replace_text(prs,app_tag,'cloud_ready_blocker_total',f"{blockers:,f}")

        pass

    def replace_text(self,prs, prefix, item, data):
        tag = f'{{{prefix}_{item}}}'
        self.log.debug(f'{tag}: {data}')
        prs.replace_text(tag,data)

# from os.path import abspath
# from cast_common.util import format_table
# from pandas import ExcelWriter

# ppt = PowerPoint(r'E:\work\Decks\highlight-test.pptx',r'E:\work\Decks\test\highlight.pptx')

# app = 'CollabServer'
                            
# hl = HighlightSummary('n.kaplan+insightsoftwareMinerva@castsoftware.com','vadKpBFAZ8KIKb2f2y',hl_instance=383,hl_base_url='https://app.casthighlight.com',log_level=DEBUG)
# hl.report(app,1,ppt,r'E:\work\Decks\test')
# ppt.save()

