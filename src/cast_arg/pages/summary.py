
from cast_common.highlight import Highlight
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint

from pandas import json_normalize

class HighlightSummary(Highlight):

    def report(self,app_name:str,app_no:int,prs:PowerPoint,output:str) -> bool:
        df = json_normalize(self._get_metrics(app_name))

        for item in ['softwareResiliency','softwareAgility','softwareElegance','cloudReady','openSourceSafety']:
            data = float(round(df[item].iloc[0] * 100,1))
            prs.replace_text(f'{{app{app_no}_{item}}}',data)

        for item in ['totalFiles']:
            data = round(df[item].iloc[0],0)
            prs.replace_text(f'{{app{app_no}_{item}}}',data)

        tech = json_normalize(df['technologies'])
        prs.replace_text(f'{{app{app_no}_technologies}}',len(tech.columns))

        for item in ['totalFiles','totalLinesOfCode','backFiredFP']:
            data = int(df[item].iloc[0])
            prs.replace_text(f'{{app{app_no}_{item}}}',data)

        prs.replace_text(f'{{app{app_no}_components}}',len(json_normalize(df['components'][0])))


        print (df.columns)


        pass

from os.path import abspath
from cast_common.util import format_table
from pandas import ExcelWriter

ppt = PowerPoint(r'E:\work\Decks\highlight-test.pptx',r'E:\work\Decks\test\highlight.pptx')

app = 'CollabServer'
                            
hl = HighlightSummary('n.kaplan+insightsoftwareMinerva@castsoftware.com','vadKpBFAZ8KIKb2f2y',hl_instance=383,hl_base_url='https://app.casthighlight.com',log_level=DEBUG)
hl.report(app,1,ppt,r'E:\work\Decks\test')
ppt.save()

