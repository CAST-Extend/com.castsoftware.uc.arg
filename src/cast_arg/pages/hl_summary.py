from cast_common.highlight import Highlight
from cast_arg.pages.hl_report import HLPage
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint
from cast_common.util import list_to_text,convert_LOC
from pandas import concat,DataFrame
from math import ceil

from pandas import json_normalize
from pptx.enum.shapes import MSO_SHAPE_TYPE

class HighlightSummary(HLPage):

    def __init__(self,day_rate:int,output:str=None,ppt:PowerPoint=None):
        super().__init__(output=output,ppt=ppt)
        self._day_rate = day_rate
        pass        

    def report(self,app_name:str|list=None,app_no:int=0) -> bool:
        if type(app_name) is list:
            self.tag_prefix = 'port'
        else:
            self.tag_prefix = f'app{app_no}'
            app_name = [app_name]

        #create list of technolgies sorted by LOC in decending order
        tech_df = DataFrame()
        comp_total = 0
        # oss_cve_df = DataFrame()
        cloud_df = DataFrame()
        green_df = DataFrame()

        t_health=t_cloud=t_oss=t_green=0
        t_high_license = t_medium_license = t_low_license = t_license = 0
        low_health={}
        oss_cve_counts={}
        t_scores = self.calc_scores(app_name)

        scores = {}

        oss={}
        oss['cve']={}
        oss['license']={}
        for app in app_name:
            df = self.get_technology(app)
            tech_df = concat([tech_df,df])
            #

            scores[app] = self.calc_scores([app])

            t_high_license += len(self.get_license_high(app))
            t_medium_license += len(self.get_license_medium(app))
            t_low_license += len(self.get_license_low(app))
            t_license += t_high_license + t_medium_license + t_low_license

            health = self.get_software_health_score(app)
            #save information to be used for ranking
            low_health[app]=health # save the health score for ranking
            
            comp_total += self.get_component_total(app)
#            oss_score += self.get_software_oss_safty_score(app) 

            """
                get the Common Vulnerabilty and Exposue information by priority
                
                The highlight class has three methods get_cve_critical, get_cve_high
                and get_cve_medum. (we don't care about the low priority items). Use 
                them to retrieve the information and store it in a dictionary for later 
                use.
            """
            for crit in ['critical','high','medium']:
                mth = getattr(self,f'get_cve_{crit}')
                df = mth(app)
                if df is not None:
                    if crit in oss['cve'].keys():
                        oss['cve'][crit] = concat([oss['cve'][crit],df])
                    else: 
                        oss['cve'][crit] = df
            pass

            df = self.get_cloud_detail(app)
            df = df[df['cloudRequirement.criticality'].isin(['Critical','High'])] 
            cloud_df = concat([cloud_df,df])

            green_df = concat([green_df,self.get_green_detail(app)])
            pass

        text = {
            'quality':{'high':'high','medium':'moderate','low':'low-level'},
            'improvement':{'high':'no immediate action required','medium':'room for improvement','low':'ample opportunity for improvement'},
            'maintain':{'high':'highly maintainable','medium':'maintainable but needs improvement','low':'is not maintainable'},

            'quality_alt_1':{'high':'well','medium':'fair','low':'bad'},
            'quality_alt_2':{'high':'impressive','medium':'fair','low':'poor'},
            'quality_alt_3':{'high':'stands out','medium':'average','low':'in need of improvement'},
            'maturity':{'high':'high','medium':'medium','low':'low'},
            'effort':{'high':'minimal','medium':'medium','low':'considerable'},
            'risk':{'high':'low amount of','medium':'average','low':'very high'}
        }

        t_apps = len(app_name)
        for key in self.grades:
            score = t_scores[key] 
            self.replace_text(f'{key}_score',score,shape=True)

            # calculate the "BEST" and "WORST" grades for each tile
            high=0
            low = 100
            for app in app_name:
                a = scores[app]
                g = a[key]
                if g < low: low = g
                if g > high: high = g

            self.replace_text(f'bmw_{key}_score',low,shape=True)
            self.replace_text(f'bmb_{key}_score',high,shape=True)
            self.replace_text(f'bmi_{key}_score',round(self._benchmark.loc[key]['avg']*100,2),shape=True)

            # is this grade score above, below or equal to the industry average?
            bm = round(Highlight._benchmark.loc[key]['avg']*100,2)
            score_bm_hml = 'equal'
            if score > bm:
                score_bm_hml = 'high'
            elif score < bm: 
                score_bm_hml = 'low'
            self.replace_text(f'{key}_bm_hle',score_bm_hml)

            if key == 'openSourceSafety':
                hml_risk = self.get_get_software_oss_risk(score=score)
                self.replace_text(f'{key}_hml_risk',hml)
                self.replace_text(f'{key}_risk',hml)

                hml_score={'high':'low','medium':'medium','low':'high'}
                self.replace_text(f'{key}_hml_score',hml_score[hml_risk])

                total_cves = 0
                total_cmpnts_df = DataFrame(columns=['component'])
                for key in oss['cve'].keys():
                    eff = 0
                    cmpnt_df=DataFrame()
                    df = oss['cve'][key]
                    if df is not None:
                        cmpnt_df['component']=df['component']
                        if not cmpnt_df.empty:
                            total_cmpnts_df = concat([cmpnt_df,total_cmpnts_df])
                            eff = ceil(len(cmpnt_df['component'].unique())/2)
                        cnt = len(df['cve'].unique())
                        total_cves += cnt

                    self.replace_text(f'{key}_cve_total',cnt)
                    self.replace_text(f'{key}_cve_effort',eff)
                    pass
    
                self.replace_text('cve_total',total_cves)
                total_eff = 0
                if not total_cmpnts_df.empty:
                    total_eff = ceil(len(total_cmpnts_df['component'].unique())/2)
                    self.replace_text(f'{key}_total_cve_effort',total_eff)


                
                self.replace_text('high_license_total',f'{t_high_license:,}')
                self.replace_text('oss_effort',ceil(total_cves/2))
                


                # oss_cve_counts[app]=len(cve_df['cve'].unique()) 
                # cve_df=cve_df['cve']
                # oss_cve_df = concat([oss_cve_df,cve_df])
            else:
                threshold = self.grades[key]['threshold']
                if len(threshold)>1:
                    if score < threshold[0]:
                        hml = 'low'
                    elif score > threshold[1]:
                        hml = 'high'
                    else:
                        hml = 'medium'
                    color = self.get_hml_color(hml)
                    PowerPoint.ppt.fill_text_box_color(f'{self.tag_prefix}_{key}_tile',color)

                    for t in text:
                        self.replace_text(f'{key}_{t}',text[t][hml])

        self.replace_text('app_count',t_apps)

        tech_df = tech_df.groupby(['technology']).sum().reset_index()   \
            [['technology', 'totalLinesOfCode', 'totalFiles']]          \
            .sort_values(by=['totalLinesOfCode'],ascending=False)

        tech_list = list_to_text(tech_df['technology'].to_list())
        self.replace_text('technology',tech_list)
        self.replace_text('technology_count',len(tech_list))

        (total_loc,unit) = convert_LOC(int(tech_df['totalLinesOfCode'].sum()))
        self.replace_text('total_loc',f'{total_loc} {unit}')

        total_files = int(tech_df['totalFiles'].sum())
        self.replace_text('total_files',f'{total_files:,}')
        self.replace_text('oss_total_components',f'{comp_total:,}')
        self.replace_text('oss_total_licenses',f'{t_license:,}')

        # if oss_cve_df.empty:
        #     oss_crit_vio_total = 0
        # else:
        #     try:
        #         oss_crit_vio_total = len(oss_cve_df[0].unique())
        #     except KeyError:
        #         oss_crit_vio_total = 0
        
        boosters = len(cloud_df[cloud_df['cloudRequirement.ruleType']=='BOOSTER'])
        blockers = len(cloud_df[cloud_df['cloudRequirement.ruleType']=='BLOCKER'])
        self.replace_text('cloud_booster_total',boosters)
        self.replace_text('cloud_blocker_total',blockers)

        if green_df.empty:
            boosters = 0
            blockers = 0
        else:
            boosters = len(green_df[green_df['greenRequirement.ruleType']=='BOOSTER'])
            blockers = len(green_df[green_df['greenRequirement.ruleType']=='BLOCKER'])
        self.replace_text('green_booster_total',boosters)
        self.replace_text('green_blocker_total',blockers)

        if not green_df.empty:
            self.replace_text('green_hml',self.get_software_green_hml(score=t_green))

        if self.tag_prefix == 'port_hl':
            (health_low_app,health_low_score,health_high_app,health_high_score) = self._get_high_low_factors(low_health)
            self.replace_text('softwareHealth_low_app',health_low_app)
            self.replace_text('softwareHealth_high_app',health_high_app)
            self.replace_text('softwareHealth_low_score',health_low_score)
            self.replace_text('softwareHealth_high_score',health_high_score)

            (oss_low_app,oss_low_crit_total,oss_high_app,oss_high_crit_total) = self._get_high_low_factors(oss_cve_counts)
            self.replace_text('oss_low_app',oss_low_app)
            self.replace_text('oss_high_app',oss_high_app)
            self.replace_text('oss_low_critical_total',oss_low_crit_total)
            self.replace_text('oss_high_critical_total',oss_high_crit_total)

    def _get_high_low_factors(self,factor:list):
        if len(factor)==0:
            return (0,0,0,0)
        low_app = min(factor, key=factor.get)
        low_score = round(factor[low_app],1)
        high_app = max(factor, key=factor.get)
        high_score = round(factor[high_app],1)
        return (low_app,low_score,high_app,high_score)


# from os.path import abspath
# from cast_common.util import format_table
# from pandas import ExcelWriter

# ppt = PowerPoint(r'E:\work\Decks\highlight-test.pptx',r'E:\work\Decks\test\highlight.pptx')

# app = 'CollabServer'
                            
# hl = HighlightSummary('n.kaplan+insightsoftwareMinerva@castsoftware.com','vadKpBFAZ8KIKb2f2y',hl_instance=383,hl_base_url='https://app.casthighlight.com',log_level=DEBUG)
# hl.report(app,1,ppt,r'E:\work\Decks\test')
# ppt.save()

