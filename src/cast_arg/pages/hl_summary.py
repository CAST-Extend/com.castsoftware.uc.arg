from cast_common.highlight import Highlight
from cast_common.util import list_to_text,convert_LOC

from pandas import DataFrame, concat
from cast_common.powerpoint import PowerPoint
from pages.hl_report import HLPage
from math import ceil
from inspect import currentframe

class HighlightSummary(HLPage):

    def __init__(self,day_rate:int,output:str=None,ppt:PowerPoint=None):
        super().__init__(output=output,ppt=ppt)

        self._day_rate = day_rate
        pass        

    def report(self, app_name: str | list = None, app_no: int = 0) -> bool:
        """
        Generate a highlight summary report for the specified application(s).
        """
        if type(app_name) is list:
            self.tag_prefix = 'port'
        else:
            self.tag_prefix = f'app{app_no}'
            app_name = [app_name]

        tech_df, comp_total = self._get_technology_data(app_name)
        cloud_df = self._get_cloud_detail(app_name)
        green_df = self._get_green_detail(app_name)
        oss_data = self._get_oss_data(app_name)
        app_scores = self._calculate_scores(app_name)

        self._process_data(app_name,tech_df, comp_total, cloud_df, green_df, oss_data,app_scores)

        return True

    def _get_technology_data(self, app_name: list) -> tuple[DataFrame, int]:
        """
        Retrieve and process technology data for the specified applications.
        """
        tech_df = DataFrame()
        comp_total = 0
        for app in app_name:
            df = self.get_technology(app)
            tech_df = concat([tech_df, df])
            comp_total += self.get_component_total(app)
        return tech_df, comp_total

    def _get_cloud_detail(self, app_name: list) -> DataFrame:
        """
        Retrieve and process cloud requirement details for the specified applications.
        """
        cloud_df = DataFrame()
        for app in app_name:
            df = self.get_cloud_detail(app)
            df = df[df['cloudRequirement.criticality'].isin(['Critical', 'High'])]
            cloud_df = concat([cloud_df, df])
        return cloud_df

    def _get_green_detail(self, app_name: list) -> DataFrame:
        """
        Retrieve and process green requirement details for the specified applications.
        """
        green_df = DataFrame()
        for app in app_name:
            green_df = concat([green_df, self.get_green_detail(app)])
        return green_df

    def _get_oss_data(self, app_name: list) -> dict:
        """
        Retrieve and process open source safety data for the specified applications.
        """
        oss = {'cve': {}, 'total': {}}
        t_high_license = t_medium_license = t_low_license = t_license = comp_total = 0

        for app in app_name:
            t_high_license += len(self.get_license_high(app))
            t_medium_license += len(self.get_license_medium(app))
            t_low_license += len(self.get_license_low(app))
            t_license += t_high_license + t_medium_license + t_low_license
            comp_total += self.get_component_total(app)

            for crit in ['critical','high','medium']:
                mth = getattr(self,f'get_cve_{crit}')
                df = mth(app)
                if df is not None:
                    if crit in oss['cve'].keys():
                        oss['cve'][crit] = concat([oss['cve'][crit],df])
                    else: 
                        oss['cve'][crit] = df

        oss['license'] = {'high': t_high_license, 'medium': t_medium_license, 'low': t_low_license, 'total': t_license}   
        return oss

    def _calculate_scores(self, app_name: list) -> dict:
        """
        Calculate software quality scores for the specified applications.
        """
        scores = {}
        for app in app_name:
            scores[app] = self.calc_scores([app])
        return scores

    def _process_data(self, app_name: list, tech_df, comp_total, cloud_df, green_df, oss_data,app_scores):
        """
        Process the collected data and update the report placeholders.
        """
        low_health={}
        t_apps = len(app_name)
        port_scores = self.calc_scores(app_name)
        for key in self.grades:
            try:
                score = port_scores[key] 
                hml = self.update_grade_score(key, score)

                # calculate the "BEST" and "WORST" grades for each tile
                high=0
                low = 100
                for app in app_name:
                    g = app_scores[app][key]
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
                    for oss_key in oss_data['cve'].keys():
                        eff = 0
                        cmpnt_df=DataFrame()
                        df = oss_data['cve'][oss_key]
                        if df is not None:
                            cmpnt_df['component']=df['component']
                            if not cmpnt_df.empty:
                                total_cmpnts_df = concat([cmpnt_df,total_cmpnts_df])
                                eff = ceil(len(cmpnt_df['component'].unique())/2)
                            cnt = len(df['cve'].unique())
                            total_cves += cnt

                        self.replace_text(f'{oss_key}_oss_total',cnt)
                        self.replace_text(f'{oss_key}_oss_effort',eff)
                        pass
        
                    self.replace_text('cve_total',total_cves)
                    total_eff = 0
                    if not total_cmpnts_df.empty:
                        total_eff = ceil(len(total_cmpnts_df['component'].unique())/2)
                        self.replace_text(f'{key}_total_cve_effort',total_eff)
                    
                    for lic_key in oss_data['license'].keys():
                        self.replace_text(f'{lic_key}_license_total', oss_data['license'][lic_key])

                    pass
                    self.replace_text('oss_effort',ceil(total_cves/2))

                    health = self.get_software_health_score(app)
                    #save information to be used for ranking
                    low_health[app]=health # save the health score for ranking

                pass

            except (KeyError) as ex:
                self.log.error(f'{repr(ex)} in {currentframe().f_code.co_name}')
            except Exception as ex:
                self.log.error(f'unknown: {repr(ex)} in {currentframe().f_code.co_name}')

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
        # self.replace_text('oss_total_licenses',f'{t_license:,}')

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

        # if not green_df.empty:
        #     self.replace_text('green_hml',self.get_software_green_hml(score=t_green))

        # if self.tag_prefix == 'port_hl':
        #     (health_low_app,health_low_score,health_high_app,health_high_score) = self._get_high_low_factors(low_health)
        #     self.replace_text('softwareHealth_low_app',health_low_app)
        #     self.replace_text('softwareHealth_high_app',health_high_app)
        #     self.replace_text('softwareHealth_low_score',health_low_score)
        #     self.replace_text('softwareHealth_high_score',health_high_score)

        #     (oss_low_app,oss_low_crit_total,oss_high_app,oss_high_crit_total) = self._get_high_low_factors(oss_cve_counts)
        #     self.replace_text('oss_low_app',oss_low_app)
        #     self.replace_text('oss_high_app',oss_high_app)
        #     self.replace_text('oss_low_critical_total',oss_low_crit_total)
        #     self.replace_text('oss_high_critical_total',oss_high_crit_total)

    def update_grade_score(self, key, score):
        """
        Update the score, color, and text for a given grade.

        Args:
            key (str): The grade key (e.g., 'health_score', 'transfer_score', etc.).
            score (float): The calculated score for the grade.
        """
        threshold = self.grades[key]['threshold']
        hml = self._get_hml_category(score, threshold)
        color = self.get_hml_color(hml)
        self._update_tile_color(key, color)
        self._update_score_text(key, score)
        self._replace_from_options(key, hml)
        self.replace_text(f'{key}_score', f'{score}%', shape=True)
        return hml

    def _get_hml_category(self, score, threshold):
        """
        Determine the High/Medium/Low (HML) category based on the score and threshold.

        Args:
            score (float): The calculated score.
            threshold (list): A list containing the low and high threshold values.

        Returns:
            str: The HML category ('low', 'medium', or 'high').
        """
        if len(threshold) > 1:
            if score < threshold[0]:
                return 'low'
            elif score > threshold[1]:
                return 'high'
            else:
                return 'medium'
        else:
            raise ValueError("Invalid threshold format")

    def _update_tile_color(self, key, color):
        """
        Update the color of the tile for the given grade.

        Args:
            key (str): The grade key.
            color (str): The color to be applied to the tile.
        """
        tile_name = f'{self.tag_prefix}_{key}_tile'
        PowerPoint.ppt.fill_text_box_color(tile_name, color)

    def _update_score_text(self, key, score):
        """
        Update the text displaying the score for the given grade.

        Args:
            key (str): The grade key.
            score (float): The calculated score.
        """
        score_text = f'{key}_score'
        self.replace_text(score_text, f'{score}%', shape=True)

    def _replace_from_options(self, key, hml):
        """
        Replace the text based on the HML category for the given grade.

        Args:
            key (str): The grade key.
            hml (str): The HML category ('low', 'medium', or 'high').
        """
        self.replace_from_options(key, hml)

