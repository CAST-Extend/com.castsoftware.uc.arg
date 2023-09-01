from cast_arg.restCall import AipData,HLData
#from cast_arg.powerpoint import PowerPoint
from cast_arg.actionPlan import ActionPlan
from cast_arg.config import Config
from cast_arg.pages.greenIt import GreenIt
from cast_arg.pages.summary import HighlightSummary

from cast_arg.stats import OssStats,AIPStats,LicenseStats
from cast_common.logger import Logger,DEBUG, INFO, WARN
from cast_common.util import find_nth, no_dups, list_to_text,format_table
from cast_arg.powerpoint import PowerPoint
from cast_common.highlight import Highlight


from pandas import DataFrame
from pptx import Presentation
from pptx.dml.color import RGBColor
from IPython.display import display
from os import getcwd
from os.path import abspath,dirname,exists
from site import getsitepackages

import pandas as pd
import numpy as np 
import math
import json
import datetime

__author__ = "Nevin Kaplan"
__email__ = "n.kaplan@castsoftware.com"
__copyright__ = "Copyright 2022, CAST Software"

class GeneratePPT(Logger):
    _ppt = None
    _aip_data = None
    _effort_df = None
    _output_folder = None
    _project_name = None
    _template = None

    _generate_HL = _hl_base_url = _hl_user = hl_pswd = None
    _generate_AIP = _aip_base_url = _aip_user = _aip_pswd = None
    _hl_instance = None
    _hl_apps_df = pd.DataFrame()
    _hl_app_list = []

    day_rate=1600

    def __init__(self, config:Config):
        super().__init__("generate",config.logging_generate)
        self._config = config

        out = f"{config.output}/Project {config.project} - Tech DD Findings.pptx"
        self.info(f'Generating {out}')

        self._ppt = PowerPoint(config.template, out)

        # TODO: Handle cases where on HL data is needed and not AIP.
        self.hl_pages = []
        if config.aip_active:
            self.info("Collecting AIP Data")
            self._aip_data = AipData(config,log_level=config.logging_aip)
        if config.hl_active:
            self.info("Collecting Highlight Data")
            hl = Highlight(hl_base_url=config.hl_url,hl_user=config.hl_user,hl_pswd=config.hl_password, \
                           hl_instance=config.hl_instance,hl_apps=config.hl_list)
            self.hl_pages = [
                HighlightSummary(self.day_rate),
                GreenIt()
            ]
            self._hl_data = HLData(config,log_level=config.logging_highlight)

        #project level work
        app_cnt = len(config.application)

        # self.remove_proc_slides(self._generate_procs)

        self._ppt.duplicate_slides(app_cnt)
        self._ppt.copy_block("each_app",["app"],app_cnt)

        self._ppt.replace_text("{app#_","{app1_")

        self.replace_all_text()

    def remove_proc_slides(self,keep_it):
        indexes=[]
        for idx, slide in enumerate(self._ppt._prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text == "{proc_analysis}":
                            self._ppt.replace_slide_text(slide,"{proc_analysis}","")
                            indexes.append(idx)

        for idx in reversed(indexes):
            if not keep_it:
                self._ppt.delete_slide(idx)

    def save_ppt(self):
        self._ppt.save()

    def replace_all_text(self):
        app_cnt = len(self._config.application)
        self._ppt.replace_text("{project}", self._config.project)
        self._ppt.replace_text("{app_count}",app_cnt)
        self._ppt.replace_text("{company}",self._config.company)

        mydate = datetime.datetime.now()
        month = mydate.strftime("%B")
        year = f'{mydate.year} '
        self._ppt.replace_text("{month}",month)
        self._ppt.replace_text("{year}",year)

        # replace AIP data global to all applications
        if self._config.aip_active:
            all_apps_avg_grade = self._aip_data.calc_grades_all_apps()
#            self._ppt.replace_text("{all_apps}",self._aip_data.get_all_app_text())
            self._ppt.replace_risk_factor(all_apps_avg_grade,search_str="{summary_")
            risk_grades = self._aip_data.calc_health_grades_high_risk(all_apps_avg_grade)
            if risk_grades.empty:
                risk_grades = self._aip_data.calc_health_grades_medium_risk(all_apps_avg_grade)
            self._ppt.replace_text("{summary_at_risk_factors}",list_to_text(risk_grades.index.tolist()).lower())
        
            if risk_grades.empty:
                self._ppt.replace_block('{immediate_action}','{end_immediate_action}','')
            else:
                self._ppt.replace_text("{immediate_action}","") 
                self._ppt.replace_text("{end_immediate_action}","") 
                self._ppt.replace_text("{high_risk_grade_names}",self._aip_data.text_from_list(risk_grades.index.values.tolist()))

        app_list = self._config.aip_list
        hl_list = self._config.hl_list

        day_rate = self.day_rate

        summary_near_term=AIPStats(day_rate,logger_level=self._config.logging_generate)
        summary_fix_now=AIPStats(day_rate,logger_level=self._config.logging_generate)
        summary_mid_long_term=AIPStats(day_rate,logger_level=self._config.logging_generate)
        hl_summary=OssStats('',day_rate,logger_level=self._config.logging_generate)
        hl_summary_critical=OssStats('',day_rate,logger_level=self._config.logging_generate)
        hl_summary_high_near=OssStats('',day_rate,logger_level=self._config.logging_generate)
        lic_summary = LicenseStats(logger_level=self._config.logging_generate)
        summary_components = 0

        self._ppt.replace_text('{app_cnt}',app_cnt)
        for idx in range(0,app_cnt):
            # create instance of action plan class 
            self.ap = ActionPlan (app_list,self._aip_data,self._ppt,self._config.output,day_rate)
            fix_now_total=AIPStats(day_rate,logger_level=self._config.logging_generate)
            near_term_total=AIPStats(day_rate,logger_level=self._config.logging_generate)
            mid_long_term=AIPStats(day_rate,logger_level=self._config.logging_generate)
            summary_total=AIPStats(day_rate,logger_level=self._config.logging_generate)

            hl_near_term_total=AIPStats(day_rate,logger_level=self._config.logging_generate)

            # replace application specific AIP data
            app_no = idx+1
            app_id = app_list[idx]
            hl_id = hl_list[idx]
            app_title = self._config.title_list[idx]
            self.info(f'********************* Working on pages for {app_title} ******************************')
            self._ppt.replace_text(f'{{app{app_no}_name}}',app_title)

            if self._config.hl_active:
                for proc in self.hl_pages:
                    proc.report(hl_id,app_no,self._ppt,self._config.output)

            if self._config.aip_active:
                if self._aip_data.has_data(app_id):
                    self.info('Preparing AIP Data')
                    #self.info(f'Working on {app_id} ({self._appl_title[app_id]})')

                    self.info('Filling risk factors for the executive summary page')
                    # do risk factors for the executive summary page
                    self.fill_aip_grades(self._aip_data,app_id, app_no)
                    risk_grades = self.each_risk_factor(self._aip_data,app_id, app_no)
                    self._ppt.replace_text(f'{{app{app_no}_high_risk_grade_names}}',list_to_text(risk_grades.index.values))


                    self.info('Filling Technical details TABLE')
                    # Technical Overview - Technical details TABLE
                    grade_all = self._aip_data.get_app_grades(app_id)
                    #self._ppt.replace_risk_factor(grade_all,app_no)
                    grade_by_tech_df = self._aip_data.get_grade_by_tech(app_id)
                    grades = grade_by_tech_df.drop(['Documentation',"ISO","ISO_EFF","ISO_MAINT","ISO_REL","ISO_SEC"],axis=1)
                    self._ppt.update_table(f'app{app_no}_grade_by_tech_table',grades)

                    if not grade_by_tech_df.empty:
                        #add appmarq technology
                        self._ppt.replace_text(f'{{app{app_no}_largest_tech}}',grade_by_tech_df.index[0])

                        self.info('Filling Technical Overview')
                        # Technical Overview - Lines of code by technology GRAPH
                        self._ppt.update_chart(f'app{app_no}_sizing_pie_chart',DataFrame(grade_by_tech_df['LOC']))

                    snapshot = self._aip_data.snapshot(app=app_id)
                    # app_name = snapshot['name']
                    # self._ppt.replace_text(f'{{app{app_no}_name}}',app_name)
                    self._ppt.replace_text(f'{{app{app_no}_all_technogies}}',list_to_text(snapshot['technology']))

                    #calculate high and medium risk factors
                    risk_grades = self._aip_data.calc_health_grades_high_risk(grade_all)
                    if risk_grades.empty:
                        risk_grades = self._aip_data.calc_health_grades_medium_risk(grade_all)
                    self._ppt.replace_text(f'{{app{app_no}_at_risk_grade_names}}',list_to_text(risk_grades.index.tolist()).lower())

                    self.fill_strengh_improvement_tbl(app_id,app_no)
                    
                    """
                        Populate the document insites page
                        The necessary data is found in the loc_tbl

                        This section fetches data for Documentation slide, excludes certain columns, sorts with few column, 
                        and based on score of particular element seggregates colors(Red, Yellow & Green) and update to app1_doc_table element.
                        255,168,168 - Red shades
                        255,234,168 - Yellow shades
                        168,228,195 - Green shades
                    """
                    doc_df = self._aip_data.doc_compliance(app_id)
                    doc_df.drop(columns=['Key','Total','Weight'],inplace=True) #'Detail',
                    doc_df.sort_values(by=['Score','Rule'], inplace=True)
                    doc_df['RGB'] = np.where(doc_df.Score >= 3,'194,236,213',np.where(doc_df.Score < 2,'255,210,210','255,240,194'))
                    doc_df.Score = doc_df.Score.map('{:.2f}'.format)
                    self._ppt.update_table(f'app{app_no}_doc_table',doc_df,include_index=False,background='RGB')
                    


                   
                    loc_df = self._aip_data.get_loc_sizing(app_id)
                    if len(loc_df) > 0:
                        loc = loc_df['Number of Code Lines']
                        self._ppt.replace_loc(loc,app_no)

                        loc_tbl = pd.DataFrame.from_dict(data=self._aip_data.get_loc_sizing(app_id),orient='index').drop('Critical Violations')
                        loc_tbl = loc_tbl.rename(columns={0:'loc'})
                        loc_tbl['percent'] = round((loc_tbl['loc'] / loc_tbl['loc'].sum()) * 100,2)
                        loc_tbl['loc']=pd.Series(["{0:,.0f}".format(val) for val in loc_tbl['loc']], index = loc_tbl.index)

                        percent_comment = loc_tbl.loc['Number of Comment Lines','percent']
                        percent_comment_out = loc_tbl.loc['Number of Commented-out Code Lines','percent']

                        if percent_comment < 15:
                            comment_level='low'
                        elif percent_comment > 15 and percent_comment <= 20:
                            comment_level='good'
                        else:
                            comment_level='high'
                    
                        self._ppt.replace_text(f'{{app{app_no}_comment_hl}}',comment_level)
                        self._ppt.replace_text(f'{{app{app_no}_comment_level}}',comment_level)
                        self._ppt.replace_text(f'{{app{app_no}_comment_pct}}',percent_comment)
                        self._ppt.replace_text(f'{{app{app_no}_comment_out_pct}}',percent_comment_out)

                        loc_tbl['percent']=pd.Series(["{0:.2f}%".format(val) for val in loc_tbl['percent']], index = loc_tbl.index)
                        self._ppt.update_table(f'app{app_no}_loc_table',loc_tbl,has_header=False)
                        self._ppt.update_chart(f'app{app_no}_loc_pie_chart',DataFrame(loc_tbl['loc']))

                        # self._ppt.replace_grade(grade_all,app_no)

                        self.fill_sizing(app_id,app_no)
                        self.fill_violations(app_id,app_no)
    
                    self.fill_critical_rules(app_id,app_no)

                    """
                        Get Action Plan data and combine it for the various slide combinations

                        Executive Summary uses a combination of extreme and high violation data
                        Action Plan midigation slice is divided into three sections
                            1. Fix before sigining contains only extreme violation data
                            2. Near term contains high violation data
                            3. Long term contains medium and low violation data
                            4. Excecutive Summary contains fix now and high violation data
                    """
                    self.ap.fill_action_plan(app_id,app_no)
                    fix_now_total.add_effort(self.ap.fix_now.effort)
                    fix_now_total.add_violations(self.ap.fix_now.violations)

                    near_term_total.add_effort(self.ap.high.effort)
                    near_term_total.add_violations(self.ap.high.violations)

                    mid_long_term.add_effort(self.ap.medium.effort)
                    mid_long_term.add_violations(self.ap.medium.violations)
                    mid_long_term.add_data(self.ap.medium.data)
                    mid_long_term.add_effort(self.ap.low.effort)
                    mid_long_term.add_violations(self.ap.low.violations)
                    mid_long_term.add_data(self.ap.low.data)

                    # summary_near_term.add_effort(self.ap.high.effort)
                    # summary_near_term.add_effort(self.ap.medium.effort)
                    # summary_near_term.add_violations(self.ap.high.violations)
                    # summary_near_term.add_violations(self.ap.medium.violations)

                    """
                        ISO-5055 slide
                            - most of the work is being done in the restCall class
                            - use iso_rules to retrieve the data
                            - add it to the app1_iso5055 table in the template
                    """
                    iso_df = self._aip_data.iso_rules(app_id)
                    if not iso_df.empty:
                        iso_df.loc[iso_df['violation']=='','background']='205,218,226'
                        iso_df.loc[iso_df['violation']!='','background']='255,255,255'
                        self._ppt.update_table(f'app{app_no}_iso5055',iso_df,
                                            include_index=False,background='background')
                                       
                        pourcentage_iso5055 = iso_df["count"].sum()
                        
                        iso_Maintainaility = iso_df[iso_df.category == 'Maintainability' ]
                        iso_MaintainailityCall = iso_Maintainaility["count"].sum()
                        self._ppt.replace_text(f'{{app{app_no}_ISO_5055}}', round((iso_MaintainailityCall/(pourcentage_iso5055/2))*100,1))
                    
            #replaceHighlight application specific data
            if self._config.hl_active and self._hl_data.has_data(hl_id):
                try:
                    (oss_crit,oss_high,oss_med,lic,components) = self.oss_risk_assessment(hl_id,app_no,day_rate)
                    fix_now_total.add_effort(oss_crit.effort)
                    fix_now_total.add_violations(oss_crit.violations)

                    near_term_total.add_effort(oss_high.effort)
                    near_term_total.add_effort(oss_med.effort)

                    hl_near_term_total.add_effort(oss_high.effort)
                    hl_near_term_total.add_effort(oss_med.effort)
                    hl_near_term_total.add_violations(oss_high.violations)
                    hl_near_term_total.add_violations(oss_med.violations)
                    hl_near_term_total.replace_text(self._ppt,app_no,'hl_near_term_total')

                    hl_summary_critical.add_components(oss_crit.components)
                    hl_summary_critical.add_violations(oss_crit.violations)

                    hl_summary_high_near.add_components(oss_high.components)
                    hl_summary_high_near.add_components(oss_med.components)

                    hl_summary.add_components(components)

#                    if not lic.empty:
                    lic_summary.add_high(lic.high)
                    lic_summary.add_medium(lic.medium)
                    lic_summary.add_low(lic.low)


                except KeyError as ex:
                    self.warning(f'OSS information not found {str(ex)}')
                """
                    Cloud ready excel sheet generation
                """
                try:
                    cloud = self._hl_data.get_cloud_info(hl_id)
                    cloud = cloud[['cloudRequirement.display','Technology','cloudRequirement.ruleType','cloudRequirement.criticality','contributionScore','roadblocks']]
                    file_name = f'{self._config.output}/cloud-{self._config.title_list[idx]}.xlsx'
                    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
                    col_widths=[50,10,10,10,10,10,10]
                    cloud_tab = format_table(writer,cloud,'Cloud Data',col_widths)
                    writer.close()
                except Exception as e:
                    self.error(f'unknown error while processing cloud ready data: {str(e)}')


            summary_fix_now.add_effort(fix_now_total.effort)
            summary_fix_now.add_violations(fix_now_total.violations)
            summary_near_term.add_effort(near_term_total.effort)
            summary_near_term.add_violations(near_term_total.violations)
            summary_mid_long_term.add_effort(mid_long_term.effort)
            summary_mid_long_term.add_violations(mid_long_term.violations)

            summary_total.add_effort(summary_fix_now.effort)
            summary_total.add_effort(summary_near_term.effort)
            summary_total.add_effort(summary_mid_long_term.effort)
            summary_total.add_violations(summary_fix_now.violations)
            summary_total.add_violations(summary_near_term.violations)
            summary_total.add_violations(summary_mid_long_term.violations)
            summary_total.add_data(summary_fix_now.data)
            summary_total.add_data(summary_near_term.data)
            summary_total.add_data(summary_mid_long_term.data)
            summary_total.replace_text(self._ppt,app_no,'summary_total')


            #replace text for all aip and HL statistics in the powerpoint document
            self.ap.fix_now.replace_text(self._ppt,app_no,'fix_now')
            self.ap.high.replace_text(self._ppt,app_no,'high')
            self.ap.medium.replace_text(self._ppt,app_no,'medium')
            self.ap.low.replace_text(self._ppt,app_no,'low')

            mid_long_term.replace_text(self._ppt,app_no,'mid_long_term')

            fix_now_total.replace_text(self._ppt,app_no,'fix_now_total')
            near_term_total.replace_text(self._ppt,app_no,'near_term_total')

            # short_term.replace_text(self._ppt,app_no,'st')
            # long_term.replace_text(self._ppt,app_no,'lt')
        
        #summary_fix_now_aip_hl.replace_text(self._ppt,'','hl_summary_fix_now')
        summary_fix_now.replace_text(self._ppt,'','summary_fix_now')
        summary_near_term.replace_text(self._ppt,'','summary_near_term')

        hl_summary_critical.replace_text(self._ppt,'_summary_crit')
        hl_summary_high_near.replace_text(self._ppt,'_summary_high_near')
        lic_summary.replace_text(self._ppt,'_summary')
        hl_summary.replace_text(self._ppt,'_summary')
        self._ppt.replace_text('{app_summary_comp_tot}',summary_components)

        show_stopper_flg = 'some'
        if self.ap.fix_now.effort == 0:
            show_stopper_flg = 'no'
        self._ppt.replace_text('{show_stopper_flg}',show_stopper_flg)  

        avg_cost = float(summary_fix_now.cost/app_cnt)   
        if avg_cost < 50:
            hml_cost_flg = 'low'
        elif avg_cost > 50 and avg_cost < 100:
            hml_cost_flg = 'medium'
        else:
            hml_cost_flg = 'high'
        self._ppt.replace_text('{hml_cost_flag}',hml_cost_flg)  

        self._ppt.replace_text('{daily_rate}',self.ap._day_rate)  

    def each_risk_factor(self, aip_data, app_id, app_no):
        # collect the high risk grades
        # if there are no high ris grades then use medium risk
        app_level_grades = aip_data.get_app_grades(app_id)
        risk_grades = aip_data.calc_health_grades_high_risk(app_level_grades)
        risk_catagory="high"
        if risk_grades.empty:
            risk_catagory="medium"
            risk_grades = aip_data.calc_health_grades_medium_risk(app_level_grades)
        
        #in the event all health risk factors are low risk
        if risk_grades.empty:
            self._ppt.replace_block(f'{{app{app_no}_risk_detail}}',
                            f'{{end_app{app_no}_risk_detail}}',
                            "no high-risk health factors")
        else: 
            rpl_str = f'{{app{app_no}_risk_category}}'
            self._ppt.replace_text(rpl_str,risk_catagory)
            self.debug(f'replaced {rpl_str} with {risk_catagory}')

            self._ppt.copy_block(f'app{app_no}_each_risk_factor',["_risk_name","_risk_grade"],len(risk_grades.count(axis=1)))
            f=1
            for index, row in risk_grades.T.items():
                rpl_str = f'{{app{app_no}_risk_name{f}}}'
                self._ppt.replace_text(rpl_str,index)
                self.debug(f'replaced {rpl_str} with {index}')

                rpl_str = f'{{app{app_no}_risk_grade{f}}}'
                value = row['All'].round(2)
                self._ppt.replace_text(rpl_str,value)
                self.debug(f'replaced {rpl_str} with {value}')
                f=f+1

        self._ppt.remove_empty_placeholders()
        return risk_grades

    def get_grade_color(self,grade):
        rgb = 0
        if grade > 3:
            rgb = RGBColor(0,176,80) # light green
        elif grade <3 and grade > 2:
            rgb = RGBColor(214,142,48) # yellow
        else:
            rgb = RGBColor(255,0,0) # red
        return rgb

    def fill_aip_grades(self,aip_data, app_id, app_no):
        self.info('Filling AIP grades data')
        app_level_grades = aip_data.get_app_grades(app_id)
        for name, value in app_level_grades.T.items():
            # fill grades
            grade = round(value,2)
            rpl_str = f'{{app{app_no}_grade_{name}}}'
            self._ppt.replace_text(rpl_str,grade)
            self.debug(f'replaced {rpl_str} with {grade}')

            # fill grade risk factor (high, medium or low)
            rpl_str = f'{{app{app_no}_risk_{name}}}'
            risk = ''
            if grade < 2:
                risk = 'high'
            elif grade < 3:
                risk = 'medium'
            else:
                risk = 'low'
            self._ppt.replace_text(rpl_str,risk)
            self.debug(f'replaced {rpl_str} with {risk}')

            #update grade box color and slider postion 
            id_base = f'app{app_no}_grade'
            box_name = f'{id_base}_{name}_box'
            txt_name = f'{id_base}_{name}_text'
            slider_name = f'{id_base}_{name}_slider'
            color = self.get_grade_color(grade)

            for slide in self._ppt._prs.slides:
                box = self._ppt.get_shape_by_name(box_name,slide)
                if not box is None:
                    box.line.color.rgb = color

                txt = self._ppt.get_shape_by_name(txt_name,slide)
                if not txt is None and txt.has_text_frame:
                    paragraphs = txt.text_frame.paragraphs
                    self._ppt.change_paragraph_color(paragraphs[0],color)

                slider = self._ppt.get_shape_by_name(slider_name,slide)
                if not slider is None:
                    self._ppt.update_grade_slider(slider,[grade])




    def fill_strengh_improvement_tbl(self,app_id,app_no):
        self.info('Filling strength and improvment table')
        """
            Populate the strengths and improvement page
            The necessary data is found in the loc_tbl
        """
        imp_df = self._aip_data.tqi_compliance(app_id)
        imp_df.drop(columns=['Weight','Total','Succeeded','Compliance'],inplace=True)
        imp_df.sort_values(by=['Score','Rule'], inplace=True, ascending = False)

        file_name = f'{self._config.output}/health-{self._config.title_list[app_no-1]}.xlsx'
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        col_widths=[50,50,10,10,10]
        cloud_tab = format_table(writer,imp_df,'Health Data',col_widths)
        writer.close()

        imp_df.drop(columns=['Detail'],inplace=True)
        imp_df['RGB'] = np.where(imp_df.Score >= 3,'194,236,213',\
            np.where(imp_df.Score < 2,'255,210,210','255,240,194'))
        imp_df.Score = imp_df.Score.map('{:.2f}'.format)

        #cause_name = abspath(f'{dirname(__file__)}/cause.json')

        cause_name = abspath(f'{getsitepackages()[-1]}/cast_arg/cause.json')
        imp_df['Cause']=''
        with open(cause_name) as json_file:
            tech_data = json.load(json_file)
        imp_df['Cause']=imp_df['Key'].map(tech_data)

        imp_df = imp_df[['Rule','Score','Cause','Failed','RGB']]
        self._ppt.update_table(f'app{app_no}_imp_table',imp_df,include_index=False,background='RGB')

    def fill_critical_rules(self,app_id,app_no):
        self.info('Filling critical rules table')
        rules_df = self._aip_data.critical_rules(app_id)
        if not rules_df.empty:
            rules_df = rules_df[['rulePattern.name','rulePattern.critical']]
            rule_summary_df=rules_df.groupby(['rulePattern.name']).size().reset_index(name='counts').sort_values(by=['counts'],ascending=False)
            rule_summary_df=rule_summary_df.head(5)
            self._ppt.update_table(f'app{app_no}_top_violations',rule_summary_df,include_index=False)
        else:
            self.warning('This application contains no critical violations')
        # if not rules_df.empty:
        #     critical_rule_df = pd.json_normalize(rules_df['rulePattern'])
        #     critical_rule_df = critical_rule_df[['name','critical']]

        #     #pourcentage_iso5055 = critical_rule_df['name']
        #     #self._ppt.replace_text(f'{{app{app_no}_ISO_5055}}',rule_summary_df)

    def fill_violations(self,app_id,app_no):
        self.info('Filling violation table')
        violation_df = pd.DataFrame(self._aip_data.violation_sizing(app_id),index=[0])
        violation_df['Violation Count']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Violation Count']])
        violation_df[' per file']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per file']])
        violation_df[' per kLoC']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per kLoC']])
        violation_df['Complex objects']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Complex objects']])
        violation_df[' With violations']=pd.Series(["{0:,.0f}".format(val) for val in violation_df[' With violations']])
        self._ppt.update_table(f'app{app_no}_violation_sizing',violation_df.transpose())
        self._ppt.replace_text(f'{{app{app_no}_critical_violations}}',violation_df['Violation Count'].loc[0])

        #violation_df['Security']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Security']])
        
        #print('#########################################################################################')
        #print('coucou le resultat est',violation_df['Security'])
        #print('#########################################################################################')
        ######################################################################

        ##code added for is 5055
        #pourcentage_iso5055 = violation_df[' per file'].sum()
        #self._ppt.replace_text(f'{{app{app_no}_ISO_5055}}',violation_df['Violation Count'].sum())
        ######################################################################################
        

    def fill_sizing(self,app_id,app_no):
        self.info('Filling sizing table')
        sizing_df = pd.DataFrame(self._aip_data.tech_sizing(app_id),index=[0])
        sizing_df['LoC']=pd.Series(["{0:,.0f} K".format(val / 1000) for val in sizing_df['LoC']])
        sizing_df['Files']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Files']])
        sizing_df['Classes']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Classes']])
        sizing_df['SQL Artifacts']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['SQL Artifacts']])
        sizing_df['Tables']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Tables']])
        sizing_df = sizing_df.transpose()
        self._ppt.update_table(f'app{app_no}_tech_sizing',sizing_df)

    def oss_risk_assessment(self,hl_id,app_no,day_rate):
        self.info('Filling OSS risk assessment table')
        lic_df=self._hl_data.get_lic_info(hl_id)
        lic_df=self._hl_data.sort_lic_info(lic_df)
        oss_df=self._hl_data.get_cve_info(hl_id)
        # lic_summary = pd.DataFrame(columns=['License Type','Risk Factor','Component Count','Example'])

        oss_crit = OssStats(hl_id,day_rate,self._hl_data,'crit',logger_level=self._config.logging_generate)
        oss_crit_comp_tot = self._hl_data.get_cve_crit_comp_tot(hl_id)

        oss_high = OssStats(hl_id,day_rate,self._hl_data,'high',logger_level=self._config.logging_generate)
        oss_high_comp_tot = self._hl_data.get_cve_high_comp_tot(hl_id)

        oss_med = OssStats(hl_id,day_rate,self._hl_data,'med',logger_level=self._config.logging_generate)
        oss_med_comp_tot = self._hl_data.get_cve_med_comp_tot(hl_id)

        total_components = self._hl_data.get_oss_cmpn_tot(hl_id)

        oss_crit.replace_text(self._ppt,app_no)
        oss_high.replace_text(self._ppt,app_no)
        oss_med.replace_text(self._ppt,app_no)

        self._ppt.replace_text(f'{{app{app_no}_oss_cmpn_tot}}',total_components)
        if not oss_df.empty:
            self._ppt.update_table(f'app{app_no}_HL_table_CVEs',oss_df,include_index=False)

        self.info('Filling OSS license table')
        '''
            License compliance table
        
        '''
        lic_summary=self._hl_data.get_lic_info(hl_id)
        if not lic_summary.empty:
            lic_summary=lic_summary[['component','version','release','license','risk']].drop_duplicates()
            lic_summary.sort_values(['risk','license','component','version'],inplace=True)

            lic_summary = lic_summary.groupby(['risk','license'])['component'].apply(lambda x: ','.join(x)).reset_index()
            lic_summary['comp count']=lic_summary['component'].str.count(',')+1
            
            #remove duplicates but show full count
            lic_summary['component']=lic_summary['component'].map(lambda x: no_dups(x,',',True))
            
            #only show the first 5 components
            lic_summary['component']=lic_summary['component'].map(lambda x: x[:find_nth(x,',',6)])
            lic_summary=lic_summary[['license','risk','comp count','component']]
            lic_summary.sort_values(['risk','license','comp count'],inplace=True)

            #remove low and undefined records from the table
            lic_summary=lic_summary[lic_summary["risk"].str.contains("Low")==False]
            lic_summary=lic_summary[lic_summary["risk"].str.contains("Undefined")==False]

            # are there any records left?
            lic = LicenseStats(logger_level=self._config.logging_generate)
            if not lic_summary.empty:
                #modify the forground color
                lic_summary.loc[lic_summary['risk']=='High','forground']='211,76,76'
                lic_summary.loc[lic_summary['risk']=='Medium','forground']='127,127,127'

                #update the powerpoint table
                self._ppt.update_table(f'app{app_no}_HL_table_lic_risks',lic_summary,include_index=False)
            
                lic.high = lic_summary[lic_summary['risk']=='High']['comp count'].sum()
                lic.medium = lic_summary[lic_summary['risk']=='Medium']['comp count'].sum()
                lic.low = lic_summary[lic_summary['risk']=='Low']['comp count'].sum()

                # #add the high and medium license risk counts to the deck
                # self._ppt.replace_text(f'{{app{app_no}_high_lic_tot}}',
                #     lic_summary[lic_summary['risk']=='High']['comp count'].sum())
                # self._ppt.replace_text(f'{{app{app_no}_med_lic_tot}}',
                #     lic_summary[lic_summary['risk']=='Medium']['comp count'].sum())

            lic.replace_text(self._ppt,app_no)
        else:
            lic = DataFrame()
        return (oss_crit,oss_high,oss_med,lic,total_components)





