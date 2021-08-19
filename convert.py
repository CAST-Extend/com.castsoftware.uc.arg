from logging import info, warn
from pandas.core.frame import DataFrame
from restCall import AipRestCall
from restCall import AipData
from restCall import HLRestCall
from restCall import HLData
from pptx import Presentation
from powerpoint import PowerPoint
from jproperties import Properties 
from pptx.dml.color import RGBColor
from actionPlan import ActionPlan
from logger import Logger
from IPython.display import display


import pandas as pd
import numpy as np 
import util 
import math
import argparse
import json
import datetime


class GeneratePPT(Logger):
    _app_list = []
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

    def __init__(self, config):
        super().__init__("convert")
        self.read_config(config)

        out = f"{self._output_folder}/{self._project_name}.pptx"
        self.info(f'Generating {out}')


        # TODO: Handle cases where on HL data is needed and not AIP.

        if self._generate_AIP:
            self.info("Collecting AIP Data")
            self._aip_data = AipData(self._aip_base_url, self._aip_user, self._aip_pswd, self._app_list)
        if self._generate_HL:
            self.info("Collecting Highlight Data")
            self._hl_data = HLData(self._hl_base_url, self._hl_user, self._hl_pswd, self._hl_instance, self._app_list,self._hl_app_list)

        self._ppt = PowerPoint(self._template, out)
        #project level work
        app_cnt = len(self._app_list)

        self.remove_proc_slides(self._generate_procs)

        self._ppt.duplicate_slides(app_cnt)
        self._ppt.copy_block("each_app",["app"],app_cnt)

        self._ppt.replace_text("{app#_","{app1_")

        self.replace_all_text(app_cnt)

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

    def read_config(self, config):
        """
        Read entries from the config file and save the values in class/instance vars.
        """

        # TODO: handle undefined entries
        self._project_name = config.get('project').data
        self._template = config.get('template').data
        self._app_list = config.get('appl.list').data.strip().split(',')
        self._appl_title = json.loads(config.get('appl.title').data)
        self._hl_app_list = json.loads(config.get('appl.highlight').data)

        #set default values
        for appl in self._app_list:
            #test the application title list, if not found set it to the list value 
            try:
                test = self._appl_title[appl]
            except (KeyError):
                self._appl_title[appl]=appl
            try:
                test = self._hl_app_list[appl]
            except (KeyError):
                self._hl_app_list[appl]=appl

            #test the application highlight list, if not found set it to the list value 
            try:
                test = self._hl_app_list[appl]
            except (KeyError):
                self._hl_app_list[appl]=appl
            try:
                test = self._hl_app_list[appl]
            except (KeyError):
                self._hl_app_list[appl]=appl


        self._output_folder = config.get('output.folder').data

        self._generate_AIP = config.get('output.aip').data
        self._generate_HL = config.get('output.hl').data
        #self._generate_procs = config.get('output.procs').data
        self._generate_procs=False
        
        if self._generate_HL.lower() == 'yes':
            self._generate_HL = True
            self._hl_base_url = config.get('hl.base_url').data
            self._hl_user = config.get('hl.user').data
            self._hl_pswd = config.get('hl.pswd').data
            self._hl_instance = config.get('hl.instance').data
            # self._hl_app_list={}
            # self._hl_app_list = json.loads(config.get('hl.application.list').data)
        else:     
            self._generate_HL = False

        if self._generate_AIP.lower() == 'yes':
            self._generate_AIP = True
            self._aip_base_url = config.get('aip.base_url').data
            self._aip_user = config.get('aip.user').data
            self._aip_pswd = config.get('aip.pswd').data
        else:     
            self._generate_AIP = False

    def replace_all_text(self,app_cnt):
        self._ppt.replace_text("{project}", self._project_name)
        self._ppt.replace_text("{app_count}",app_cnt)

        mydate = datetime.datetime.now()
        month = mydate.strftime("%B")
        self._ppt.replace_text("{month}",month)
        self._ppt.replace_text("{year}",mydate.year)

        # replace AIP data global to all applications
        if self._generate_AIP:
            all_apps_avg_grade = self._aip_data.calc_grades_all_apps()
            self._ppt.replace_text("{all_apps}",self._aip_data.get_all_app_text())
            self._ppt.replace_risk_factor(all_apps_avg_grade,search_str="{summary_")
            risk_grades = self._aip_data.calc_health_grades_high_risk(all_apps_avg_grade)
            if risk_grades.empty:
                risk_grades = self._aip_data.calc_health_grades_medium_risk(all_apps_avg_grade)
            self._ppt.replace_text("{summary_at_risk_factors}",util.list_to_text(risk_grades.index.tolist()).lower())
        
            if risk_grades.empty:
                self._ppt.replace_block('{immediate_action}','{end_immediate_action}','')
            else:
                self._ppt.replace_text("{immediate_action}","") 
                self._ppt.replace_text("{end_immediate_action}","") 
                self._ppt.replace_text("{high_risk_grade_names}",self._aip_data.text_from_list(risk_grades.index.values.tolist()))

        # create instance of action plan class 
        ap = ActionPlan (self._app_list,self._aip_data,self._ppt,self._output_folder)

        for app_no in range(0,app_cnt):
            # replace application specific AIP data
            app_id = self._app_list[app_no]
            self.info(f'Working on pages for {self._appl_title[app_id]}')
            self._ppt.replace_text(f'{{app{app_no+1}_name}}',self._appl_title[app_id])

            fix_now_eff = fix_now_cost = fix_now_vio_cnt = 0
            fix_now_data = DataFrame()
            fix_now_bus_txt = ''
            fix_now_vio_txt = ''

            near_term_eff = near_term_cost = near_term_vio_cnt = 0
            near_term_data = DataFrame()
            near_term_bus_txt = ''
            near_term_vio_txt = ''

            mid_eff = mid_cost = mid_vio_cnt = 0
            mid_data = DataFrame()
            mid_bus_txt = ''
            mid_vio_txt = ''

            low_eff = low_cost = low_vio_cnt = 0
            low_data = DataFrame()
            low_bus_txt = ''
            low_vio_txt = ''

            long_term_eff = long_term_cost = long_term_vio_cnt = 0
            long_term_data = DataFrame()
            long_term_bus_txt = ''
            long_term_vio_txt = ''

            summary_eff = summary_cost = summary_vio_cnt = 0
            summary_data = DataFrame()
            summary_bus_txt = ''
            summary_vio_txt = ''
            summary_total_cost = 0

            if self._generate_AIP:
                if self._aip_data.has_data(app_id):
                    #self.info(f'Working on {app_id} ({self._appl_title[app_id]})')

                    # do risk factors for the executive summary page
                    risk_grades = util.each_risk_factor(self._ppt, self._aip_data,app_id, app_no)
                    self._ppt.replace_text(f'{{app{app_no+1}_high_risk_grade_names}}',util.list_to_text(risk_grades.index.values))

                    grade_all = self._aip_data.get_app_grades(app_id)
                    self._ppt.replace_risk_factor(grade_all,app_no)
                    grade_by_tech_df = self._aip_data.get_grade_by_tech(app_id)

                    self._ppt.update_table(f'app{app_no+1}_grade_by_tech_table',grade_by_tech_df.drop(['Documentation'],axis=1))
                    self._ppt.update_chart(f'app{app_no+1}_sizing_pie_chart',grade_by_tech_df['LOC'])

                    snapshot = self._aip_data.snapshot(app=app_id)
                    # app_name = snapshot['name']
                    # self._ppt.replace_text(f'{{app{app_no+1}_name}}',app_name)
                    self._ppt.replace_text(f'{{app{app_no+1}_all_technogies}}',util.list_to_text(snapshot['technology']))

                    #calculate high and medium risk factors
                    risk_grades = self._aip_data.calc_health_grades_high_risk(grade_all)
                    if risk_grades.empty:
                        risk_grades = self._aip_data.calc_health_grades_medium_risk(grade_all)
                    self._ppt.replace_text(f'{{app{app_no+1}_at_risk_grade_names}}',util.list_to_text(risk_grades.index.tolist()).lower())


                    """
                        Populate the strengths and improvement page
                        The necessary data is found in the loc_tbl
                    """
                    imp_df = self._aip_data.tqi_compliance(app_id)
                    imp_df.drop(columns=['Key','Total','Weight'],inplace=True)
                    imp_df.sort_values(by=['Score','Rule'], inplace=True)
                    imp_df['RGB'] = np.where(imp_df.Score >= 3,'168,228,195',\
                        np.where(imp_df.Score < 2,'255,168,168','255,234,168'))
                    imp_df.Score = imp_df.Score.map('{:.2f}'.format)
                    self._ppt.update_table(f'app{app_no+1}_imp_table',imp_df,include_index=False,background='RGB')


                    """
                        Populate the document insites page
                        The necessary data is found in the loc_tbl
                    """
                    loc_df = self._aip_data.get_loc_sizing(app_id)
                    loc = loc_df['Number of Code Lines']
                    self._ppt.replace_loc(loc,app_no+1)

                    loc_tbl = pd.DataFrame.from_dict(data=self._aip_data.get_loc_sizing(app_id),orient='index').drop('Critical Violations')
                    loc_tbl = loc_tbl.rename(columns={0:'loc'})
                    loc_tbl['percent'] = round((loc_tbl['loc'] / loc_tbl['loc'].sum()) * 100,2)
                    loc_tbl['loc']=pd.Series(["{0:,.0f}".format(val) for val in loc_tbl['loc']], index = loc_tbl.index)

                    percent_comment = loc_tbl.loc['Number of Comment Lines','percent']
                    percent_comment_out = loc_tbl.loc['Number of Commented-out Code Lines','percent']
                    if percent_comment < 15:
                        comment_level='low'
                    else:
                        comment_level='good'
                    self._ppt.replace_text(f'{{app{app_no+1}_comment_level}}',comment_level)
                    self._ppt.replace_text(f'{{app{app_no+1}_comment_pct}}',percent_comment)
                    self._ppt.replace_text(f'{{app{app_no+1}_comment_out_pct}}',percent_comment_out)

                    loc_tbl['percent']=pd.Series(["{0:.2f}%".format(val) for val in loc_tbl['percent']], index = loc_tbl.index)
                    self._ppt.update_table(f'app{app_no+1}_loc_table',loc_tbl,has_header=False)
                    self._ppt.update_chart(f'app{app_no+1}_loc_pie_chart',loc_tbl['loc'])

                    self._ppt.replace_grade(grade_all,app_no+1)

                    self.fill_sizing(app_no)
                    self.fill_critical_rules(app_no)


                    """
                        Get Action Plan data and combine it for the various slide combinations

                        Executive Summary uses a combination of extreme and high violation data
                        Action Plan midigation slice is divided into three sections
                            1. Fix before sigining contains only extreme violation data
                            2. Near term contains high violation data
                            3. Long term contains medium and low violation data
                            4. Excecutive Summary contains fix now and high violation data
                    """
                    ap.fill_action_plan(app_no)
                    self.fill_violations(app_no)

                    #calculate and print aip fix now action plan midigation items
                    (aip_fix_now_eff, aip_fix_now_cost, aip_fix_now_vio_cnt, aip_fix_now_data) = \
                        ap.get_extreme_costing()
                    aip_fix_now_bus_txt = util.list_to_text(ap.business_criteria(aip_fix_now_data)) + ' '
                    aip_fix_now_vio_txt = ap.list_violations(aip_fix_now_data)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_fn_eff}}',aip_fix_now_eff)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_fn_cost}}',aip_fix_now_cost)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_fn_vio_cnt}}',aip_fix_now_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_fn_bus_txt}}',aip_fix_now_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_fn_vio_txt}}',aip_fix_now_vio_txt)
                        

                    (aip_near_term_eff, aip_near_term_cost, aip_near_term_vio_cnt, aip_near_term_data) = \
                        ap.get_high_costing()
                    aip_near_term_bus_txt = util.list_to_text(ap.business_criteria(aip_near_term_data)) + ' '
                    aip_near_term_vio_txt = ap.list_violations(near_term_data)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_eff}}',aip_near_term_eff)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_cost}}',aip_near_term_cost)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_vio_cnt}}',aip_near_term_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_bus_txt}}',aip_near_term_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_vio_txt}}',aip_near_term_vio_txt)


                    (aip_mid_term_eff, aip_mid_term_cost, aip_mid_term_vio_cnt, aip_mid_term_data) = ap.get_med_costing()
                    aip_mid_term_bus_txt = util.list_to_text(ap.business_criteria(aip_mid_term_data)) + ' '
                    aip_mid_term_vio_txt = ap.list_violations(aip_mid_term_data)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_mt_eff}}',aip_mid_term_eff)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_mt_cost}}',aip_mid_term_cost)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_mt_vio_cnt}}',aip_mid_term_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_mt_bus_txt}}',aip_mid_term_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_mt_vio_txt}}',aip_mid_term_vio_txt)

                    (aip_low_eff, aip_low_cost, aip_low_vio_cnt, aip_low_data) = ap.get_low_costing()
                    aip_low_bus_txt = util.list_to_text(ap.business_criteria(aip_low_data)) + ' '
                    aip_low_vio_txt = ap.list_violations(aip_low_data)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_eff}}',aip_low_eff)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_cost}}',aip_low_cost)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_vio_cnt}}',aip_low_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_bus_txt}}',aip_low_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_nt_vio_txt}}',aip_low_vio_txt)

                    aip_long_term_data = pd.concat([low_data,mid_data],ignore_index=True)
                    aip_long_term_bus_txt = util.list_to_text(ap.business_criteria(aip_long_term_data)) + ' '
                    aip_long_term_vio_txt = ap.list_violations(aip_long_term_data)
                    aip_long_term_eff = int(mid_eff) + int(low_eff)
                    aip_long_term_cost = float(mid_cost) + float(low_cost)
                    aip_long_term_vio_cnt = int(mid_vio_cnt) + int(low_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_lt_eff}}',aip_long_term_eff)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_lt_cost}}',aip_long_term_cost)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_lt_vio_cnt}}',aip_long_term_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_lt_bus_txt}}',aip_long_term_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_aip_lt_vio_txt}}',aip_long_term_vio_txt)

                    summary_data = pd.concat([aip_fix_now_data,aip_near_term_data],ignore_index=True)
                    summary_bus_txt = util.list_to_text(ap.business_criteria(summary_data)) + ' '
                    summary_vio_txt = ap.list_violations(summary_data)
                    summary_eff = int(aip_fix_now_eff) + int(aip_near_term_eff)
                    summary_cost = float(aip_fix_now_cost) + float(aip_near_term_cost)
                    summary_vio_cnt = int(aip_fix_now_vio_cnt) + int(aip_near_term_vio_cnt)

            #replaceHighlight application specific data
            if self._generate_HL and self._hl_data.has_data(app_id):
                lic_df=self._hl_data.get_lic_info(app_id)
                lic_df=self._hl_data.sort_lic_info(lic_df)
                oss_df=self._hl_data.get_cve_info(app_id)
                lic_summary = pd.DataFrame(columns=['License Type','Risk Factor','Component Count','Example'])

                crit_cve = self._hl_data.get_cve_crit_tot(app_id)
                crit_comp_tot = self._hl_data.get_cve_crit_comp_tot(app_id)

                high_cve = self._hl_data.get_cve_high_tot(app_id)
                high_comp_tot = self._hl_data.get_cve_high_comp_tot(app_id)

                med_cve = self._hl_data.get_cve_med_tot(app_id)
                med_comp_tot = self._hl_data.get_cve_med_comp_tot(app_id)

                oss_cmpnt_tot = self._hl_data.get_oss_cmpn_tot(app_id)

                self._ppt.replace_text(f'{{app{app_no+1}_crit_sec_tot}}',crit_cve)
                self._ppt.replace_text(f'{{app{app_no+1}_high_sec_tot}}',high_cve)
                self._ppt.replace_text(f'{{app{app_no+1}_med_sec_tot}}',med_cve)

                self._ppt.replace_text(f'{{app{app_no+1}_crit_cve_comp_ct}}',crit_comp_tot)
                self._ppt.replace_text(f'{{app{app_no+1}_high_cve_comp_ct}}',high_comp_tot)
                self._ppt.replace_text(f'{{app{app_no+1}_med_cve_comp_ct}}',med_comp_tot)

                self._ppt.replace_text(f'{{app{app_no+1}_high_lic_tot}}',self._hl_data.get_lic_high_tot(app_id))
                self._ppt.replace_text(f'{{app{app_no+1}_oss_cmpn_tot}}',oss_cmpnt_tot)

                if not lic_df.empty:
                    for ln in lic_df['license'].unique():
                        data={}
                        data['License Type']=ln

                        lic_type=lic_df[lic_df['license']==ln]
                        lic_cnt = len(lic_type)
                        data['Component Count']=lic_cnt
                        if lic_cnt > 0:
                            data['Risk Factor']=lic_type.iloc[0]['compliance']
                            data['Example']=", ".join(lic_type.head(3)['component'].tolist())
                            lic_summary=lic_summary.append(data,ignore_index=True)

                    #app1_HL_table_lic_risks
                    if len(lic_summary.loc[lic_summary['Risk Factor']=='High'])>0:
                        lic_summary.loc[lic_summary['Risk Factor']=='High','forground']='255,0,0'
                    
                    if len(lic_summary.loc[lic_summary['Risk Factor']=='Medium'])>0:
                        lic_summary.loc[lic_summary['Risk Factor']=='Medium','forground']='209,125,13'

                    self._ppt.update_table(f'app{app_no+1}_HL_table_lic_risks',
                                        lic_summary,include_index=False,
                                        forground='forground')
                
                    high_lic_total = len(lic_summary.loc[lic_summary['Risk Factor']=='High'])
                else:
                    self.info('No license risks found')


                self._ppt.update_table(f'app{app_no+1}_HL_table_CVEs',oss_df,include_index=False)

                if crit_cve is None:
                    crit_cve_eff = 0
                    crit_cve_cost = 0
                else:
                    crit_cve_eff = math.ceil(crit_cve/2)
                    crit_cve_cost = crit_cve_eff * ap._day_rate /1000

                if high_cve is None:
                    high_cve_eff = 0
                    high_cve_cost = 0
                else:
                    high_cve_eff = math.ceil(high_cve/2)
                    high_cve_cost = high_cve_eff * ap._day_rate /1000

                if med_cve is None:
                    med_cve_eff = 0
                    med_cve_cost = 0
                else:   
                    med_cve_eff = math.ceil(med_cve/2)
                    med_cve_cost = med_cve_eff * ap._day_rate /1000

                summary_eff = int(summary_eff) + int(crit_cve_eff)
                summary_cost = round(summary_cost + crit_cve_cost,2)

                fix_now_eff = int(aip_fix_now_eff) + int(crit_cve_eff)
                fix_now_cost = round(aip_fix_now_cost + crit_cve_cost,2)

                near_term_eff = int(aip_near_term_eff) + int(high_cve_eff) + int(med_cve_eff)
                near_term_cost = round(aip_near_term_cost + high_cve_cost + med_cve_cost,2)

                self._ppt.replace_text(f'{{app{app_no+1}_high_sec_tot}}','{high_cve}')
                self._ppt.replace_text(f'{{app{app_no+1}_med_sec_tot}}','{mid__cve}')
                self._ppt.replace_text(f'{{app{app_no+1}_hl_fn_eff}}',crit_cve_eff)
                self._ppt.replace_text(f'{{app{app_no+1}_hl_nt_eff}}',int(high_cve_cost + med_cve_cost))

                self._ppt.replace_text(f'{{app{app_no+1}_fn_tot_cost}}',fix_now_cost)
                self._ppt.replace_text(f'{{app{app_no+1}_fn_tot_eff}}',fix_now_eff)

                self._ppt.replace_text(f'{{app{app_no+1}_nt_tot_cost}}',near_term_cost)
                self._ppt.replace_text(f'{{app{app_no+1}_nt_tot_eff}}',near_term_eff)

            # if not self._generate_AIP and self._generate_HL:
            #     #This deck is for HL only, lets make some adjustments
            #     fix_now_bus_txt=near_term_bus_txt=long_term_bus_txt=None

            #both AIP and HL data
            # ap.fill_action_plan_tags(app_no,'fix_now', \
            #     fix_now_eff, fix_now_cost, fix_now_vio_cnt,fix_now_bus_txt,fix_now_vio_txt)
            # ap.fill_action_plan_tags(app_no,'near_term', \
            #     near_term_eff, near_term_cost, near_term_vio_cnt,near_term_bus_txt,near_term_vio_txt)
            # ap.fill_action_plan_tags(app_no,'long_term', \
            #     long_term_eff, long_term_cost, long_term_vio_cnt,long_term_bus_txt,long_term_vio_txt)
            # ap.fill_action_plan_tags(app_no,'mid', mid_eff, mid_cost, mid_vio_cnt,mid_bus_txt,mid_vio_txt)
            # ap.fill_action_plan_tags(app_no,'low', low_eff, low_cost, low_vio_cnt,low_bus_txt,low_vio_txt)

            # ap.fill_action_plan_tags(app_no,'summary', \
            #     summary_eff, summary_cost, summary_vio_cnt,summary_bus_txt,summary_vio_txt)

            # ap.fill_action_plan_tags(app_no,'summary', \
            #     summary_eff, summary_cost, summary_vio_cnt,summary_bus_txt,summary_vio_txt)

            summary_total_cost = summary_total_cost + summary_cost
            self._ppt.replace_text(f'{{app{app_no+1}_summary_eff}}',summary_eff)
            self._ppt.replace_text(f'{{app{app_no+1}_summary_cost}}',summary_cost)
            self._ppt.replace_text(f'{{app{app_no+1}_summary_vio_cnt}}',summary_vio_cnt)
            self._ppt.replace_text(f'{{app{app_no+1}_summary_bus_txt}}',summary_bus_txt)
            self._ppt.replace_text(f'{{app{app_no+1}_summary_vio_txt}}',summary_vio_txt)

        self._ppt.replace_text('{summary_total_cost}',summary_total_cost)  
        if fix_now_eff > 0:
            show_stopper_flg = 'no'
        else:
            show_stopper_flg = ''
        self._ppt.replace_text('{show_stopper_flg}',show_stopper_flg)  

        avg_cost = float(summary_total_cost/app_cnt)   
        if avg_cost < 50:
            hml_cost_flg = 'low'
        elif avg_cost > 50 and avg_cost < 100:
            hml_cost_flg = 'medium'
        else:
            hml_cost_flg = 'high'
        self._ppt.replace_text('{hml_cost_flag}',hml_cost_flg)  


    def fill_critical_rules(self,app_no):
        app_id = self._app_list[app_no]
        rules_df = self._aip_data.critical_rules(app_id)
        if not rules_df.empty:
            critical_rule_df = pd.json_normalize(rules_df['rulePattern'])
            critical_rule_df = critical_rule_df[['name','critical']]
            rule_summary_df=critical_rule_df.groupby(['name']).size().reset_index(name='counts').sort_values(by=['counts'],ascending=False)
            rule_summary_df=rule_summary_df.head(10)
            self._ppt.update_table(f'app{app_no+1}_top_violations',rule_summary_df,include_index=False)

    def fill_violations(self,app_no):
        app_id = self._app_list[app_no]
        violation_df = pd.DataFrame(self._aip_data.violation_sizing(app_id),index=[0])
        violation_df['Violation Count']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Violation Count']])
        violation_df[' per file']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per file']])
        violation_df[' per kLoC']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per kLoC']])
        violation_df['Complex objects']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Complex objects']])
        violation_df[' With violations']=pd.Series(["{0:,.0f}".format(val) for val in violation_df[' With violations']])
        self._ppt.update_table(f'app{app_no+1}_violation_sizing',violation_df.transpose())
        self._ppt.replace_text(f'{{app{app_no+1}_critical_violations}}',violation_df['Violation Count'].loc[0])


    def fill_sizing(self,app_no):
        app_id = self._app_list[app_no]

        sizing_df = pd.DataFrame(self._aip_data.tech_sizing(app_id),index=[0])
        sizing_df['LoC']=pd.Series(["{0:,.0f} K".format(val / 1000) for val in sizing_df['LoC']])
        sizing_df['Files']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Files']])
        sizing_df['Classes']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Classes']])
        sizing_df['SQL Artifacts']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['SQL Artifacts']])
        sizing_df['Tables']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Tables']])
        sizing_df = sizing_df.transpose()
        self._ppt.update_table(f'app{app_no+1}_tech_sizing',sizing_df)

if __name__ == '__main__':
    print('\nCAST Assessment Deck Generation Tool')
    print('Copyright (c) 2021 CAST Software Inc.\n')
    print('If you need assistance, please contact Nevin Kaplan (NKA) or Guru Pai (GPR) from the CAST US PS team\n')

    parser = argparse.ArgumentParser(description='Assessment Deck Generation Tool')
    parser.add_argument('-c','--config', required=True, help='Configuration properties file')
    args = parser.parse_args()

    config = Properties()
    with open(args.config, 'rb') as config_file:
        config.load(config_file)

    #GeneratePPT(config)

    ppt = GeneratePPT(config)

    """
    # Retreive HL data and generated HL specific slides.
    # TODO: Only if HL generation is enabled.
    if ppt._generate_HL:
        ppt.get_hl_data()
    """

    ppt.save_ppt()

    """
    if generate_AIP:
        if generate_HL:
            GeneratePPT(project_name, app_list, template_name, output_folder, gen_hl = True, hl_user = hl_user, hl_pswd = hl_pswd)
        else:
            # TODO: Pass in AIP rest URL, too.
            GeneratePPT(project_name, app_list, template_name, output_folder)
    else:
        pass
    """


