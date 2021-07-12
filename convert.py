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

MAX_LICS_PER_SLIDE = 12
MAX_CVES_PER_SLIDE = 12
class GeneratePPT(Logger):
    _app_list = []
    _ppt = None
    _aip_data = None
    _effort_df = None
    _output_folder = None
    _project_name = None
    _template = None

    _generate_HL, _hl_base_url, _hl_user, hl_pswd = None, None, None, None
    _generate_AIP, _aip_base_url, _aip_user, _aip_pswd = None, None, None, None
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

        self._ppt.duplicate_slides(app_cnt)
        self._ppt.copy_block("each_app",["app"],app_cnt)

        self._ppt.replace_text("{app#_","{app1_")

        self.replace_all_text(app_cnt)

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
            self._ppt.replace_text(f'{{app{app_no+1}_name}}',self._appl_title[app_id])

            extrm_effort = 0
            extrm_vio_cnt = 0
            extrm_vio_amt = 0
            extrm_violation_business_txt = ''

            high_effort = 0
            high_vio_cnt = 0
            high_vio_amt = 0
            high_violation_business_txt = ''

            crit_cve = 0

            summary_days = 0
            summary_cost = 0.0
            summary_total_cost = 0.0

            fix_now_days = 0
            fix_now_cost = 0.0

            if self._generate_AIP:
                if self._aip_data.has_data(app_id):
                    self.info(f'Working on {app_id} ({self._appl_title[app_id]})')

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

                    loc_df = self._aip_data.get_loc_sizing(app_id)
                    loc = loc_df['Number of Code Lines']
                    self._ppt.replace_loc(loc,app_no+1)

                    """
                        Populate the document insites page
                        The necessary data is found in the loc_tbl
                    """
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
                    self._ppt.replace_text(f'{{app{app_no+1}_comment_loc}}',percent_comment)

                    loc_tbl['percent']=pd.Series(["{0:.2f}%".format(val) for val in loc_tbl['percent']], index = loc_tbl.index)
                    self._ppt.update_table(f'app{app_no+1}_loc_table',loc_tbl,has_header=False)
                    self._ppt.update_chart(f'app{app_no+1}_loc_pie_chart',loc_tbl['loc'])

                    self._ppt.replace_grade(grade_all,app_no+1)

                    self.fill_sizing(app_no)
                    self.fill_critical_rules(app_no)


                    """
                        Get Action Plan data
                    """
                    ap.fill_action_plan(app_no)
                    self.fill_violations(app_no)

                    (extrm_effort, extrm_cost, extrm_vio_cnt, extrm_data) = \
                        ap.get_extreme_costing()

                    (high_effort, high_cost, high_vio_cnt, high_data) = \
                        ap.get_high_costing()

                    (med_effort, med_cost, med_vio_cnt, med_data) = \
                        ap.get_med_costing()

                    (low_effort, low_cost, low_vio_cnt, low_data) = \
                        ap.get_low_costing()

                    """
                        Now combine it for the various slide combinations

                        Executive Summary uses a combination of extreme and high violation data
                        Action Plan midigation slice is divided into three sections
                            1. Fix before sigining contains only extreme violation data
                            2. Near term contains high and medium violation data
                            3. Long term contains low violation data
                    """
                    #Executive Summary Side
                    summary_data = pd.concat([extrm_data,high_data],ignore_index=True)
                    summary_business_txt = util.list_to_text(ap.business_criteria(summary_data))
                    summary_days = int(extrm_effort) + int(high_effort)
                    summary_cost = float(extrm_cost) + float(high_cost)
                    summary_vio_cnt = int(extrm_vio_cnt) + int(high_vio_cnt)
                    
                    self._ppt.replace_text(f'{{app{app_no+1}_summary_vio_cnt}}',summary_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_summary_bus_text}}',summary_business_txt)

                    # Action Plan mitigation slide
                    fix_now_business_txt = util.list_to_text(ap.business_criteria(extrm_data))
                    fix_now_days = extrm_effort
                    fix_now_cost = extrm_cost
                    fix_now_vio_cnt = extrm_vio_cnt

                    self._ppt.replace_text(f'{{app{app_no+1}_fix_now_bus_crit}}',fix_now_business_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_fix_now_vio_cnt}}',fix_now_vio_cnt)

                    near_term_data = pd.concat([high_data,med_data],ignore_index=True)
                    near_term_bus_txt = util.list_to_text(ap.business_criteria(near_term_data))
                    near_term_days = int(high_effort) + int(med_effort)
                    near_term_cost = float(high_cost) + float(med_cost)
                    near_term_vio_cnt = int(high_vio_cnt) + int(med_vio_cnt)
                    near_term_vio_txt = ap.list_violations(near_term_data)

                    self._ppt.replace_text(f'{{app{app_no+1}_near_term_bus_txt}}',near_term_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_near_term_vio_cnt}}',near_term_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_near_term_vio_text}}',near_term_vio_txt)

                    long_term_bus_txt = util.list_to_text(ap.business_criteria(low_data))
                    long_term_days = low_effort
                    long_term_cost = low_cost
                    long_term_vio_cnt = low_vio_cnt
                    long_term_vio_txt = ap.list_violations(low_data)

                    self._ppt.replace_text(f'{{app{app_no+1}_long_term_bus_text}}',long_term_bus_txt)
                    self._ppt.replace_text(f'{{app{app_no+1}_long_term_vio_cnt}}',long_term_vio_cnt)
                    self._ppt.replace_text(f'{{app{app_no+1}_long_term_vio_text}}',long_term_vio_txt)


            #replaceHighlight application specific data
            if self._generate_HL:
                lic_df=self._hl_data.get_lic_info(app_id)
                lic_df=self._hl_data.sort_lic_info(lic_df)
                oss_df=self._hl_data.get_cve_info(app_id)
                lic_summary = pd.DataFrame(columns=['License Type','Risk Factor','Component Count','Example'])

                crit_cve = self._hl_data.get_cve_crit_tot(app_id)
                high_cve = self._hl_data.get_cve_high_tot(app_id)
                med_cve = self._hl_data.get_cve_med_tot(app_id)
                oss_cmpnt_tot = self._hl_data.get_oss_cmpn_tot(app_id)

                self._ppt.replace_text(f'{{app{app_no+1}_crit_sec_tot}}',crit_cve)
                self._ppt.replace_text(f'{{app{app_no+1}_high_sec_tot}}',high_cve)
                self._ppt.replace_text(f'{{app{app_no+1}_med_sec_tot}}',med_cve)

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

                crit_cve_days = math.ceil(crit_cve/5)
                crit_cve_cost = crit_cve_days * ap._day_rate /1000

                high_cve_days = math.ceil(high_cve/5)
                high_cve_cost = high_cve_days * ap._day_rate /1000

                med_cve_days = math.ceil(med_cve/5)
                med_cve_cost = med_cve_days * ap._day_rate /1000

                summary_days = int(summary_days) + int(crit_cve_days)
                summary_cost = round(summary_cost + crit_cve_cost,2)

                fix_now_days = int(fix_now_days) + int(crit_cve_days)
                fix_now_cost = round(fix_now_cost + crit_cve_cost,2)

                near_term_days = int(near_term_days) + int(high_cve_days)
                near_term_cost = round(near_term_cost + high_cve_cost + med_cve_cost,2)

                self._ppt.replace_text(f'{{app{app_no+1}_high_sec_tot}}','{high_cve}')
                self._ppt.replace_text(f'{{app{app_no+1}_med_sec_tot}}','{med_cve}')

            #both AIP and HL data
            summary_total_cost = summary_total_cost + summary_cost

            self._ppt.replace_text(f'{{app{app_no+1}_summary_days}}',summary_days)
            self._ppt.replace_text(f'{{app{app_no+1}_summary_cost}}',summary_cost)

            self._ppt.replace_text(f'{{app{app_no+1}_fix_now_days}}',fix_now_days)
            self._ppt.replace_text(f'{{app{app_no+1}_fix_now_cost}}',f'${fix_now_cost}K')

            self._ppt.replace_text(f'{{app{app_no+1}_near_term_days}}',near_term_days)
            self._ppt.replace_text(f'{{app{app_no+1}_near_term_cost}}',f'${near_term_cost}K')

        self._ppt.replace_text(f'{summary_total_cost}',summary_total_cost)
            
               

                


    # def get_hl_data(self):
    #     # self._hl_data = HLData(self._hl_base_url, self._hl_user, self._hl_pswd, self._hl_instance,self._project_name, self._hl_app_list)

    #     # Retreive the app ids for given instance.
    #     # TODO: try-except
    #     self._hl_apps_df = self._hl_data.get_app_ids(self._hl_instance)

    #     # App counter, used to address the next blank HL table on a slide.
    #     app_no = 0
    #     for i in range(len(self._hl_apps_df)):
    #         app_name = self._hl_apps_df[i]['name']

    #         if app_name in self._hl_app_list:
    #             print(f'HL - Processing {app_name}')
    #             app_id = self._hl_apps_df[i]['id']
    #             app_no = app_no + 1
    #             # Retreive CVE info

    #             temp_cve_df = pd.DataFrame()

    #             # Always retrieve all of the CRITICAL sev CVEs.
    #             try:
    #                 self._crit_cves_df = pd.DataFrame() 
    #                 self._high_cves_df = pd.DataFrame() 
    #                 self._med_cves_df = pd.DataFrame() 

    #                 self._crit_cves_df = self._hl_data.get_cves(app_id, 'critical')

    #                 # Only retrieve the HIGH sev CVEs, if we do not have enough of the CRITICAL sevs and,
    #                 # only retrieve enough to fill one slide.
    #                 # TODO: Do this only if third-party data was found.
    #                 if len(self._crit_cves_df) < MAX_CVES_PER_SLIDE:
    #                     self._high_cves_df = self._hl_data.get_cves(app_id, 'high', (MAX_CVES_PER_SLIDE - len(self._crit_cves_df)))

    #                 # Only retrieve the MEDIUM sev CVEs when we do not have enough of the CRITICAL and the HIGH sevs and,
    #                 # only retrieve enough to fill one slide.
    #                 if len(self._crit_cves_df) + len(self._high_cves_df) < MAX_CVES_PER_SLIDE:
    #                     self._med_cves_df = self._hl_data.get_cves(app_id, 'medium', 
    #                                     (MAX_CVES_PER_SLIDE - (len(self._crit_cves_df) + len(self._high_cves_df))))

    #                 temp_cve_df = self._crit_cves_df
    #                 temp_cve_df = temp_cve_df.append(self._high_cves_df)
    #                 temp_cve_df = temp_cve_df.append(self._med_cves_df)
    #             except:
    #                 print('ERROR: No CVE info found')

    #             # Retreive the risky license info

    #             try:
    #                 self._high_lic_df = self._hl_data.get_lics(app_id, 'high')

    #                 temp_lic_df = self._high_lic_df

    #                 if len(self._high_lic_df) < MAX_LICS_PER_SLIDE:
    #                     # Retrieve medium risk licenses only if we do no have enough highs.
    #                     # TODO: errorhandling
    #                     self._med_lic_df = self._hl_data.get_lics(app_id, 'medium', (MAX_LICS_PER_SLIDE - len(self._high_lic_df)))

    #                     # Always print all components wth high risk licneses and print mediums only 
    #                     # when there are not enough highs.

    #                     if len(self._med_lic_df) > 0:
    #                         temp_lic_df = temp_lic_df.append(self._med_lic_df)
    #             except:
    #                 print('ERROR: No license info found')

    #             if len(temp_lic_df) > 0:
    #                 self._ppt.update_table(f'app{app_no}_HL_table_lic_risks', temp_lic_df, include_index=False)
    #             if len(temp_cve_df) > 0:
    #                 self._ppt.update_table(f'app{app_no}_HL_table_CVEs', temp_cve_df, include_index=False)
    #         else:
    #             print(f'App {app_name}, found in HL, but report not requested. Skipping')

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


