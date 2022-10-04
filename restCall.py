from io import RawIOBase

from pandas import DataFrame
from pandas import json_normalize
from pandas import concat
from pandas import notnull
from pptx.chart.data import Category
from IPython.display import display

from requests.auth import HTTPBasicAuth 
from time import perf_counter, ctime
from copy import copy
from logger import Logger
from logging import DEBUG, INFO, ERROR, warning

from restAPI import RestCall
from aipRestCall import AipRestCall
from hlRestCall import HLRestCall

import requests
import enum
import urllib.parse
import numpy as np 
import json

"""
    This class is used to retrieve information from the CAST AIP REST API
"""
class AipData(AipRestCall):
    _data={}
    _base=[]

    _sizing = {
       '10151':'Number of Code Lines', 
       '10107':'Number of Comment Lines', 
       '10109':'Number of Commented-out Code Lines',
       '67011':'Critical Violations'
    }

    _tech_sizing = {
        '10151':'LoC',
        '10154':'Files',
        '10155':'Classes',
        '10158':'SQL Artifacts',
        '10163':'Tables'
    }

    _imp_list = ['66070','61001','61013','61004','66009',
                 '61009','61029','61031','61010','61011',
                 '61026','61027','66068','61018','61019',
                 '61014','61015','61020','61003','61024',
                 '66069','66063','66066','66062','66065',
                 '66064']

    _doc_list = ['61008','61007','61017','61017','61006','61028','61022','61023']

    _health_grade_ids = ['Efficiency','Robustness','Security','Changeability','Transferability']

    def __init__(self, config, timer_on=False,log_level=INFO):
        super().__init__(config.aip_url, config.aip_user, config.aip_password, timer_on,log_level)

        #self._rest=rest
        #self._base=app_list

        for s in config.aip_list:
            self.info(f'*******Collecting AIP data for {s}********')
            self._data[s]={}
            self._data[s]['has data'] = False
            central_schema = f'{s}_central'
            domain_id = self.get_domain(central_schema)
            if domain_id == -1:
                raise SystemExit  #connection failed, exit here
            if domain_id is not None:
                self._data[s]['domain_id']=domain_id
                snapshot = self.get_latest_snapshot(domain_id)
                self._data[s]['snapshot']=snapshot
                if self._data[s]['snapshot']:
                    self.info('AIP is active')
                    self._data[s]['has data'] = True

                    self.info('TQI compliance data')
                    self._data[s]['tqi_compliance']=self.aggregate_violation_ratio(domain_id,snapshot['id'],'60017',self._imp_list)

                    self.info('documentation compliance data')
                    self._data[s]['doc_compliance']=self.aggregate_violation_ratio(domain_id,snapshot['id'],'66033',self._doc_list,False)

                    self.info('Grades by technololgy')
                    self._data[s]['grades']=self.get_grades_by_technology(domain_id,self._data[s]['snapshot'])

                    self.info('Sizing by technical')
                    self._data[s]['sizing']=self.get_sizing_by_technology(domain_id,self._data[s]['snapshot'],self._sizing)

                    self.info('LOC data')
                    self._data[s]['loc_sizing']=self.get_sizing(domain_id,self._sizing) 

                    self.info('Technical sizing data')
                    self._data[s]['tech_sizing']=self.get_sizing(domain_id,self._tech_sizing) 

                    self.info('Violation sizing data')
                    self._data[s]['violation_sizing']=self.get_violation_CR(domain_id)

                    self.info('Critical rules')
                    self._data[s]['critical_rules']=self.get_rules(domain_id,self._data[s]['snapshot']['id'],60017,non_critical=False)

                    self.info('ISO rules')
                    self._data[s]['ISO']=self.__get_iso_rules(domain_id,snapshot['id'])

                    self.info('action plan data')
                    (ap_df,ap_summary_df) = self.get_action_plan(domain_id,self._data[s]['snapshot']['id']) 
                    self._data[s]['action_plan']=ap_df
                    self._data[s]['action_plan_summary']=ap_summary_df
                    self.info('AIP data retrieval complete')
            else:
                self.logger.warn(f'Domain not found for {s}')

    def has_data(self, app):
        return self._data[app]['has data']

    def data(self,app):
        return self._data[app]

    def iso_rules(self, app):
        return self._data[app]['ISO']

    def __get_iso_rules(self,domain_id,snapshot_id):
        iso={1061004:"Security",
             1061003:"Reliabllity",
             1061002:"Performance-Efficiency",
             1061001:"Maintainability",
        }

        rslt_df = DataFrame()
        rslt_df.style.set_properties(subset=['text'],**{'text-align': 'left'})
        for key, value in iso.items():
            try:
                temp = DataFrame(columns=['category','violation'])
                rp = json_normalize(self.get_rules(domain_id,snapshot_id,key)['rulePattern'])
                temp['violation'] = rp['name']
                temp['category'] = value
                temp = temp.groupby(['category','violation']).size().reset_index(name='count') 

                total = temp.groupby(['category'])['count'].sum().reset_index(name='count') 
                total['violation']=''
                total = total[['category','violation','count']]

                rslt_df = concat([rslt_df,total,temp])
            except KeyError as e:
                self.warning(f'no iso rules for {value} ({e})')

        if 'category' in rslt_df.columns:
            rslt_df['category'] = rslt_df['category'].mask(rslt_df['category'].ne(rslt_df['category'].shift()).cumsum().duplicated(), '')
        else:
            self.warning('No ISO rules found')
            
        return rslt_df

    def domain(self, app):
        return self.data(app)['domain_id']

    def snapshot(self, app):
        return self.data(app)['snapshot']

    def grades(self, app):
        return self.data(app)['grades']

    def sizing(self, app):
        return self.data(app)['sizing']

    def critical_rules(self, app):
        return self.data(app)['critical_rules']

    def tqi_compliance(self, app):
        return self.data(app)['tqi_compliance']

    def doc_compliance(self, app):
        return self.data(app)['doc_compliance']

    def aggregate_violation_ratio(self,domain_id, snapshot_id,key,sub_keys,crit_only=True):
        self.debug(f'aggregating violation ration information for {domain_id}{snapshot_id}')
        # sid = self.snapshot(app)['id']
        # domain_id = self.domain(app)
        
        df = DataFrame(columns=['Key','Rule','Detail','Weight','Total','Failed','Succeeded','Compliance','Score'])
        indicators = self.get_quality_indicators(domain_id,snapshot_id, key)

        for ind in indicators:
            ind_key = ind['key']
            if ind_key in sub_keys:
                ind_name = ind['name'] #.replace('-','\n  ',1)
                ind_weight = ind['weight']
                vr = self.get_violation_ratio(domain_id,ind_key)
                try:
                    if vr:
                        temp = json_normalize(vr)
                        if crit_only:
                            temp = temp[temp['reference.critical']==True]
                        agr = temp.agg({ 
                            'result.violationRatio.totalChecks': ['sum'], 
                            'result.violationRatio.failedChecks': ['sum'], 
                            'result.violationRatio.successfulChecks': ['sum'] 
                            })
        #                agr['compliance']=(agr['result.violationRatio.successfulChecks']/agr['result.violationRatio.totalChecks'])*100
                        grade_set=self.get_grade(domain_id,ind_key)
                        grade=round(grade_set[0]['result']['grade'],2)

                        compliance=grade_set[0]['result']['score']*100

                        # temp = temp[temp['result.grade']<=2]
                        # Get Top 3 rules, based on max weight of Failed check logic  
                        temp = temp.sort_values('result.violationRatio.failedChecks', ascending = False).head(3)
                        
                        detail = "\n".join(temp['reference.name'].tolist())
                        total=int(agr.iloc[0]['result.violationRatio.totalChecks'])
                        if total == 0:
                            continue
                        failed=int(agr.iloc[0]['result.violationRatio.failedChecks'])
                        success=int(agr.iloc[0]['result.violationRatio.successfulChecks'])

                        df.loc[len(df)] = [
                            ind_key,
                            ind_name,
                            detail,
                            ind_weight,
                            f'{total:,}',
                            f'{failed:,}',
                            f'{success:,}',
                            f'{int(compliance)}%',
                            grade
                        ]                
                except (KeyError):
                    self.logger.warn(f'Key Error while aggregating {ind_name}')

        # for ind in indicators:
        #     key = ind['key']
        #     name = ind['name']
        #     weight = ind['weight']
        #     vr = self.get_violation_ratio(domain_id,key)
        #     if vr:
        #         temp = json_normalize(vr)
        #         temp['key']=key
        #         temp['name']=name
        #         temp['weight']=weight
        #         df = concat([df,temp])

        # df.rename(columns={"name":"parent","reference.name":"name","reference.critical":"critical"},inplace=True)
        # df.rename(columns={"result.grade":"grade","result.violationRatio.totalChecks":"totalChecks"},inplace=True)
        # df.rename(columns={"result.violationRatio.failedChecks":"failedChecks"},inplace=True)
        # df.rename(columns={"result.violationRatio.successfulChecks":"successfulChecks"},inplace=True)
        # df.rename(columns={"result.violationRatio.violationOccurrences":"violationOccurrences"},inplace=True)
        # df.rename(columns={"result.violationRatio.ratio":"ratio"},inplace=True)

        # return df[[ 'type','key', 'parent', 'weight','name','critical', 'grade','totalChecks', 'failedChecks', \
        #             'violationOccurrences', 'successfulChecks','ratio', 'result.score']]

        return df

    def action_plan(self, app):
        ap_df = self.data(app)['action_plan']
        ap_summary_df = self.data(app)['action_plan_summary']
        return (ap_df,ap_summary_df)
        
    def get_app_grades(self, app, sort=False):
        app_grades = self.grades(app).loc['All']
        if sort:
            return app_grades
        else:
            return app_grades.sort_values()

    def calc_grades_all_apps(self):
        all_app=DataFrame()
        for row in self._data:
            if self.has_data(row):
                app_name=self.snapshot(row)['name']
                grades=self.grades(row)
                all_app = concat([all_app,grades[grades.index.isin(['All'])].rename(index={'All': app_name})]).drop_duplicates()
        return all_app[all_app.columns].mean(axis=0)

    def calc_grades_health(self,grade_all):
        grade_df = DataFrame(grade_all)
        grade_health=grade_df[grade_df.index.isin(self._health_grade_ids)]
        return grade_health

    def calc_health_grades_high_risk(self,grade_all):
        grade_health = self.calc_grades_health(grade_all)
        grade_at_risk=grade_health[grade_health < 2].dropna()
        return grade_at_risk

    def calc_health_grades_medium_risk(self,grade_all):
        grade_health = self.calc_grades_health(grade_all)
        grade_at_risk=grade_health[grade_health > 2].dropna()
        grade_at_risk=grade_health[grade_health < 3].dropna()
        return grade_at_risk

    def get_loc_sizing(self,app):
        return self.data(app)['loc_sizing']

    def tech_sizing(self, app):
        return self.data(app)['tech_sizing']

    def violation_sizing(self, app):
        return self.data(app)['violation_sizing']

    def get_all_app_text(self):
        rslt = ""

        data = self._data
        l = len(self._base)
        if l == 1:
            return self.snapshot(self._base[0])['name']

        last_name = self._base[-2]
        for a in self._base:
            if self.has_data(a):
                rslt = rslt + self.snapshot(a)['name']
                if l >= 2 and a == last_name:
                    rslt = rslt + " and "
                elif a != self._base[-1]:
                    rslt = rslt + ", "
            else:
                rslt = "NO SNAPSHOT INFORMATION AVAILABLE"
        return rslt

    def get_grade_by_tech(self,app):
        grade_df = self.grades(app).round(2).applymap('{:,.2f}'.format)
        grade_df = self.grades(app).round(2)
        grade_df = grade_df[grade_df.index.isin(['All'])==False]

        sizing_df = DataFrame(self.sizing(app))
        sizing_df = sizing_df[sizing_df.index.isin(['All'])==False]
        sizing_df = DataFrame(sizing_df["Number of Code Lines"].rename("LOC")).dropna()
        sizing_df = sizing_df.applymap('{:,.0f}'.format)
        
        tech = sizing_df.join(grade_df) 

        sizing_df = DataFrame(self.sizing(app)) 
        sizing_df = sizing_df[sizing_df.index.isin(['All'])==False]
        sizing_df = DataFrame(sizing_df["Critical Violations"]).dropna()
        sizing_df = sizing_df.applymap('{:,.0f}'.format)

        tech = tech.join(sizing_df)
        tech.sort_values(by='LOC',ascending=False,inplace=True)
#        tech = tech.applymap('{:,.2f}'.format)

        
        return tech

    def get_high_risk_grade_text(self, grades):
        grade_at_risk=self.calc_health_grades_high_risk(grades)
        if grade_at_risk.empty:
            return None
        else:
            return self.text_from_list(grade_at_risk.index.values.tolist())

    def get_medium_risk_grade_text(self, grades):
        grade_at_risk=self.calc_health_grades_medium_risk(grades)
        if grade_at_risk.empty:
            return None
        else:
            return self.text_from_list(grade_at_risk.index.values.tolist())

    def text_from_list(self,list):
        rslt = ""
        total_items = len(list)
        if total_items == 1:
            return list[0]

        last_item = list[-1]
        sec_last_item = list[-2]
        for name in list:
            rslt = rslt + name
            if total_items >= 2 and name == sec_last_item:
                rslt = rslt + " and "
            elif name != last_item:
                rslt = rslt + ", "
        return rslt



"""

"""
class HLData(HLRestCall):
    _data={}

    def __init__(self, config, timer_on=False,log_level=INFO):
        super().__init__(config.hl_url, config.hl_user, config.hl_password, config.hl_instance, timer_on,log_level)

        for s in config.hl_list:
            self.info(f'Collecting Highlight data for {s}')
            self._data[s]={}
            self._data[s]['has data'] = False

            app_id = self.get_app_id(s)
            if app_id:
                (lic,cves,total_components) = self.get_third_party(app_id)
                
                self._data[s]['has data'] = True
                self._data[s]['app_id']=app_id
                self._data[s]['cves']=cves
                self._data[s]['licenses']=lic
                self._data[s]['cloud']=self.get_cloud_data(app_id)

                if cves.empty:
                    self._data[s]['cve_crit_tot']=0
                    self._data[s]['cve_crit_comp_tot']=0
                    self._data[s]['cve_high_tot']=0
                    self._data[s]['cve_high_comp_tot']=0
                    self._data[s]['cve_med_tot']=0
                    self._data[s]['cve_med_comp_tot']=0
                else:
                    self._data[s]['cve_crit_tot']=len(cves[cves['criticity']=='CRITICAL']['cve'].unique())
                    self._data[s]['cve_crit_comp_tot']=len(cves[cves['criticity']=='CRITICAL']['component'].unique())
                    self._data[s]['cve_high_tot']=len(cves[cves['criticity']=='HIGH']['cve'].unique())
                    self._data[s]['cve_high_comp_tot']=len(cves[cves['criticity']=='HIGH']['component'].unique())
                    self._data[s]['cve_med_tot']=len(cves[cves['criticity']=='MEDIUM']['cve'].unique())
                    self._data[s]['cve_med_comp_tot']=len(cves[cves['criticity']=='MEDIUM']['component'].unique())
                if lic.empty:
                    self._data[s]['lic_high_tot']=0
                    self._data[s]['lic_med_tot']=0
                else:
                    # make some adjustment since HL data has changed
                    lic['compliance']=lic['compliance'].str.replace('compliant','high')
                    lic['compliance']=lic['compliance'].str.replace('partial','medium')
                    lic['compliance']=lic['compliance'].str.replace('notCompliant','low')

                    self._data[s]['lic_high_tot']=len(lic[lic['compliance']=='low']['component'].unique())
                    self._data[s]['lic_med_tot']=len(lic[lic['compliance']=='medium']['component'].unique())

                self._data[s]['total_components']=total_components
            else:
                self.error(f'Highlight Application Id not found for {s}')    
        #    self._base = schema
            self._data_retrieved = False
            self._app_id = None
            self._got_data = False
            self._has_crit_sev_cves = False
            self._has_high_sev_cves = False
            self._has_med_sev_cves = False
            self._has_high_risk_lics = False
            self._has_med_risk_lics = False
            self._third_party_df = DataFrame()
            self._cve_df = DataFrame()
            self._lic_df = DataFrame()

    def get_cve_data(self,app_id):
        return self._data[app_id]['cves']

    def has_data(self,app_id):
        return self._data[app_id]['has data'] 

    def get_data(self,app_id):
        return self._data[app_id]

    def get_cloud_info(self,app_id):
        return self._data[app_id]['cloud'] 

    def get_cve_crit_tot(self,app_id):
        return self._data[app_id]['cve_crit_tot']
    def get_cve_crit_comp_tot(self,app_id):
        return self._data[app_id]['cve_crit_comp_tot']

    def get_cve_high_tot(self,app_id):
        return self._data[app_id]['cve_high_tot']
    def get_cve_high_comp_tot(self,app_id):
        return self._data[app_id]['cve_high_comp_tot']

    def get_cve_med_tot(self,app_id):
        return self._data[app_id]['cve_med_tot']
    def get_cve_med_comp_tot(self,app_id):
        return self._data[app_id]['cve_med_comp_tot']

    def get_lic_high_tot(self,app_id):
        return self._data[app_id]['lic_high_tot']

    def get_lic_med_tot(self,app_id):
        return self._data[app_id]['lic_med_tot']

    def get_oss_cmpn_tot(self,app_id):
        return self._data[app_id]['total_components']
        

    def get_third_party_info(self,app_id):
        return DataFrame(self._data[app_id]['components'])

    def get_lic_info(self,app_name):
        """
            Extract all license relavent columns from the components DF 
        """
        lic = self._data[app_name]['licenses']

        #adjust license risk factors
        if lic is None:
            lic = DataFrame()
        else:
            lic.loc[lic['compliance']=='high','risk']='Low'
            lic.loc[lic['compliance']=='medium','risk']='Medium'
            lic.loc[lic['compliance']=='low','risk']='High'
            lic.loc[lic['compliance']=='undefined','risk']='Undefined'
       
        return lic

    def get_cve_info(self,app_name):
        """
            Extract all license relavent columns from the components DF 
        """
        cves = self._data[app_name]['cves']
        oss_df = DataFrame(columns=['component','critical','high','medium'])

        if cves is not None and not cves.empty:

            oss_work_df = cves[['component','cve','criticity']].copy()
            oss_work_df.drop_duplicates(['component','cve','criticity'],inplace=True)
            oss_work_df.sort_values(by=['component'],inplace=True)

            cve_critical_df = oss_work_df[oss_work_df['criticity']=='CRITICAL']
            cve_critical_df = cve_critical_df.groupby('component')['cve'].apply(', '.join).reset_index()
            cve_critical_df.rename(columns={'cve':'critical'},inplace=True)

            cve_high_df = oss_work_df[oss_work_df['criticity']=='HIGH']
            cve_high_df = cve_high_df.groupby('component')['cve'].apply(', '.join).reset_index()
            cve_high_df.rename(columns={'cve':'high'},inplace=True)

            cve_medium_df = oss_work_df[oss_work_df['criticity']=='MEDIUM']
            cve_medium_df = cve_medium_df.groupby('component')['cve'].apply(', '.join).reset_index()
            cve_medium_df.rename(columns={'cve':'medium'},inplace=True)

            oss_df = cve_critical_df.merge(cve_high_df,on='component',how='outer')
    #        oss_df = oss_df.merge(cve_high_df,on='component',how='outer')
            oss_df = oss_df.merge(cve_medium_df,on='component',how='outer')
            oss_df = oss_df.where(notnull(oss_df),'')
            oss_df = oss_df[['component','critical','high','medium']]

        return oss_df



        # lic = pd.json_normalize(components,
        #                         record_path =['licenses'],
        #                         meta=['componentId','name','version','release',
        #                               'languages','lastVersion','lastRelease',
        #                               'nbVersionPreviousYear'],
        #                         meta_prefix='comp.',
        #                         errors='ignore')

        # lic=lic[['comp.componentId', 'comp.name', 'comp.version',
        #     'comp.release', 'comp.languages', 'comp.lastVersion',
        #     'comp.lastRelease', 'comp.nbVersionPreviousYear','name', 'compliance']]

        # lic.rename(columns={'name':'lic.name','compliance':'compliance'},inplace=True)

        # #adjust license risk factors
        # lic.loc[lic['compliance']=='high','compliance']='Low'
        # lic.loc[lic['compliance']=='medium','compliance']='Medium'
        
        # return lic

    def sort_lic_info(self, lic_df):
        if lic_df is None:
            return lic_df
        else:
            lic_all = lic_df
            try:
                lic_all.sort_values(by=['component','release'],inplace=True)
                lic_high = lic_all[lic_all['compliance']=="High"]
                lic_medium = lic_all[lic_all['compliance']=="Medium"]
                return concat([lic_high,lic_medium])
            except (KeyError):
                return lic_df


    def get_app_ids(self, instance_id):
        # TODO: try-except
        return self._get_app_ids(instance_id)

    def _get_third_party(self, app_id):

        if self._app_id != app_id:
            self._app_id = app_id
            self._got_data = False

        if not self._got_data:
            # If we do not have the data already for this app, retrieve it first.

            self._third_party_df = DataFrame(self.get_third_party(app_id)['thirdParties'])
            # TODO: Confirm that data was retrieved before setting this to True
            self._got_data = True

            self._cve_df = self._third_party_df.loc[:, ['name', 'cve']]
            self._cve_df.dropna(axis = 0, how = 'any', inplace = True)
            self._lic_df = self._third_party_df.loc[:, ['name', 'licenses']]
            self._lic_df.dropna(axis = 0, how = 'all', inplace = True)

        return

    def get_cves(self, components, type, limit = 0):
        """
        Returns CVE info for a give app, if there are Critical, High and medium CVEs.
        Unless the 'all' argument is provided, High CVEs are returned only when there are no/not enough critical CVEs
        to fill the slide. Similarly, medium CVEs are returned only where there are no/limited critical or high CVEs.

        Note that low CVEs are ignored for assessment purposes.
        """

        _cve_df = DataFrame()
        sev_type = type.lower()

        # If the request is for critical sev CVEs, override the limit and return all rows.
        if sev_type == 'critical':
            limit = 0

        # Do we have the data retrieved for the app? If not, auto-retrieve.
        # If all == False, limit the number of CVEs returned.

        # try:
        #     self._get_third_party(app_id)
        # except:
        #     print('ERROR - no thirdparty data')
        #     raise

        cve_df = components.loc[:, ['name', 'cve']]
        cve_df.dropna(axis = 0, how = 'any', inplace = True)


        i = 0
        prev_comp = ''
        comp_changed = False

        for tup in self._cve_df.itertuples():
            crit_cve_str, high_cve_str, med_cve_str = '', '', ''
            comp_name = tup[1]

            for cve in tup[2]['vulnerabilities']:
                # If critical sev is the type requested, also grab high and medium sev CVEs for the same component.
                # Similarly, if high sev is the type requested, also grab medium sev CVEs for the same component.
                criticality = cve['criticity'].lower()

                # Ignore components with low criticality, as we do not need to report those.
                if criticality == 'low':
                    continue

                # Note that are not considering the type requested here, but simply storing the CVE in its
                # appropriate slot. That part is handled further below, before they are added to the dataframe.
                if criticality == 'critical':
                    crit_cve_str += cve['name'] + ', '
                if criticality == 'high':
                    high_cve_str += cve['name'] + ', '
                elif criticality == 'medium':
                    med_cve_str += cve['name'] + ', '

            # Before adding the results into the dataframe, ensure that we have matches for the
            # crticality requested. For example, if critical CVEs were requested, we should have
            # values in the crit_cve_str. # If not, do not add.
            if sev_type == 'critical' and len(crit_cve_str) > 0 or \
                sev_type == 'high' and len(high_cve_str) > 0    or \
                sev_type == 'medium' and len(med_cve_str) > 0:

                # Strip the comma-spaces at the end of the CVE strings.
                crit_cve_str = crit_cve_str[:-2]
                high_cve_str = high_cve_str[:-2]
                med_cve_str = med_cve_str[:-2]

                if len(crit_cve_str) == 0:
                    crit_cve_str = 'N/A'
                if len(high_cve_str) == 0:
                    high_cve_str = 'N/A'
                if len(med_cve_str) == 0:
                    med_cve_str = 'N/A'

                _cve_df = _cve_df.append({'Component Name': comp_name, 'Critical Sev CVEs': crit_cve_str,
                            'High Sev CVEs': high_cve_str, 'Medium Sev CVEs': med_cve_str}, ignore_index = True)

            i += 1

            # Return only the max number of rows requested. In case the sev type is critical, always return all rows.
            if (limit > 0) and (i == limit):
                break

        #print(_cve_df)
        return _cve_df

    def get_lics(self, app_id, type = 'high', limit = 0):
        """
        """

        compliance = ''
        _lic = []
        _lic_df = DataFrame()

        # Do we have the data retrieved for the app? If not, auto-retrieve.
        # If all == False, limit the number of CVEs returned.

        try:
            self._get_third_party(app_id)
        except:
            print('ERROR - no thirdparty data')
            raise

        i = 0
        for tup in self._lic_df.itertuples():
            high_risk_lic_str = ''
            med_risk_lic_str = ''
            comp_name = tup[1]

            for dic in tup[2]:
                if type == 'high':
                    # A note on risk-complaince relation.
                    # A riskier license translates to lower compliance. The REST API returns compliance info.
                    # So, if the type being sought is high-risk license, we return low compliance license.
                    #
                    # If tbe requested type is high risk licenses and we also see
                    # medium risk licenses for the same component, include those as well.
                    # Leveraging the content of high_risk_lic_str for this purpose.

                    if dic['compliance'] == 'low':
                        high_risk_lic_str += dic['name'] + '\n'
                    elif dic['compliance'] == 'medium' and len(high_risk_lic_str) > 0:
                        med_risk_lic_str += dic['name'] + '\n'
                elif type == 'medium':
                    # A medium risk likcense has a compliance of medium.
                    if dic['compliance'] == 'medium': 
                        med_risk_lic_str += dic['name'] + '\n'
            
            if len(high_risk_lic_str) > 0 or len(med_risk_lic_str) > 0:
                # Strip the trailing newline.
                high_risk_lic_str = high_risk_lic_str[:-1]
                med_risk_lic_str = med_risk_lic_str[:-1]

                if len(high_risk_lic_str) == 0:
                    high_risk_lic_str = 'N/A'

                if len(med_risk_lic_str) == 0:
                    med_risk_lic_str = 'N/A'

                _lic_df = _lic_df.append({'Component Name': comp_name, 'High Risk': high_risk_lic_str,
                            'Medium Risk': med_risk_lic_str}, ignore_index = True)
                i += 1

            if (limit > 0) and (i == limit):
                break

        """
        for elem in _lics:
            print(elem.key, elem.value)
        for idx in range(len(_lics)):
            for key in _lics[idx]:
                print(key, _lics[idx][key])

            #for j in range(len(high_risk_lic_df['LicenseInfo'][i])):
                #print(high_risk_lic_df['LicenseInfo'][i][j])
                #_lic_str += high_risk_lic_df['name'][i][j] + '\n'

            #_lic_str = _lic_str[:-1]
            #_lics = _lics.append({'name:', high_risk_lic_df['licenses'], 'licenses:', _lic_str})
        """
        #print(_lic_df)
        return _lic_df

# rest = AipData('http://sha-dd-node02:8080/CAST-RESTAPI-integrated/rest/','cast','cast',['MainWebApplication'],log_level=DEBUG)
# vr = rest.aggregate_violation_ratio('MainWebApplication',"60017")
