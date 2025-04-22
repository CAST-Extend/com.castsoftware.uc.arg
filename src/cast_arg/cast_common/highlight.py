from cast_common.restAPI import RestCall
from cast_common.logger import Logger, INFO,DEBUG

from requests import codes
from pandas import ExcelWriter,DataFrame,json_normalize,concat
from pptx.dml.color import RGBColor
from typing import List
from json import loads,dumps

class Highlight(RestCall):

    _data = {}
    _third_party = {}
    _cves = {}
    _license = {}

    _tags = []
    _cloud = []
    _oss = []
    _elegance = []
    _apps = None
    _apps_full_list = None
    _instance_id = 0
    _base_url = None
    _hl_basic_auth = None
    _hl_user = None
    _hl_pswd = None

    log = None

    grades = {
        'softwareHealth':{'threshold':[53,75]},
        'softwareResiliency':{'threshold':[65,87]},
        'softwareAgility':{'threshold':[54,69]},
        'softwareElegance':{'threshold':[39,70]},
        'cloudReady':{'threshold':[50,75]},
        'openSourceSafety':{'threshold':[50,75]},
        'greenIndex':{'threshold':[50,75]}
    }

    def __init__(self,  
                 hl_user:str=None, hl_pswd:str=None,hl_basic_auth=None, hl_instance:int=0,
                 hl_apps:str=[],hl_tags:str=[], 
                 hl_base_url:str=None, 
                 log_level=INFO, timer_on=False):

        # general message to be used by ValueError exception
        msg='must be supplied with the first Highlight class instance'

        if Highlight.log is None:
            Highlight._log_level = log_level
            Highlight.log = Logger('Highlight',log_level)

        reset_data = False
        if Highlight._instance_id == 0:
            if hl_instance == 0:
                raise ValueError(f'Domain id {msg}') 

            Highlight._instance_id = hl_instance
            Highlight._data = {}
            reset_data = True
        
        #set the base url, if not already set from a prev instance of the Highlight class
        if Highlight._base_url is None:
            if hl_base_url is None:
                raise ValueError(f'Base url {msg}') 

            Highlight._base_url = hl_base_url
            if Highlight._base_url.endswith('/'):
                Highlight._base_url=Highlight._base_url[:-1]
                            
            if not Highlight._base_url.endswith('/WS2'):
                Highlight._base_url=f'{Highlight._base_url}/WS2/'

        #set the user name/password or Basic Autorization 
        if Highlight._hl_basic_auth is None and \
            Highlight._hl_user is None and \
            Highlight._hl_pswd is None:

            if hl_user is None and \
                hl_pswd is None and \
                hl_basic_auth is None:
                raise ValueError(f'UserName/password or Basic Authorization {msg}')

            if not hl_user is None:
                Highlight._hl_user=hl_user

            if not hl_pswd is None:
                Highlight._hl_pswd=hl_pswd

            if not hl_basic_auth is None:
                Highlight._hl_basic_auth=hl_basic_auth

        super().__init__(base_url=Highlight._base_url, user=Highlight._hl_user, password=Highlight._hl_pswd, 
                         basic_auth=Highlight._hl_basic_auth,track_time=timer_on,log_level=Highlight._log_level)

        if reset_data:

            Highlight._benchmark = DataFrame(self._get(r'/benchmark'))

            Highlight._tags = self._get_tags()
            Highlight._apps_full_list = []

            # retrieve all applications from HL REST API
            Highlight._apps_full_list = DataFrame(self._get_applications())

            # if the tag parameter is passed then filter applictions by the tag
            if type(hl_tags) is not list:
                hl_tags = [hl_tags]
            hl_apps = []
            stop = False
            for tag in hl_tags:
                for idx,app in Highlight._apps_full_list.iterrows():
                    id = app['id']
                    rslt,json = self.get(f'domains/{Highlight._instance_id}/applications/{id}/tags')
                    if rslt == codes.ALL_GOOD:
                        for t in json:
                            if t['label'] == tag:
                                for item in t['applications']:
                                    hl_apps.append(item['name'])
                                stop = True
                                break
                    if stop:
                        break
                    pass


            self.info(f'Found {len(Highlight._apps_full_list)} analyzed applications')

            # if the apps is not already of list type then convert it now
            if not isinstance(hl_apps,list):
                hl_apps=list(hl_apps.split(','))

            #if hl_apps is empty, include all applications prevously retrieved
            #otherwise filter the list down to include only the selected applications
            if len(hl_apps) > 0:
                Highlight._apps = Highlight._apps_full_list[Highlight._apps_full_list['name'].isin(hl_apps)]
            else:
                Highlight._apps = Highlight._apps_full_list

            if Highlight._apps is not None:
                self.info(f'{len(Highlight._apps)} applications selected: {",".join(Highlight._apps["name"])}')

        pass

    @property
    def app_list(self) -> DataFrame:
        return Highlight._apps

    def _get_application_data(self,app:str=None):
        self.debug('Retrieving all HL application specific data')
        
        if app is None:
            df = Highlight._apps
        else:
            df = Highlight._apps[Highlight._apps['name']==app]

        for app_name in df['name']:
            if app_name not in Highlight._data:
                try:
                    if not app_name in Highlight._data:
                        self.info(f'Loading Highlight data for {app_name}')
                        Highlight._data[app_name]=self._get_app_from_rest(app_name)
                except KeyError as ke:
                    self.warning(str(ke))
                    pass
                except Exception as ex:
                    self.error(str(ex))
        pass

    def _get_metrics(self,app_name:str) -> dict:
        try:
            if app_name not in Highlight._apps['name'].to_list():
                raise ValueError(f'{app_name} is not a selected application')

            self._get_application_data(app_name)
            data = Highlight._data[app_name]
            metrics = data['metrics'][0][0]
            return metrics.copy()
        except KeyError as ke:
            self.warning(f'{app_name} has no Metric Data')
            return {}

    def _get_third_party_data(self,app_name:str=None):
        self.debug('Retrieving third party HL data')

        if app_name not in Highlight._apps['name'].to_list():
            raise ValueError(f'{app_name} is not a selected application')

        if app_name is None:
            df = Highlight._apps
        else:
            df = Highlight._apps[Highlight._apps['name']==app_name]

        for idx,app in df.iterrows():
            app_name = app['name']
            if app_name not in Highlight._third_party:
                try:
                    if not app_name in Highlight._third_party:
                        self.info(f'Loading Highlight component data for {app_name}')
                        data = self._get_third_party_from_rest(app_name)
                        Highlight._third_party[app_name]=data
                except KeyError as ke:
                    self.warning(str(ke))
                    pass
                except Exception as ex:
                    self.error(str(ex))
        pass

    def _get_third_party(self,app_name:str) -> dict:
        try:
            if app_name not in Highlight._apps['name'].to_list():
                raise ValueError(f'{app_name} is not a selected application')

            self._get_third_party_data(app_name)
            data = Highlight._third_party[app_name]
#            third_party = data['metrics'][0]
            return data
        except KeyError as ke:
            self.error(f'{app_name} has no Metric Data')
            raise ke

    def _get_third_party_from_rest(self,app:str) -> dict:
        return self._get(f'domains/{Highlight._instance_id}/applications/{self.get_app_id(app)}/thirdparty')

    def _get(self,url:str,header=None) -> DataFrame:
        (status, json) = self.get(url,header)
        if status == codes.ok or codes.no_content:
#            return DataFrame(json)
            return json
        else:
            raise KeyError (f'Server returned a {status} while accessing {url}')

    def _get_applications(self):
        return DataFrame(self._get(f'domains/{Highlight._instance_id}/applications/', {'Accept': 'application/vnd.castsoftware.api.basic+json'} ))
        
    def _get_tags(self) -> DataFrame:
        return DataFrame(self._get(f'domains/{Highlight._instance_id}/tags/'))

    def _get_app_from_rest(self,app:str) -> DataFrame:
        ap_data = self._get(f'domains/{Highlight._instance_id}/applications/{self.get_app_id(app)}')
        df = DataFrame.from_dict(ap_data, orient='index')
        df = df.transpose()
        return df

    def get_app_id(self,app_name:str) -> int:
        """get the application id

        Args:
            app_name (str): application name 

        Raises:
            KeyError: application not found

        Returns:
            int: highlight application id
        """
        if len(Highlight._apps)==0:
            Highlight._apps = self._get_applications()

        series = Highlight._apps[Highlight._apps['name']==app_name]
        if series.empty:
            raise KeyError (f'Highlight application not found: {app_name}')
        else:
            return int(series.iloc[0]['id'])
            
    def get_tag_id(self,tag_name:str):
        """get the application id

        Args:
            tag_name (str): tag name 

        Raises:
            KeyError: tag not found

        Returns:
            int: highlight tag id
        """
        if len(Highlight._tags)==0:
            Highlight._tags = self._get_tags()

        series = Highlight._tags[Highlight._tags['label']==tag_name]
        if series.empty:
            raise KeyError (f'Highlight tag not found: {tag_name}')
        else:
            return int(series.iloc[0]['id'])
            
    def add_tag(self,app_name,tag_name) -> bool:
        """Add tag for application 

        Args:
            app_name (_type_): applicaton name
            tag_name (_type_): tag name
        Returns:
            bool: True if tag added
        """
        app_id = self.get_app_id(app_name)
        tag_id = self.get_tag_id(tag_name)

        url = f'domains/{Highlight._instance_id}/applications/{app_id}/tags/{tag_id}'
        (status,json) = self.post(url)
        if status == codes.ok or status == codes.no_content:
            return True
        else:
            return False 
    
    def get_tech_debt_advisor(self,order:List[str],tag:str=None) -> DataFrame:
        #https://rpa.casthighlight.com/WS2/domains/1271/technicalDebt/aggregated?tagIds=685&activePagination=false&pageOffset=0&maxEntryPerPage=9999&maxPage=999&order=healthFactor&order=alert
        tag_part = ""
        if tag is not None:
            tag_part = f'tagIds={self.get_tag_id(tag)}&'
        url = f'domains/{Highlight._instance_id}/technicalDebt/aggregated?{tag_part}&activePagination=false&pageOffset=0&maxEntryPerPage=9999&maxPage=999'
        for o in order:
            url=f'{url}&order={o}'
        
        (status, json) = self.get(url)
        if status == codes.ok:
            d1 = json_normalize(json,['subLevel'], max_level=1)
            d1.rename(columns={'id':'Health Factor'},inplace=True)
            d1.drop(columns=['levelType'], inplace=True)

            d2=d1.explode('subLevel')
            d2=json_normalize(d2.to_dict('records'), max_level=1)
            d2=d2.rename(columns={'subLevel.id':'Alert','subLevel.detail':'detail'})
            d2=d2.drop(columns=['subLevel.levelType'])

            d3=json_normalize(d2.to_dict('records'),['detail'],meta=['Health Factor','Alert'])
            d3=d3.rename(columns={'application_id':'App Id','effort':'Effort'})
            d3=d3.drop(columns=['counter'])

            return d3

        else:
            raise KeyError (f'Server returned a {status} while accessing {url}')

    """ **************************************************************************************************************
                                            Third Party Component Data 
    ************************************************************************************************************** """
    def get_component_total(self,app_name:str) -> int:
        third_party_data = self._get_third_party(app_name)
        if 'thirdParties' in third_party_data:
            return len(third_party_data['thirdParties'])
        else:
            return 0

        # third_party = len(self._get_third_party(app_name)['thirdParties'])
        # return int(third_party)

    def get_cve_data(self,app_name:str) -> DataFrame:
        data = self._get_third_party(app_name)
        if 'thirdParties' in data:
            third_party = data['thirdParties']
        else:
            third_party = []        
        if not app_name in Highlight._cves:
            cves = DataFrame()
            for tp in third_party:
                try:
                    if 'cve' in tp:
                        cve_df = json_normalize(tp['cve']['vulnerabilities'])
                        cve_df.rename(columns={'name':'cve'},inplace=True)
                        
                        cve_df['component']=self.in_collection(tp,'name',tp['name'])   
                        cve_df['version']=self.in_collection(tp,'version',tp['name']) 
                        cve_df['languages']=self.in_collection(tp,'languages',tp['name']) 
                        cve_df['release']=self.in_collection(tp,'release',tp['name'])
                        cve_df['origin']=self.in_collection(tp,'origin',tp['name']) 
                        cve_df['lastVersion']=self.in_collection(tp,'lastVersion',tp['name'])

                        cves=concat([cves,cve_df],ignore_index=True)
                except KeyError as e:
                    self.warning(f'Key Error while retrieving cve information: {e} {tp}')    
                    pass            
            if not cves.empty and 'component' in cves.columns:
                cves=cves[['component','version','languages','release','origin','lastVersion','cve', 'description', 'cweId', 'cweLabel', 'criticity', 'cpe']]

            Highlight._cves[app_name]=cves
        else:
            cves = Highlight._cves[app_name]
        return cves            

    def get_cve_critical(self, app_name:str) -> DataFrame:
        cves = self.get_cve_data(app_name)
        if not cves.empty:
            return cves[cves['criticity']=='CRITICAL']
        
    def get_cve_high(self, app_name:str) -> DataFrame:
        cves = self.get_cve_data(app_name)
        if not cves.empty:
            return cves[cves['criticity']=='HIGH']
    
    def get_cve_medium(self, app_name:str) -> DataFrame:
        cves = self.get_cve_data(app_name)
        if not cves.empty:
            return cves[cves['criticity']=='MEDIUM']

    def get_license_data(self,app_name:str) -> DataFrame:
        data = self._get_third_party(app_name)
        if 'thirdParties' in data:
            third_party = data['thirdParties']
        else:
            third_party = []
        
        if not app_name in Highlight._license:
            lic = DataFrame()
            for tp in third_party:
                if 'licenses' in tp:
                    lic_df = json_normalize(tp['licenses'], \
                        meta = ['name','version','languages','release','origin','lastVersion'])
                    lic_df.rename(columns={'name':'license'},inplace=True)
                    lic_df['component']=tp['name']
                    load_df_element(tp,lic_df,'version')
                    load_df_element(tp,lic_df,'languages')
                    load_df_element(tp,lic_df,'release')
                    load_df_element(tp,lic_df,'origin')
                    load_df_element(tp,lic_df,'lastVersion')
                    lic=concat([lic,lic_df],ignore_index=True)
            
            if not lic.empty and  'component' in lic.columns:
                lic=lic[['component','version','languages','release','origin','lastVersion','license','compliance']] 
                lic['compliance']=lic['compliance'].str.replace('compliant','high')
                lic['compliance']=lic['compliance'].str.replace('partial','medium')
                lic['compliance']=lic['compliance'].str.replace('notCompliant','low')

            Highlight._license[app_name]=lic
        else:
            lic = Highlight._license[app_name]
        return lic            

    def get_license_high(self,app_name:str) -> DataFrame:
        lic = self.get_license_data(app_name)
        if lic.empty:
            return DataFrame()
        return lic[lic['compliance']=='high'] 

    def get_license_medium(self,app_name:str) -> DataFrame:
        lic = self.get_license_data(app_name)
        if lic.empty:
            return DataFrame()
        return lic[lic['compliance']=='medium']   

    def get_license_low(self,app_name:str) -> DataFrame:
        lic = self.get_license_data(app_name)
        if lic.empty:
            return DataFrame()
        return lic[lic['compliance']=='low']   

    """ **************************************************************************************************************
                                            Green Detail Data 
    ************************************************************************************************************** """
    def get_green_detail(self,app_name:str)->DataFrame:
        """Highlight green impact data

        Args:
            app_name (str): name of the application

        Returns:
            DataFrame: flattened version of the Highlight green IT data
        """
        try:
            return json_normalize(self._get_metrics(app_name)['greenDetail'],'greenIndexDetails')
        except KeyError as ke:
            self.warning(f'{app_name} has no Green Impact Data')
            return None

    """ **************************************************************************************************************
                                            Cloud Ready Data 
    ************************************************************************************************************** """
    _cloud={}
    def get_cloud_detail(self,app_name:str)->DataFrame:
        if app_name in Highlight._cloud:
            return Highlight._cloud[app_name]
        

        columns = ['technology','cloudRequirement.display','cloudRequirement.ruleType','roadblocks','cloudEffort','cloudRequirement.criticality','files']

        data = self._get_metrics(app_name) 
        rslt = DataFrame()
        if 'cloudReadyDetail' in data:
            cloud_data = json_normalize(data['cloudReadyDetail'])
            for index,row in cloud_data.iterrows():
                tech = row['technology']
                db = json_normalize(row['cloudReadyDetails'])
                db['technology']=tech
                db = db[columns]

                rslt = concat([rslt,db],ignore_index=True)
                pass

        Highlight._cloud[app_name]=rslt
        return rslt
    
    _container={}
    def get_cloud_container(self,app_name:str)->DataFrame:
        if app_name in Highlight._container:
            return Highlight._container[app_name]

        columns = ['display','technology','impacts','criticality','roadblocks','cloudEffort']

        app_id = self.get_app_id(app_name)
        url = f'/domains/{self._instance_id}/applications/{app_id}/containerization'
        df = DataFrame(self._get(url))

        #is technology is misssing from rest call results
        if not 'technology' in df.columns:
            df['technology'] = ''

        #filter dataframe columns 
        df=df[columns]
        df['impacts'] = [','.join(map(str, l)) for l in df['impacts']]
        Highlight._container[app_name]=df

        return df

    """ **************************************************************************************************************
                                            General Metrics Data 
    ************************************************************************************************************** """
    def get_technology(self,app_name:str) -> DataFrame:
        """Highlight application technology data

        Args:
            app_name (str): _description_

        Returns:
            DataFrame: _description_
        """
        df = json_normalize(self._get_metrics(app_name))
        if not df.empty:
            tech = json_normalize(df['technologies'])
            tech = tech.transpose()
            tech = json_normalize(tech[0])
            return tech.sort_values(by=['totalLinesOfCode'],ascending=False)
        else:
            return DataFrame()

    def get_total_lines_of_code(self,app_name:str) -> int:
        return json_normalize(self._get_metrics(app_name))['totalLinesOfCode']    

    """ ******************************************************************************
                        Highlight Scores
    ****************************************************************************** """
    def calc_scores(self, app_list:list) -> dict:
        """ for each applicaiton in the Highligh portfolio retrieve and
            calculate the grade.  

        Args:
            app_list (list): list of application name in the portfilio

        Returns:
            dict: list of scores
        """
        t_scores= {}
        for key in self.grades:
            t_scores[key]=0

        for app in app_list:
            metrics = self._get_metrics(app)
            for key in self.grades:
                if key in metrics and metrics[key] is not None: 
                    t_scores[key]+=metrics[key]

        t_apps = len(app_list)
        for key in self.grades:
            t_scores[key]=round(t_scores[key]*100/t_apps,1)
            score = t_scores[key] 

        return t_scores

    def get_hml_color(self, hml:str):
        if hml == 'high':
            color = RGBColor(25,182,152)
        elif hml == 'medium':
            color = RGBColor(255,138,60)
        else: 
            color = RGBColor(234,97,83)
        return color

    def get_software_health_score(self,app_name:str) -> float:
        metrics = self._get_metrics(app_name)
        return float(metrics['softwareHealth'])*100

    def get_software_health_hml(self,app_name:str=None,score=None) -> str:
        if score is None:
            score = self.get_software_health_score(app_name)
        if score > 75:
            return 'high'
        elif score > 52:
            return 'medium'
        else:
            return 'low'

    def get_software_health_color(self,app_name:str=None,score=None) -> RGBColor:
        return self.get_hml_color(self.get_software_health_hml(app_name,score))



    # def get_software_agility_score(self,app_name:str) -> float:
    #     metrics = self._get_metrics(app_name)
    #     return float(metrics['softwareAgility'])*100

    # def get_software_agility_hml(self,app_name:str) -> str:
    #     score = self.get_software_agility_score(app_name)
    #     if score > 69:
    #         return 'high'
    #     elif score > 54:
    #         return 'medium'
    #     else:
    #         return 'low'

    # def get_software_elegance_score(self,app_name:str) -> float:
    #     metrics = self._get_metrics(app_name)
    #     return float(metrics['softwareElegance'])*100

    # def get_software_elegance_hml(self,app_name:str) -> str:
    #     score = self.get_software_elegance_score(app_name)
    #     if score > 70:
    #         return 'high'
    #     elif score > 39:
    #         return 'medium'
    #     else:
    #         return 'low'

    # def get_software_resiliency_score(self,app_name:str) -> float:
    #     metrics = self._get_metrics(app_name)
    #     return float(metrics['softwareResiliency'])*100
    
    # def get_software_resiliency_hml(self,app_name:str) -> str:
    #     score = self.get_software_resiliency_score(app_name)
    #     if score > 87:
    #         return 'high'
    #     elif score > 65:
    #         return 'medium'
    #     else:
    #         return 'low'

    # def get_software_cloud_score(self,app_name:str) -> float:
    #     metrics = self._get_metrics(app_name)
    #     return float(metrics['cloudReady'])*100
    # def get_get_cloud_hml(self,app_name:str=None,score=None) -> str:
    #     if score is None:
    #         score = self.get_software_cloud_score(app_name)
    #     if score > 75:
    #         return 'high'
    #     elif score > 52:
    #         return 'medium'
    #     else:
    #         return 'low'

    # def get_software_oss_safty_score(self,app_name:str) -> float:
    #     metrics = self._get_metrics(app_name)
    #     return float(metrics['openSourceSafety'])*100

    
    def get_get_software_oss_risk(self,app_name:str=None,score=None) -> str:
        if score is None:
            score = self.get_software_oss_safty_score(app_name)
        if score > 75:
            return 'low'
        elif score > 52:
            return 'medium'
        else:
            return 'high'

    def get_software_oss_license_score(self,app_name:str) -> float:
        metrics = self._get_metrics(app_name)
        return float(metrics['openSourceLicense'])*100

    def get_software_oss_obsolescence_score(self,app_name:str) -> float:
        metrics = self._get_metrics(app_name)
        return float(metrics['openSourceObsolescence'])*100

    def get_software_green_score(self,app_name:str) -> float:
        metrics = self._get_metrics(app_name)
        green_index = metrics['greenIndex']
        if green_index is None: green_index = 0
        return float(green_index)*100

    def get_software_green_hml(self,app_name:str=None,score=None) -> str:
        if score is None:
            score = self.get_software_green_score(app_name)
        if score > 75:
            return 'high'
        elif score > 52:
            return 'moderate'
        else:
            return 'low'

    def in_collection(self,collection, key,id) -> str:
        if key in collection:
            return collection[key]
        else:
            self.log.warning(f'Key {key} not found in component {id}')
            return ''




def load_df_element(src,dst,name):
    if not (src.get(name) is None):
        dst[name]=src[name] 


