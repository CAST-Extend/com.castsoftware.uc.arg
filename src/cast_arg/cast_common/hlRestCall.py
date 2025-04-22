from cast_common.restAPI import RestCall
from requests import codes, post
from pandas import DataFrame
from pandas import json_normalize
from pandas import concat
from logging import DEBUG, INFO, ERROR, warning


class HLRestCall(RestCall):
    """
    Class to handle HL REST API calls.
    """
    def __init__(self, hl_base_url:str, hl_user:str, hl_pswd:str, hl_instance:int, timer_on=False,log_level=INFO):

        if hl_base_url.endswith('/'):
            hl_base_url = hl_base_url[:-1]
        if not hl_base_url.endswith(r'/WS2'):
            hl_base_url=f'{hl_base_url}/WS2'

        super().__init__(base_url=hl_base_url,user=hl_user,password=hl_pswd,track_time=timer_on,log_level=log_level)

        self._hl_instance = hl_instance
        self._hl_data_retrieved = False
    
    def _get_app_ids(self, instance_id):
        # Retrieve the HL app id for the application.

        try:
            # TODO: remove the hard-coding
            # TODO: Get the app id.
            url = f'domains/{instance_id}/applications'

            (status, json) = self.get(url,headers={'Accept': '*/*'})

            # TODO: Handle exceptions
            if status == codes.ok and len(json) > 0:
                pass
        except:
            # TODO
            print('Oopsi.. caught an exception')
            raise

        return json

    def get_appl_third_party(self,appl_id):
        url = f'domains/{self._hl_instance}/applications/{appl_id}/thirdparty'
        (status, json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            return json
        else:
            raise KeyError (f'Application not found {appl_id}')


    def get_appls(self):
        url = f'domains/{self._hl_instance}/applications/'
        (status, json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            return json
        else:
            raise KeyError (f'No applications not found')

    def get_app_id(self,app_name):
        url = f'domains/{self._hl_instance}/applications/'
        (status, json) = self.get(url)

        # TODO: Handle exceptions
        if status == codes.ok and len(json) > 0:
            for id in json:
                if id['name'].lower()==app_name.lower():
                    return id['id']
            #raise KeyError (f'Highlight application not found: {app_name}')
            return None

    def get_cloud_data(self,app_id):
        url = f'domains/{self._hl_instance}/applications/{app_id}'
        (status, json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            rslt = DataFrame()
            try:
                cloud_data = json['metrics'][0]['cloudReadyDetail']
                for d in cloud_data:
                    cd = json_normalize(d['cloudReadyDetails'])
                    cd['Technology']=d['technology']
                    cd['Scan']=d['cloudReadyScan']

                    if not rslt.empty and not cd.empty:
                        rslt = concat([rslt, cd], ignore_index=True)
                    elif rslt.empty and not cd.empty:
                        rslt = cd.copy()
                    pass
                    # if rslt.empty:
                    #     rslt = cd.copy()
                    # elif(not cd.empty):
                    #     rslt = concat([rslt,cd],ignore_index=True)
            except KeyError as e:
                self.warning('Error retrieving cloud ready information')
            return rslt
        else: 
            return None

    def in_collection(self,collection, key,id) -> str:
        if key in collection:
            return collection[key]
        else:
            self.warning(f'Key {key} not found in component {id}')
            return ''

    def get_third_party(self, app_id):
        cves = DataFrame()
        lic = DataFrame()

        self.info(f'Collecting third party information')

        url = f'domains/{self._hl_instance}/applications/{app_id}/thirdparty'
        (status, json) = self.get(url)

        third_party = []
        try:
            if status == codes.ok and len(json) > 0:
                third_party = json['thirdParties']
                for tp in third_party:
                    try:
                        if 'cve' in tp:
                            cve_df = json_normalize(tp['cve']['vulnerabilities'])
                            cve_df.rename(columns={'name':'cve'},inplace=True)
                            
                            cve_df['component']=self.in_collection(tp,'name',tp['name'])
                            cve_df['version']=self.in_collection(tp,'version',tp['name'])
                            cve_df['languages']=self.in_collection(tp, 'languages',tp['name'])  
                            cve_df['release']=self.in_collection(tp,'version',tp['name']) 
                            cve_df['origin']=self.in_collection(tp,'origin',tp['name']) 
                            cve_df['lastVersion']=self.in_collection(tp,'lastVersion',tp['name'])

                            cves=concat([cves,cve_df],ignore_index=True)

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
                    except KeyError as e:
                        self.warning(f'Key Error while retrieving third party information: {e} {tp}')    
                        pass                        
                if not cves.empty and 'component' in cves.columns:
                    cves=cves[['component','version','languages','release','origin','lastVersion','cve', 'description', 'cweId', 'cweLabel', 'criticity', 'cpe']]
                else:
                    cves=None
                    
                if not lic.empty and  'component' in lic.columns:
                    lic=lic[['component','version','languages','release','origin','lastVersion','license','compliance']] 
                else:
                    lic=None
        except Exception as e:
            self.error('Error retrieving third party information')
            ex_type, ex_value, ex_traceback = exc_info()
            self.log.warning(f'{ex_type.__name__}: {ex_value} while in {__class__}')
            raise e 
                    
        return lic,cves,len(third_party)
    
    def create_an_app(self, instance_id, app_name):
        url = f'{self._base_url}/domains/{instance_id}/applications/'
        payload =[{"name": app_name,"domains": [{"id": instance_id}]}]

        resp = post(url, auth=self._auth, json=payload)        
        #print(resp.status_code)

        return resp.status_code

def load_df_element(src,dst,name):
    if not (src.get(name) is None):
        dst[name]=src[name] 
