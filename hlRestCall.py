from restAPI import RestCall
from requests import codes
from pandas import DataFrame
from pandas import json_normalize
from pandas import concat
from logging import DEBUG, INFO, ERROR, warning


class HLRestCall(RestCall):
    """
    Class to handle HL REST API calls.
    """
    def __init__(self, hl_base_url, hl_user, hl_pswd, hl_instance, timer_on=False,log_level=INFO):
        super().__init__(hl_base_url, hl_user, hl_pswd, timer_on,log_level)

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
            raise KeyError (f'Application not found: {app_name}')

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
                    rslt = concat([rslt,cd],ignore_index=True)
            except KeyError as e:
                self.warning('Error retrieving cloud ready information')
            return rslt
        else: 
            return None


    def get_third_party(self, app_id):
        cves = DataFrame()
        lic = DataFrame()

        url = f'domains/{self._hl_instance}/applications/{app_id}/thirdparty'
        (status, json) = self.get(url)

        third_party = []
        if status == codes.ok and len(json) > 0:
            third_party = json['thirdParties']
            for tp in third_party:
                if 'cve' in tp:
                    cve_df = json_normalize(tp['cve']['vulnerabilities'])
                    cve_df.rename(columns={'name':'cve'},inplace=True)
                    
                    cve_df['component']=tp['name']
                    cve_df['version']=tp['version']
                    cve_df['languages']=tp['languages']
                    cve_df['release']=tp['release']
                    cve_df['origin']=tp['origin']
                    cve_df['lastVersion']=tp['lastVersion']

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

            if 'component' in cves.columns:
                cves=cves[['component','version','languages','release','origin','lastVersion','cve', 'description', 'cweId', 'cweLabel', 'criticity', 'cpe']]
            if 'component' in lic.columns:
                lic=lic[['component','version','languages','release','origin','lastVersion','license','compliance']] 

        return lic,cves,len(third_party)

def load_df_element(src,dst,name):
    if not (src.get(name) is None):
        dst[name]=src[name] 
