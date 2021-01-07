import requests
from requests.auth import HTTPBasicAuth 
from time import perf_counter, ctime
import pandas as pd
import logging
import enum

class RestCall:

    _base_url = None
    _auth = None
    _time_tracker_df  = pd.DataFrame()
    _track_time = True

    def __init__(self, base_url, user, password, track_time=False):
        self._logger = logging.getLogger(__name__)
        shandler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(filename)s [%(funcName)30s:%(lineno)-4d] %(levelname)-8s - %(message)s')
        shandler.setFormatter(formatter)
        self._logger.addHandler(shandler)

        self._base_url = base_url
        self._track_time = track_time
        self._auth = HTTPBasicAuth(user, password)
        self.setLoggingLevel(logging.INFO)

    def setLoggingLevel(self,level):
        self._logger.setLevel(level)

    def get(self,url=""):
        start_dttm = ctime()
        start_tm = perf_counter()

        # TODO: Errorhandling
        
        headers = {'Accept': 'application/json'}
        u = f'{self._base_url}{url}'
        resp = requests.get(url= u, auth = self._auth, headers=headers)

        # Save the duration, if enabled.
        if (self._track_time):
            end_tm = perf_counter()
            end_dttm = ctime()
            duration = end_tm - start_tm

            #print(f'Request completed in {duration} ms')
            self._time_tracker_df = self._time_tracker_df.append({'Application': 'ALL', 'URL': url, 'Start Time': start_dttm \
                                                        , 'End Time': end_dttm, 'Duration': duration}, ignore_index=True)
        
        # TODO: Errorhandling

        return resp.status_code,resp.json()


class AipRestCall(RestCall):
    _measures = {
        '60017':'TQI',
        '60012':'Changeability',
        '60014':'Efficiency',
        '60013':'Robustness',
        '60016':'Security',
        '60011':'Transferability',
        '60015':'SEI Maintainability',
    }

    def getDomain(self,schema_name):
        domain_id = None
        (status,json) = self.get()
        if status == requests.codes.ok:
            try: 
                domain_id = list(filter(lambda x:x["schema"]==schema_name,json))[0]['name']
            except IndexError:
                self._logger.error(f'Domain not found for schema {schema_name}')
                
        return domain_id

    def getLatestSnapshot(self,domain_id):
        snapshot = {}
        (status,json) = self.get(f'{domain_id}/applications/3/snapshots')
        if status == requests.codes.ok and len(json) > 0:
            snapshot['id'] = json[0]['number']
            snapshot['name'] = json[0]['name']
            snapshot['technology'] = json[0]['technologies']
            snapshot['module_href'] = json[0]['moduleSnapshots']['href']
            snapshot['result_href'] = json[0]['results']['href'] 
        return snapshot 

    def getGradesByTechnolgy(self,domain_id,snapshot):
        first_tech=True
        grade = pd.DataFrame(columns=list(self._measures.values()))
        for tech in snapshot['technology']:
            t={}
            a={}
            for key in self._measures: 
                url = f'{domain_id}/applications/3/results?quality-indicators={key}&technologies={tech}'
                (status,json) = self.get(url)
                if status == requests.codes.ok and len(json) > 0:
                    try:
                        t[self._measures[key]]=json[0]['applicationResults'][0]['technologyResults'][0]['result']['grade']
                        if first_tech==True:
                            a[self._measures[key]]=json[0]['applicationResults'][0]['result']['grade']
                    except IndexError:
                        self._logger.debug(f'{domain_id} no grade available for {key} {tech}')
            if first_tech==True:
                grade.loc['All'] = a
            grade.loc[tech] = t
            first_tech=False
        return grade

    def getLOC(self,domain_id):
        loc = 0
        (status,json) = self.get(f'{domain_id}/applications/3/results?sizing-measures=10151&snapshots=-1')
        if status == requests.codes.ok and len(json) > 0:
            loc = json[0]['applicationResults'][0]['result']['value']
        return loc

    def getSizing(self, domain_id, input):
        rslt = {}
        for key in input: 
            (status,json) = self.get(f'{domain_id}/applications/3/results?sizing-measures={key}&snapshots=-1')
            if status == requests.codes.ok and len(json) > 0:
                rslt[input[key]]=json[0]['applicationResults'][0]['result']['value']
        return rslt


class AipData():
    _data={}
    _base=[]
    _sizing = {
       '10151':'Number of Code Lines', 
       '10107':'Number of Comment Lines', 
       '10109':'Number of Commented-out Code Lines' 
    }

    def __init__(self, rest, project, schema):
        self._base=schema
        for s in schema:
            self._data[s]={}
            central_schema = f'{s}_central'
            domain_id = rest.getDomain(central_schema)
            if domain_id is not None:
                self._data[s]['domain_id']=domain_id
                self._data[s]['snapshot']=rest.getLatestSnapshot(domain_id)
                self._data[s]['grades']=rest.getGradesByTechnolgy(domain_id,self._data[s]['snapshot'])
                self._data[s]['loc_sizing']=rest.getSizing(domain_id,self._sizing) 

    def data(self,app):
        return self._data[app]

    def domain(self, app):
        return self.data(app)['domain_id']

    def snapshot(self, app):
        return self.data(app)['snapshot']

    def grades(self, app):
        return self.data(app)['grades']

    def get_app_grades(self, app, sort=False):
        app_grades = self.grades(app).loc['All']
        if sort:
            return app_grades
        else:
            return app_grades.sort_values() 
    
    def get_loc_sizing(self,app):
        return self.data(app)['loc_sizing']

    def get_all_app_text(self):
        rslt = ""

        data = self._data
        l = len(self._base)
        last_name = self._base[-2]

        for a in self._base:
            rslt = rslt + self.snapshot(a)['name']
            if l >= 2 and a == last_name:
                rslt = rslt + " and "
            elif a != self._base[-1]:
                rslt = rslt + ", "
        return rslt




aip_rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)
apps = ["actionplatform","intersect"] 
aip_data = AipData(aip_rest,"Florence", apps)
print (aip_data.get_all_app_text())

print (aip_data.data(apps[0])['snapshot'])




