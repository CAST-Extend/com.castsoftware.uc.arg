"""
    Read and validate configuration file
"""
from cast_common.logger import Logger,DEBUG, INFO, WARN, ERROR
from json import load
from argparse import ArgumentParser
from json import JSONDecodeError

__author__ = "Nevin Kaplan"
__copyright__ = "Copyright 2022, CAST Software"
__email__ = "n.kaplan@castsoftware.com"

class Config():
    log = None
    log_translate = {} 
    def __init__(self, config:str):
        #super().__init__("Config")

        self.log_translate['info']=INFO
        self.log_translate['warn']=WARN
        self.log_translate['error']=ERROR
        self.log_translate['debug']=DEBUG

        #do all required fields contain data
        try:
            with open(config, 'rb') as config_file:
                self._config = load(config_file)

            #get logging configuration
            # is there a logging, if not add it now?
            if 'logging' not in self._config:
                self._config['logging']={}

            # if any logging entries are missing, add them now   
            log_config = self._config['logging']
            for v in ['config','aip','highlight','generate']:
                if v not in log_config:
                   log_config[v]='info'

            # convert the entries into something the logger can understand
            for idx,value in enumerate(log_config):
                log_config[value]=self.log_translate[log_config[value]]
            self.log=Logger('config',log_config['config'])

            for v in ['project','application','company','template']:
                self.log.debug(f'Validating {v}')
                if v not in self._config or len(self._config[v]) == 0:
                    raise ValueError(f"Required field '{v}' is missing from config.json")

            apps = self._config['application']
            for idx,value in enumerate(apps):
                if 'aip' in value: 
                    if len(value['aip']) == 0:
                        raise ValueError(f'No AIP id for application {idx+1}')
                    aip=value['aip']
                
                #title and highlight values are both optional
                #if they are not filled in then populate with the aip value
                if 'highlight' not in value or len(value['highlight'])==0:
                    value['highlight']=aip
                if 'title' not in value or len(value['title'])==0:
                    value['title']=aip

            if 'rest' not in self._config:
                raise ValueError(f"Required field 'rest' is missing from config.json")

            for v in ['AIP','Highlight']:
                if v not in self._config['rest']:
                    raise ValueError(f"Required field '{v}' is missing from config.json")

            self.__rest_settings(self._config['rest']['AIP'])
            self.__rest_settings(self._config['rest']['Highlight'])
            if 'instance' not in self._config['rest']['Highlight']:
                raise ValueError(f"Required field 'instance' is missing from config.json")

        except JSONDecodeError as e:
            msg = str(e)
            self.error('Configuration file must be in a JSON format')
            exit()

        except ValueError as e:
            msg = str(e)
            self.log.error(msg)
            exit()

    def __rest_settings(self,dict):
        for v in ["Active","URL","user","password"]:
            if v not in dict:
                raise ValueError(f"Required field '{v}' is missing from config.json")

    @property
    def project(self):
        return self._config['project']

    @property
    def company(self):
        return self._config['company']

    @property
    def template(self):
        return self._config['template']

    @property
    def output(self):
        return self._config['output']

    @property
    def cause(self):
        return self._config['cause']

    @property
    def application(self):
        return self._config['application']

    @property
    def aip_list(self):
        rslt = []
        for app in self.application:
            rslt.append(app['aip'])
        return rslt

    @property
    def hl_list(self):
        rslt = []
        for app in self.application:
            rslt.append(app['highlight'])
        return rslt

    @property
    def title_list(self):
        rslt = []
        for app in self.application:
            rslt.append(app['title'])
        return rslt

    @property
    def logging_aip(self):
        return self._config['logging']['aip']
    @property
    def logging_highlight(self):
        return self._config['logging']['highlight']
    @property
    def logging_generate(self):
        return self._config['logging']['generate']

    def aip_name(self,idx):
        return self.application[idx]['aip']

    def hl_name(self,idx):
        return self.application[idx]['highlight']

    def title(self,idx):
        return self.application[idx]['title']

    @property
    def rest(self):
        return self._config['rest']

    @property
    def aip_active(self):
        return self.rest['AIP']['Active']

    @property
    def aip_url(self):
        return self.rest['AIP']['URL']

    @property
    def aip_user(self):
        return self.rest['AIP']['user']

    @property
    def aip_password(self):
        return self.rest['AIP']['password']

    @property
    def hl_active(self):
        return self.rest['Highlight']['Active']

    @property
    def hl_url(self):
        return self.rest['Highlight']['URL']

    @property
    def hl_user(self):
        return self.rest['Highlight']['user']

    @property
    def hl_password(self):
        return self.rest['Highlight']['password']

    @property
    def hl_instance(self):
        return self.rest['Highlight']['instance']




"""
parser = ArgumentParser(description='CAST Assessment Report Generation (ARG)')
parser.add_argument('-c','--config', required=True, help='Configuration properties file')
args = parser.parse_args()

c = ARGConfig(args.config)
print (f'app 1: {c.aip_app(0)}')
"""
