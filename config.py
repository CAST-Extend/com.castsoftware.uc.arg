"""
    Read and validate configuration file
"""
from logging import DEBUG, info, warn, error
from logger import Logger
from json import load
from argparse import ArgumentParser
from json import JSONDecodeError

__author__ = "Nevin Kaplan"
__copyright__ = "Copyright 2022, CAST Software"
__email__ = "n.kaplan@castsoftware.com"

class Config(Logger):

    def __init__(self, config):
        super().__init__("Config")

        #do all required fields contain data
        try:
            with open(config, 'rb') as config_file:
                self.__config = load(config_file)

            for v in ['project','application','company','template']:
                if v not in self.__config or len(self.__config[v]) == 0:
                    raise ValueError(f"Required field '{v}' is missing from config.json")

            apps = self.__config['application']
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

            if 'rest' not in self.__config:
                raise ValueError(f"Required field 'rest' is missing from config.json")

            for v in ['AIP','Highlight']:
                if v not in self.__config['rest']:
                    raise ValueError(f"Required field '{v}' is missing from config.json")

            self.__rest_settings(self.__config['rest']['AIP'])
            self.__rest_settings(self.__config['rest']['Highlight'])
            if 'instance' not in self.__config['rest']['Highlight']:
                raise ValueError(f"Required field 'instance' is missing from config.json")

        except JSONDecodeError as e:
            msg = str(e)
            self.error('Configuration file must be in a JSON format')
            exit()

        except ValueError as e:
            msg = str(e)
            self.error(msg)
            exit()

    def __rest_settings(self,dict):
        for v in ["Active","URL","user","password"]:
            if v not in dict:
                raise ValueError(f"Required field '{v}' is missing from config.json")

    @property
    def project(self):
        return self.__config['project']

    @property
    def company(self):
        return self.__config['company']

    @property
    def template(self):
        return self.__config['template']

    @property
    def output(self):
        return self.__config['output']

    @property
    def application(self):
        return self.__config['application']

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

    def aip_name(self,idx):
        return self.application[idx]['aip']

    def hl_name(self,idx):
        return self.application[idx]['highlight']

    def title(self,idx):
        return self.application[idx]['title']

    @property
    def rest(self):
        return self.__config['rest']

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

c = Config(args.config)
print (f'app 1: {c.aip_app(0)}')
"""
