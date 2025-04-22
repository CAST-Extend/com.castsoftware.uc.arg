from cast_common.restAPI import RestCall
from cast_common.logger import Logger, INFO,DEBUG
from cast_common.powerpoint import PowerPoint

from requests import codes

class MRI(RestCall):

    _log = None
    _log_level = None

    _base_url = None
    _basic_auth = None
    _api_auth = None
    _user = None 
    _pswd = None
    _api_key = None
    _domain_list = None
    
    def __init__(self,base_url=None, user=None, pswd=None, basic_auth=None, token=False, track_time=False,log_level=INFO):

        # general message to be used by ValueError exception
        msg='must be supplied with the first MRI class instance'

        if MRI._log is None:
            MRI._log_level = log_level
            MRI._log = Logger('MRI Rest',log_level)

        #set the base url, if not already set from a prev instance of the MRI class
        if MRI._base_url is None:
            if base_url is None:
                raise ValueError(f'Base url {msg}') 

            MRI._base_url = base_url
            if MRI._base_url.endswith('/'):
                MRI._base_url=MRI._base_url[:-1]
                            
            if not MRI._base_url.endswith('/rest'):
                MRI._base_url=f'{MRI._base_url}/rest/'

        #set the user name/password or Basic Autorization 
        if MRI._basic_auth is None and \
            MRI._user is None and \
            MRI._pswd is None:

            if user is None and \
               pswd is None and \
               basic_auth is None:
                    raise ValueError(f'UserName/password or Basic Authorization {msg}')

            if not user is None:
                MRI._user=user

            if not pswd is None:
                MRI._pswd=pswd

        if not basic_auth is None:
            MRI._basic_auth=basic_auth

        # if MRI._template is None:
        #     if template is None:
        #         raise ValueError(f'Template {msg}') 
        #     MRI._template=template

        super().__init__(base_url=MRI._base_url, user=MRI._user, password=MRI._pswd, api_key=token,
                         basic_auth=MRI._basic_auth,track_time=track_time,log_level=MRI._log_level)

    pass
