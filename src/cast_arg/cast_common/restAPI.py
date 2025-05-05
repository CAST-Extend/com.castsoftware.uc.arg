import urllib.parse

from requests import get,exceptions,codes,Session
from requests.auth import HTTPBasicAuth 
from requests.adapters import HTTPAdapter, Retry
from requests.exceptions import ConnectionError, ConnectTimeout
import urllib3

#from requests.packages.urllib3.util.retry import Retry

from logging import INFO, error
from cast_common.logger import Logger
from pandas import DataFrame
from time import perf_counter, ctime
from base64 import b64decode
from inspect import stack

__author__ = "Nevin Kaplan"
__copyright__ = "Copyright 2022, CAST Software"
__email__ = "n.kaplan@castsoftware.com"


class RestCall(Logger):

    _base_url = None
    _auth = None
    _time_tracker_df  = DataFrame()
    _track_time = True
    _api_key = False
    _session = {}

    def __init__(self,*, base_url, user=None, password=None, basic_auth=None, api_key:bool=False, track_time=False,log_level=INFO):
        super().__init__(level=log_level)
        if base_url[-1]=='/': 
            base_url=base_url[:-1]
        self._base_url = base_url

        if base_url not in RestCall._session:

            RestCall._session[base_url] = Session()


            RestCall._session[base_url].verify = False
            urllib3.disable_warnings()


            max_retries = 5

            self._adapter = HTTPAdapter(
                    max_retries = Retry(
                        total = max_retries,
                        backoff_factor = 1,
                        status_forcelist = [408, 500, 502, 503, 504],
                    )
            )

            RestCall._session[base_url].mount('http://', self._adapter)
            RestCall._session[base_url].mount('https://', self._adapter)
            if basic_auth:
                up = b64decode(bytes(basic_auth,encoding='utf8')+b'==')
                (user,password)=up.decode().split(':')
            self._auth = HTTPBasicAuth(user, password)

            RestCall._session[base_url].headers.update({'Accept': 'application/json'})

            self._api_key=api_key
            if api_key:
                RestCall._session[base_url].headers.update({'X-Api-Key': password})
                RestCall._session[base_url].headers.update({'X-Api-User': user})
                
                #login to MRI rest api
                (status,rslt) = self.get('login')
                pass
                

    def get(self, url = "",header=None):
        start_dttm = ctime()
        start_tm = perf_counter()

        try:
            if len(url) > 0 and url[0] != '/':
                url=f'/{url}'
            u = urllib.parse.quote(f'{self._base_url}{url}',safe='/:&?=')

            if header is None:
                header={'Accept': 'application/json'}

            resp = RestCall._session[self._base_url].get(u, timeout = (5, 15),auth=self._auth,headers=header)
            resp.raise_for_status()

            if resp.status_code == codes.ok:
                return resp.status_code, resp.json()
            elif resp.status_code == codes.no_content:
                return resp.status_code, {}
            else:
                return resp.status_code,""

        except exceptions.ConnectionError as ex:
            self.error(f'Unable to connect to host {self._base_url}: {ex} ')
            exit()
        except exceptions.Timeout as ex:
            #TODO Maybe set up for a retry, or continue in a retry loop
            self.error(f'Timeout while performing api request using: {url}')
        except exceptions.TooManyRedirects:
            #TODO Tell the user their URL was bad and try a different one
            self.error(f'TooManyRedirects while performing api request using: {url}')
        except exceptions.HTTPError as e:
            if resp.status_code == 401:
                raise PermissionError(u)
            self.error(e)
        except exceptions.RequestException as e:
            # catastrophic error. bail.
            self.error(f'General Request exception while performing api request using: {u}')

        return 0, "{}"
    
    def post(self, url = "",data = {}):

        try:
            if len(url) > 0 and url[0] != '/':
                url=f'/{url}'
            u = urllib.parse.quote(f'{self._base_url}{url}',safe='/:&?=')

            resp = RestCall._session[self._base_url].post(u,data=data,timeout = (5, 15),auth=self._auth,headers={'Accept': 'application/json'})
            resp.raise_for_status()

            if resp.status_code == codes.ok:
                return resp.status_code, resp.json()
            elif resp.status_code == codes.no_content:
                return resp.status_code, {}
            else:
                return resp.status_code,""

        except exceptions.ConnectionError:
            self.error(f'Unable to connect to host {self._base_url}')
        except exceptions.Timeout:
            #TODO Maybe set up for a retry, or continue in a retry loop
            self.error(f'Timeout while performing api request using: {url}')
        except exceptions.TooManyRedirects:
            #TODO Tell the user their URL was bad and try a different one
            self.error(f'TooManyRedirects while performing api request using: {url}')
        except exceptions.HTTPError as e:
            self.error(e)
        except exceptions.RequestException as e:
            # catastrophic error. bail.
            self.error(f'General Request exception while performing api request using: {u}')

        return 0, "{}"

