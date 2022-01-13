import urllib.parse

from requests import get
from requests import exceptions
from requests.auth import HTTPBasicAuth 
from requests import codes

from logger import Logger
from logging import INFO, error
from pandas import DataFrame
from time import perf_counter, ctime



class RestCall(Logger):

    _base_url = None
    _auth = None
    _time_tracker_df  = DataFrame()
    _track_time = True

    def __init__(self, base_url, user=None, password=None, track_time=False,log_level=INFO):
        super().__init__(level=log_level)
        if base_url[-1]=='/': 
            base_url=base_url[:-1]
        self._base_url = base_url
        self._track_time = track_time
        self._auth = HTTPBasicAuth(user, password)

    def get(self, url = "", headers = {'Accept': 'application/json'}):
        start_dttm = ctime()
        start_tm = perf_counter()

        try:
            if len(url) > 0 and url[0] != '/':
                url=f'/{url}'
            u = urllib.parse.quote(f'{self._base_url}{url}',safe='/:&?=')
            if self._auth == None:
                resp = get(url= u, headers=headers)
            else:
                resp = get(url= u, auth = self._auth, headers=headers)

            # Save the duration, if enabled.
            if (self._track_time):
                end_tm = perf_counter()
                end_dttm = ctime()
                duration = end_tm - start_tm

                #print(f'Request completed in {duration} ms')
                self._time_tracker_df = self._time_tracker_df.append({'Application': 'ALL', 'URL': url, 'Start Time': start_dttm \
                                                            , 'End Time': end_dttm, 'Duration': duration}, ignore_index=True)
            if resp.status_code == codes.ok:
                return resp.status_code, resp.json()
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
        except exceptions.RequestException as e:
            # catastrophic error. bail.
            self.error(f'General Request exception while performing api request using: {url}')

        return 0, "{}"
