from cast_common.highlight import Highlight
from cast_arg.powerpoint import PowerPoint
from cast_common.logger import Logger,INFO


class HLPage(Highlight):

    _ppt=None
    _log = None
    _output = None

    @property
    def ppt(self) -> PowerPoint:
        return HLPage._ppt
    
    @property
    def output(self) -> str:
        return HLPage._output
    
    def __init__(self,output:str=None,ppt:PowerPoint=None,  
                 hl_user:str=None, hl_pswd:str=None,hl_basic_auth=None, hl_instance:int=0,
                 hl_apps:str=[],hl_tags:str=[], 
                 hl_base_url:str=None, 
                 log_level=INFO):
        
        if HLPage._log is None:
            HLPage._log = Logger('HLPage',level=log_level)    
        if self.ppt is None: 
            if ppt is None:
                raise AttributeError('PPTX must be defined in the first instance of HLPage')
            else:
                HLPage._ppt=ppt

        if self.output is None: 
            if output is None:
                raise AttributeError('output must be defined in the first instance of HLPage')
            else:
                HLPage._output=output

        super().__init__(hl_user, hl_pswd,hl_basic_auth, hl_instance,hl_apps,hl_tags, hl_base_url, log_level)

        

