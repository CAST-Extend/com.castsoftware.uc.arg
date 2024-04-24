from cast_common.highlight import Highlight
from cast_arg.powerpoint import PowerPoint
from cast_common.logger import Logger,INFO
from pandas import DataFrame


class HLPage(Highlight):

    _ppt=None
    _log = None
    _output = None
    _benchmark = None

    _tag_prefix = None

    @property
    def ppt(self) -> PowerPoint:
        return HLPage._ppt

    @property
    def tag_prefix(self) -> str:
        if self._tag_prefix is None:
            raise ValueError(f'tag prefix is not set')
        return self._tag_prefix
    @tag_prefix.setter
    def tag_prefix(self,value):
        self._tag_prefix=f'{value}_hl'
   
    @property
    def output(self) -> str:
        return HLPage._output
    
    @property
    def log(self) -> str:
        return HLPage._log
    
    @property
    def ppt(self) -> str:
        return HLPage._ppt
        
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

        if HLPage._benchmark is None:
            HLPage._benchmark = DataFrame(self._get(r'/benchmark'))


    def replace_text(self, item, data,shape=False,slide=None):
        tag = f'{self.tag_prefix}_{item}'
        self.log.debug(f'{tag}: {data}')
        if shape:
            # slide = PowerPoint.ppt._prs.slides[3]
            PowerPoint.ppt.replace_textbox(tag,data,slide=slide)
        else:
            tag = f'{{{tag}}}'
            PowerPoint.ppt.replace_text(tag,data,slide=slide)
        

