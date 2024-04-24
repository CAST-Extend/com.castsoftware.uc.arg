from cast_arg.restCall import AipData
from cast_arg.powerpoint import PowerPoint
from cast_common.logger import Logger,DEBUG,INFO



class MRIPage(AipData):
    _ppt=None
    _log = None

    @property
    def ppt(self):
        return MRIPage._ppt

    def __init__(self,ppt:PowerPoint=None,log_level=INFO):
        if MRIPage._log is None:
            MRIPage._log = Logger('MRIPage',level=log_level)    
        if MRIPage._ppt is None: 
            if ppt is None:
                raise AttributeError('PPTX must be defined in the first instance of MRIPage')
            else:
                MRIPage._ppt=ppt

            
        super().__init__()

    def run(self,app_name:str,app_no:int) -> bool:
        try:
            return self.report(app_name,app_no)
        except ValueError as e:
            self._log.warning(e)
            return False
