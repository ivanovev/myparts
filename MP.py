
import unohelper
from com.sun.star.task import XJobExecutor

class MP(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, *args):
        from mp import MyParts
        MyParts(self.ctx).execute()

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(MP, 'vnd.myparts', ('com.sun.star.task.Job',),)

