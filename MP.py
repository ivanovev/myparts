
import unohelper
from com.sun.star.task import XJobExecutor

class MP(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, *args):
        import mp
        a = list(args)
        if a[0] == 'myparts':
            mp.myparts.ms = mp.mysearch
            mp.myparts.MyParts(self.ctx).execute()
        elif a[0] == 'mysort':
            mp.mysort.mp = mp.myparts
            MS = type('MySort', (mp.mysort.MySort, mp.myparts.MyParts), {})
            MS(self.ctx).execute()
        elif a[0] == 'mybom':
            mp.mybom.mp = mp.myparts
            MB = type('MyBOM', (mp.mybom.MyBOM, mp.myparts.MyParts), {})
            MB(self.ctx).execute()

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(MP, 'vnd.myparts', ('com.sun.star.task.Job',),)
g_ImplementationHelper.addImplementation(MP, 'vnd.mysort', ('com.sun.star.task.Job',),)
g_ImplementationHelper.addImplementation(MP, 'vnd.mybom', ('com.sun.star.task.Job',),)

