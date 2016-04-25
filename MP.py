
import unohelper
from com.sun.star.task import XJobExecutor

class MP(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, *args):
        import mp
        a = list(args)
        mp.myparts.ms = mp.mysearch
        if a[0] == 'myparts':
            MP = type('MyParts', (mp.myparts.MyParts, mp.mysearch.MySearch), {})
            MP(self.ctx).execute()
        elif a[0] == 'mysort':
            mp.mysort.mp = mp.myparts
            mp.mysort.ms = mp.mysearch
            MS = type('MySort', (mp.mysort.MySort, mp.myparts.MyParts, mp.mysearch.MySearch), {})
            MS(self.ctx).execute()
        elif a[0] == 'mybom':
            mp.mybom.mp = mp.myparts
            mp.mybom.ms = mp.mysearch
            MB = type('MyBOM', (mp.mybom.MyBOM, mp.myparts.MyParts, mp.mysearch.MySearch), {})
            MB(self.ctx).execute()

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(MP, 'vnd.myparts', ('com.sun.star.task.Job',),)
g_ImplementationHelper.addImplementation(MP, 'vnd.mysort', ('com.sun.star.task.Job',),)
g_ImplementationHelper.addImplementation(MP, 'vnd.mybom', ('com.sun.star.task.Job',),)

