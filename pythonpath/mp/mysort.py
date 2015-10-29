
from .myparts import *

class MySort(MyParts):
    def __init__(self, ctx):
        MyParts.__init__(self, ctx)

    def init_dlg(self):
        '''
        ctx = self.ctx
        smgr = ctx.ServiceManager
        model = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
        model.Width = 120
        model.Height = 150
        model.Title = 'My Parts'
        self.center(model)

        dlg = self.createUnoService('com.sun.star.awt.UnoControlDialog')
        dlg.setModel(model)
        wnd = self.createUnoService('com.sun.star.awt.Toolkit')
        dlg.createPeer(wnd, None)
        self.init_rows(dlg)
        self.init_buttons(dlg)
        self.init_data(dlg)

        return dlg
        '''

def mysort(*args):
    MySort(XSCRIPTCONTEXT.getComponentContext()).execute()

