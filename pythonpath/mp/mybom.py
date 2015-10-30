
class MyBOM(object):
    def init_dlg(self):
        ctx = self.ctx
        smgr = ctx.ServiceManager
        model = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
        model.Width = 120
        model.Height = 80
        model.Title = 'BOM'
        self.center(model)

        dlg = self.createUnoService('com.sun.star.awt.UnoControlDialog')
        dlg.setModel(model)
        wnd = self.createUnoService('com.sun.star.awt.Toolkit')
        dlg.createPeer(wnd, None)
        self.init_buttons(dlg)

        return dlg

    def init_buttons(self, dlg):
        model = dlg.getModel()
        btn = self.init_button(dlg, model.Width/3, 'Check BOM')
        listener = mp.ButtonListener(self.check_bom_cb)
        btn.addActionListener(listener)
        return dlg

    def check_bom_cb(self):
        self.msgbox('check bom')

def mybom(*args):
    import os, sys
    prefix = 'file://'
    fname = __file__
    if fname.find(prefix) == 0:
        fname = __file__.replace(prefix, '', 1)
    dirname = os.path.dirname(fname)
    if dirname not in sys.path:
        sys.path.append(os.path.dirname(fname))
    global mp
    #import myparts
    mp = __import__('myparts')
    MS = type('MyBOM', (MyBOM, mp.MyParts), {})
    MS(XSCRIPTCONTEXT.getComponentContext()).execute()

