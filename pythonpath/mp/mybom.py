
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
        self.init_rows(dlg)
        self.init_buttons(dlg)

        return dlg

    def init_rows(self, dlg):
        model = dlg.getModel()
        posy = 0
        label = '# of boards'
        text = model.createInstance('com.sun.star.awt.UnoControlFixedTextModel')
        text.Name = label + 'L'
        text.Width = model.Width/3
        text.PositionX = 0
        text.PositionY = posy + 5
        text.Height = 10
        text.Label = label
        text.Align = 0
        text.Enabled = 1
        model.insertByName(text.Name, text)
        self.ll.append(dlg.getControl(text.Name))
        spin = model.createInstance('com.sun.star.awt.UnoControlNumericFieldModel')
        spin.Name = label + 'W'
        spin.Width = model.Width/2
        spin.PositionX = model.Width/2
        spin.PositionY = posy + 5
        spin.Height = 15
        spin.Enabled = 1
        spin.Spin = 1
        spin.DecimalAccuracy = 0
        spin.Value = 5
        spin.ValueMin = 1
        spin.ValueMax = 100
        spin.ValueStep = 1
        model.insertByName(spin.Name, spin)
        c = dlg.getControl(spin.Name)
        self.cc.append(c)
        label = 'BOM'
        posy = model.Height/3
        text = model.createInstance('com.sun.star.awt.UnoControlFixedTextModel')
        text.Name = label + 'L'
        text.Width = model.Width/3
        text.PositionX = 0
        text.PositionY = posy + 5
        text.Height = 10
        text.Label = label
        text.Align = 0
        text.Enabled = 1
        model.insertByName(text.Name, text)
        self.ll.append(dlg.getControl(text.Name))
        edit = model.createInstance('com.sun.star.awt.UnoControlEditModel')
        edit.Name = label + 'W'
        edit.Width = model.Width/2
        edit.PositionX = model.Width/2
        edit.PositionY = posy + 5
        edit.Height = 15
        edit.Align = 0
        edit.Enabled = 1
        edit.ReadOnly = 1
        model.insertByName(edit.Name, edit)
        self.cc.append(dlg.getControl(edit.Name))
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sht = ctrlr.ActiveSheet
        edit.Text = sht.Name

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

