
class MySort(object):
    def init_dlg(self):
        self.sortasc = True
        ctx = self.ctx
        smgr = ctx.ServiceManager
        model = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
        model.Width = 120
        model.Height = 80
        model.Title = 'Sort'
        self.center(model)

        dlg = self.createUnoService('com.sun.star.awt.UnoControlDialog')
        dlg.setModel(model)
        wnd = self.createUnoService('com.sun.star.awt.Toolkit')
        dlg.createPeer(wnd, None)
        self.init_buttons(dlg)
        self.init_rows(dlg)

        return dlg

    def init_buttons(self, dlg):
        model = dlg.getModel()
        btn = self.init_button(dlg, model.Width/6, 1, 'Sort')
        listener = mp.ButtonListener(self.parts_sort_cb)
        btn.addActionListener(listener)
        btn = self.init_button(dlg, model.Width/2, 1, 'Close')
        listener = mp.ButtonListener(self.close_cb)
        btn.addActionListener(listener)
        return dlg

    def init_row(self, dlg, posy, index):
        label = 'Sort key %d' % (index+1)
        model = dlg.getModel()
        text = model.createInstance('com.sun.star.awt.UnoControlFixedTextModel')
        text.Name = label + 'L'
        text.Width = model.Width/3
        text.PositionX = 0#model.Width/4 - text.Width/2
        text.PositionY = posy + 5
        text.Height = 10
        text.Label = label
        text.Align = 0
        text.Enabled = 1
        model.insertByName(text.Name, text)
        l = dlg.getControl(text.Name)
        self.ll.append(l)
        labels = ['A', 'B', 'C']
        cc = []
        for i in range(0, len(labels)):
            radio = model.createInstance('com.sun.star.awt.UnoControlRadioButtonModel')
            radio.Label = labels[i]
            radio.Name = label + 'W' + radio.Label
            radio.Width = 2*model.Width/3/len(labels)
            radio.PositionX = model.Width/3 + i*radio.Width
            radio.PositionY = posy + 5
            radio.Height = 15
            radio.Enabled = 1
            model.insertByName(radio.Name, radio)
            c = dlg.getControl(radio.Name)
            if i == index:
                c.State = 1
            listener = mp.ButtonListener(lambda: self.key_upd_cb(index))
            c.addActionListener(listener)
            cc.append(c)
        self.cc.append(cc)

    def init_rows(self, dlg):
        model = dlg.getModel()
        for i in range(0, mp.PART_ATTR_N):
            self.init_row(dlg, i*model.Height/(mp.PART_ATTR_N+1), i)

    def key_upd_cb(self, index):
        cc = self.cc[index]
        kk = self.get_keys()
        k0 = set(range(0, len(self.cc)))
        k0 = k0 - set(kk)
        for i in range(0, len(self.cc)):
            if i == index:
                continue
            if kk[i] == kk[index]:
                self.cc[i][k0.pop()].setState(1)
                break

    def get_keys(self):
        kk = []
        for i in range(0, len(self.cc)):
            cc = self.cc[i]
            for j in range(0, len(cc)):
                if cc[j].State == 1:
                    kk.append(j)
                    break
        return kk

    def set_part(self, sht, row, p, psz):
        if not sht:
            sht = self.get_active_sheet()
        for i in range(0, psz):
            cell = sht.getCellByPosition(i, row)
            if i < len(p):
                cell.setString(p[i])
            else:
                cell.setString('')

    def parts_sort_cb(self):
        self.dlg.endExecute()
        sht = self.get_active_sheet()
        pp = []
        row = 0
        while True:
            p = self.get_part(sht, row)
            row = row + 1
            if p:
                pp.append(p)
            else:
                break
        self.sortasc = not self.sortasc
        kk = self.get_keys()
        pp = sorted(pp, key=lambda p: '.'.join([p[k] for k in kk]), reverse=self.sortasc)
        psz = mp.PART_ATTR_LEN
        for p in pp:
            if len(p) > psz:
                psz = len(p)
        for i in range(0, len(pp)):
            self.set_part(sht, i, pp[i], psz)

def mysort(*args):
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
    MS = type('MySort', (MySort, mp.MyParts), {})
    MS(XSCRIPTCONTEXT.getComponentContext()).execute()

