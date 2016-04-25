
class MySort(object):
    def __init__(self, ctx, title='MySort'):
        self.sortasc = True
        self.x = 100
        self.y = 100
        self.w = 120
        self.h = 80
        mp.MyParts.__init__(self, ctx, title=title)

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
        for i in range(0, ms.PART_ATTR_N):
            self.init_row(dlg, i*model.Height/(ms.PART_ATTR_N+1), i)

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

    def parts_sort_cb(self):
        sht = self.get_sheet()
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
        psz = ms.PART_ATTR_LEN
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
    global ms
    ms = __import__('mysearch')
    global mp
    mp = __import__('myparts')
    mp.ms = ms
    MS = type('MySort', (MySort, mp.MyParts, ms.MySearch), {})
    MS(XSCRIPTCONTEXT.getComponentContext()).execute()

