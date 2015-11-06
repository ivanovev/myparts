
from collections import OrderedDict as OD

class MyBOM(object):
    def init_dlg(self):
        ctx = self.ctx
        smgr = ctx.ServiceManager
        model = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)
        model.Width = 120
        model.Height = 120
        model.Title = 'BOM'
        self.center(model)

        dlg = self.createUnoService('com.sun.star.awt.UnoControlDialog')
        dlg.setModel(model)
        wnd = self.createUnoService('com.sun.star.awt.Toolkit')
        dlg.createPeer(wnd, None)
        self.init_rows(dlg)
        self.init_buttons(dlg)

        return dlg

    def setup_attributes(self, wdgtm, label, posx, posy, w, h):
        if posx == 0:
            wdgtm.Name = label + 'L'
        else:
            wdgtm.Name = label + 'W'
        wdgtm.Width = w
        wdgtm.PositionX = posx
        wdgtm.PositionY = posy + 5
        wdgtm.Height = h
        if hasattr(wdgtm, 'Label'):
            wdgtm.Label = label
        wdgtm.Align = 0
        wdgtm.Enabled = 1
        return wdgtm

    def create_text_label(self, dlg, label, posy):
        model = dlg.getModel()
        text = model.createInstance('com.sun.star.awt.UnoControlFixedTextModel')
        text = self.setup_attributes(text, label, 0, posy, model.Width/2, 10)
        model.insertByName(text.Name, text)
        self.ll.append(dlg.getControl(text.Name))
        return text

    def init_rows(self, dlg):
        model = dlg.getModel()
        posy = 0
        h1 = 15
        label = '# of boards'
        self.create_text_label(dlg, label, posy)
        spin = model.createInstance('com.sun.star.awt.UnoControlNumericFieldModel')
        spin = self.setup_attributes(spin, label, model.Width/2, posy, model.Width/2, h1)
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
        posy = model.Height/6
        self.create_text_label(dlg, label, posy)
        edit = model.createInstance('com.sun.star.awt.UnoControlEditModel')
        edit = self.setup_attributes(edit, label, model.Width/2, posy, model.Width/2, h1)
        edit.ReadOnly = 1
        model.insertByName(edit.Name, edit)
        self.cc.append(dlg.getControl(edit.Name))
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sht = ctrlr.ActiveSheet
        edit.Text = sht.Name
        label = 'Search in'
        posy = model.Height/3
        self.create_text_label(dlg, label, posy)
        srcl = model.createInstance('com.sun.star.awt.UnoControlListBoxModel')
        srcl = self.setup_attributes(srcl, label, model.Width/2, posy, model.Width/2, model.Height/2 - 5)
        srcl.Border = 1
        srcl.Dropdown = 0
        #srcl.LineCount = 3
        srcl.MultiSelection = 1
        srcl.BackgroundColor = 0xFFFFFF
        names = self.get_sheet_names()
        while sht.Name in names:
            names.remove(sht.Name)
        while 'Labels' in names:
            names.remove('Labels')
        srcl.StringItemList = tuple(names)
        model.insertByName(srcl.Name, srcl)
        self.cc.append(dlg.getControl(srcl.Name))

    def init_buttons(self, dlg):
        model = dlg.getModel()
        btn = self.init_button(dlg, model.Width/3, 1, 'Check BOM')
        listener = mp.ButtonListener(self.check_bom_cb)
        btn.addActionListener(listener)
        return dlg

    def check_bom_cb(self):
        self.dlg.endExecute()
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sht = ctrlr.ActiveSheet
        partn = 0
        while True:
            part = self.get_part(sht, partn)
            if not part:
                break
            if len(part) > mp.PART_ATTR_LEN:
                col = mp.PART_ATTR_LEN
                while col <= len(part):
                    cell = sht.getCellByPosition(col, partn)
                    cell.clearContents(1023)
                    col += 1
            partn += 1
        srcl = self.cc[-1]
        shts = srcl.SelectedItems
        if len(shts) == 0:
            return
        partn = 0
        ynrow = mp.PART_ATTR_LEN
        qty = 0
        while True:
            part = self.get_part(sht, partn)
            if not part:
                break
            if len(part) < 3:
                break
            part = part[0:3]
            qty = 0
            qtyd = OD()
            index = 0
            for shtn in shts:
                sht1 = doc.Sheets.getByName(shtn)
                index = self.part_find2(sht1, part, index)
                part2 = self.get_part(sht1, index)
                if not part2:
                    continue
                qty1 = int(part2[mp.PART_ATTR_N])
                qty += qty1
                if sht1.Name not in qtyd:
                    qtyd[sht1.Name] = qty1
                else:
                    qtyd[sht1.Name] += qty1
                index += 1
            cell = sht.getCellByPosition(mp.PART_ATTR_LEN, partn)
            part = self.get_part(sht, partn)
            qtyn = int(self.cc[0].Text)*int(part[mp.PART_ATTR_N])
            if qty < qtyn:
                cell.setString('NO')
                cell.CharColor = 0xAA0000
                cell.CharWeight = 150
            else:
                cell.setString('YES')
                cell.CharColor = 0x00AA00
                cell.CharWeight = 150
            index = mp.PART_ATTR_LEN + 1
            for k,v in qtyd.items():
                cell = sht.getCellByPosition(index, partn)
                cell.setString('%d' % v)
                index += 1
                cell = sht.getCellByPosition(index, partn)
                cell.setString(k)
                index += 1
            partn += 1

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

