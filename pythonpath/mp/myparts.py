
import uno, unohelper
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.PushButtonType import STANDARD, OK, CANCEL
from com.sun.star.sheet.CellDeleteMode import UP
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.awt.MouseButton import RIGHT
from com.sun.star.awt import XActionListener, XTextListener, XMouseListener
from com.sun.star.table import CellRangeAddress
from com.sun.star.table.CellVertJustify import CENTER as VJ_CENTER
from com.sun.star.table.CellHoriJustify import CENTER as HJ_CENTER

PART_ATTR_LIST = ['Device', 'Value', 'Footprint', 'Quantity', 'Note1', 'Note2']
PART_ATTR_LEN = len(PART_ATTR_LIST)
PART_ATTR_ITEMS = [
    ('RESISTOR', 'CAPACITOR', 'INDUCTOR', 'CONNECTOR', 'LED'),
    ('1p', '1n', '1u', '1m', '1', '1k', '1M'),
    ('0402', '0603', '0805', '1206'),
    ('1', '2', '3', '4', '5'),
    (),
    ()
]
PART_ATTR_DFLT = [
    'RESISTOR',
    '1k',
    '0603',
    '1',
    '',
    ''
]

PART_ATTR_N = 3

def out(obj):
    f = open('/tmp/out', 'w')
    for i in dir(obj):
        f.write(str(i))
        f.write('\n\n\r')
    f.close()

class ButtonListener(unohelper.Base, XActionListener):
    def __init__(self, cb):
        self.cb = cb
    def actionPerformed(self, evt):
        self.cb()

class LabelListener(unohelper.Base, XMouseListener):
    def __init__(self, cb):
        self.cb = cb
    def mouseReleased(self, evt):
        self.cb(evt)

class ComboboxListener(unohelper.Base, XTextListener):
    def __init__(self, cb):
        self.cb = cb
    def textChanged(self, evt):
        self.cb(evt.Source)

def part_add_del_dec(f):
    def tmp(*args, **kwargs):
        mp = args[0]
        if not mp.ll[PART_ATTR_N].getState():
            return
        part = [a.Text for a in mp.cc]
        pn = part[PART_ATTR_N]
        part[PART_ATTR_N] = ''
        index = mp.get_selection()
        if index == None:
            mp.part_find()
            mp.cc[PART_ATTR_N].setText(pn)
            return
        sht = mp.get_active_sheet()
        if not mp.part_cmp(sht, index, part):
            mp.part_find()
            mp.cc[PART_ATTR_N].setText(pn)
            return
        p1 = mp.get_part(sht, index)
        if ''.join(p1) == '' and index > 0:
            if mp.part_find2(sht, part, 0) != index:
                mp.part_find()
                mp.cc[PART_ATTR_N].setText(pn)
                return
        return f(*args, **kwargs)
    return tmp

class MyParts(object):
    def __init__(self, ctx):
        self.ctx = ctx
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.ll = []
        self.cc = []
        self.dlg = self.init_dlg()

    def init_dlg(self):
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
        self.part_dlg_combo_upd(self.cc[0])

        return dlg

    def init_row(self, dlg, posy, label, itemlist):
        model = dlg.getModel()
        text = model.createInstance('com.sun.star.awt.UnoControlCheckBoxModel')
        text.Name = label + 'L'
        text.Width = model.Width/3
        text.PositionX = 0#model.Width/4 - text.Width/2
        text.PositionY = posy + 5
        text.Height = 10
        text.Label = label
        text.Align = 0
        text.Enabled = 1
        model.insertByName(text.Name, text)
        self.ll.append(dlg.getControl(text.Name))
        combo = model.createInstance('com.sun.star.awt.UnoControlComboBoxModel')
        combo.Name = label + 'W'
        combo.Width = 2*model.Width/3
        combo.PositionX = 2*model.Width/3 - combo.Width/2
        combo.PositionY = posy
        combo.Height = 15
        combo.Enabled = 1
        combo.Dropdown = 1
        combo.StringItemList = itemlist
        combo.Border = 1
        model.insertByName(combo.Name, combo)
        self.cc.append(dlg.getControl(combo.Name))
        return self.ll[-1], self.cc[-1]

    def init_rows(self, dlg):
        model = dlg.getModel()
        for i in range(0, PART_ATTR_LEN):
            label = PART_ATTR_LIST[i]
            #items = PART_ATTR_ITEMS[i]
            items = self.get_combo_itemlist(i)
            l,w = self.init_row(dlg, i*model.Height/(PART_ATTR_LEN + 2), label, items)
            l.setState(1)
            w.Text = PART_ATTR_DFLT[i]
            listener = LabelListener(self.part_dlg_label_upd)
            l.addMouseListener(listener)
            listener = ComboboxListener(self.part_dlg_combo_upd)
            w.addTextListener(listener)

    def init_button(self, dlg, posx, rown, label):
        model = dlg.getModel()
        btnmodel = model.createInstance('com.sun.star.awt.UnoControlButtonModel')
        btnmodel.Name = label
        btnmodel.Width = model.Width/3
        btnmodel.Height = 15
        btnmodel.PositionX = posx
        btnmodel.PositionY = model.Height - (2-rown)*btnmodel.Height
        btnmodel.Label = label
        btnmodel.PushButtonType = STANDARD
        model.insertByName(label, btnmodel)
        return dlg.getControl(label)

    def init_buttons(self, dlg):
        model = dlg.getModel()
        btn = self.init_button(dlg, 0, 0, 'Find')
        listener = ButtonListener(lambda: self.part_find())
        btn.addActionListener(listener)
        btn = self.init_button(dlg, model.Width/3, 0, 'Add')
        listener = ButtonListener(lambda: self.part_add())
        btn.addActionListener(listener)
        btn = self.init_button(dlg, 2*model.Width/3, 0, 'Del')
        listener = ButtonListener(lambda: self.part_del())
        btn.addActionListener(listener)
        btn = self.init_button(dlg, model.Width/3, 1, 'Add label')
        listener = ButtonListener(lambda: self.add_label())
        btn.addActionListener(listener)
        return dlg

    def init_data(self, dlg):
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sel = self.get_selection()
        if sel == None:
            return
        part = self.get_part(None, sel)
        l = min(len(part), PART_ATTR_LEN)
        for i in range(0, l):
            self.cc[i].setText(part[i])

    def get_combo_itemlist(self, index):
        itemlist = PART_ATTR_ITEMS[index]
        if index >= PART_ATTR_N:
            return itemlist
        ii = set()
        for i in itemlist:
            ii.add(i)
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sht = ctrlr.ActiveSheet
        r = 0
        while True:
            cell = sht.getCellByPosition(index, r)
            s = cell.getString()
            if s:
                ii.add(s)
            else:
                break
            r = r + 1
        itemlist = tuple(sorted(list(ii)))
        return itemlist

    def part_dlg_label_upd(self, evt):
        if evt.Buttons == RIGHT:
            state = evt.Source.getState()
            for l in self.ll:
                l.setState(1 - state)

    def part_dlg_combo_upd(self, source):
        if source not in self.cc:
            return
        i = self.cc.index(source)
        self.ll[i].setState(1)
        devw = self.cc[0]
        if source == devw:
            out(self.ll[0])
            note1lm = self.ll[-2].getModel()
            note1wm = self.cc[-2].getModel()
            note2lm = self.ll[-1].getModel()
            note2wm = self.cc[-1].getModel()
            if devw.Text == 'RESISTOR':
                note1lm.Label = 'Tolerance'
                note1wm.StringItemList = ('1%', '5%')
                note1wm.Text = '1%'
                note2lm.Label = PART_ATTR_LIST[-1]
                note2wm.StringItemList = PART_ATTR_ITEMS[-1]
                note2wm.Text = PART_ATTR_DFLT[-1]
            elif devw.Text == 'CAPACITOR':
                note1lm.Label = 'Voltage'
                note1wm.StringItemList = ('16V', '25V', '50V')
                note1wm.Text = '50V'
                note2lm.Label = 'Dielectric'
                note2wm.StringItemList = ('NPO', 'X5R', 'X7R', 'Y5V')
                note2wm.Text = 'X5R'
            else:
                note1lm.Label = PART_ATTR_LIST[-2]
                note1wm.StringItemList = PART_ATTR_ITEMS[-2]
                note1wm.Text = PART_ATTR_DFLT[-2]
                note2lm.Label = PART_ATTR_LIST[-1]
                note2wm.StringItemList = PART_ATTR_ITEMS[-1]
                note2wm.Text = PART_ATTR_DFLT[-1]

    def execute(self):
        self.dlg.execute()

    def get_active_sheet(self):
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        return ctrlr.ActiveSheet

    def get_selection(self):
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sel = ctrlr.getSelection()
        area = sel.getRangeAddress()
        if area.StartRow != area.EndRow or area.StartColumn != 0 or area.EndColumn != PART_ATTR_LEN - 1:
            return
            #self.msgbox('%d %d %d %d %d' % (area.StartRow, area.EndRow, area.StartColumn, area.EndColumn, len(args)))
        return area.StartRow

    def get_mypart(self):
        part1 = [a.Text for a in self.cc]
        checks = [l.getState() for l in self.ll]
        l = min(len(part1), len(checks))
        for i in range(0, l):
            if not checks[i]:
                part1[i] = ''
        return part1

    def get_part(self, sht, row):
        if not sht:
            sht = self.get_active_sheet()
        p = []
        col = 0
        while True:
            cell = sht.getCellByPosition(col, row)
            col = col + 1
            s = cell.getString()
            if s:
                p.append(s)
            elif len(p) and col <= PART_ATTR_LEN:
                p.append('')
            else:
                break
        return p

    def get_sheet_names(self):
        doc = self.desktop.getCurrentComponent()
        return [doc.Sheets.getByIndex(i).getName() for i in range(0, doc.Sheets.getCount())]

    def set_part(self, sht, row, p, psz):
        if not sht:
            doc = self.desktop.getCurrentComponent()
            ctrlr = doc.CurrentController
            sht = ctrlr.ActiveSheet
        for i in range(0, psz):
            cell = sht.getCellByPosition(i, row)
            if i < len(p):
                cell.setString(p[i])
            else:
                cell.setString('')

    def part_cmp(self, sht, index, part):
        if sht == None:
            sht = self.get_active_sheet()
        part2 = self.get_part(sht, index)
        psz = min(len(part), len(part2))
        for i in range(0, psz):
            if part[i] == '':
                continue
            if part2[i] == '':
                continue
            if part[i] != part2[i]:
                return False
        return True

    def part_find(self):
        sht = self.get_active_sheet()
        part1 = self.get_mypart()
        checks = [l.getState() for l in self.ll]
        part1[PART_ATTR_N] = ''
        index = self.get_selection()
        if index != None:
            part = self.get_part(sht, index)
            if not part:
                index = 0
            elif self.part_cmp(sht, index, part1):
                index += 1
        else:
            index = 0
        index = self.part_find2(sht, part1, index)
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        sel = sht.getCellRangeByPosition(0, index, PART_ATTR_LEN - 1, index)
        ctrlr.select(sel)
        part = self.get_part(sht, index)
        if ''.join(part) == '':
            return
        for i in range(0, len(self.cc)):
            self.cc[i].setText(part[i])
            self.ll[i].setState(checks[i])

    def part_find2(self, sht, part, start=0):
        if sht == None:
            sht = self.get_active_sheet()
        index = start
        while True:
            part2 = self.get_part(sht, index)
            if not part2:
                return index
            if self.part_cmp(sht, index, part):
                return index
            index += 1

    @part_add_del_dec
    def part_add(self):
        sht = self.get_active_sheet()
        index = self.get_selection()
        part = [a.Text for a in self.cc]
        p1 = self.get_part(sht, index)
        if not p1:
            part = [a.Text for a in self.cc]
            for i in range(0, PART_ATTR_LEN):
                cell = sht.getCellByPosition(i, index)
                cell.setString(part[i])
        else:
            cell = sht.getCellByPosition(PART_ATTR_N, index)
            n = int(cell.getString())
            n += int(part[PART_ATTR_N])
            cell.setString('%d' % n)

    @part_add_del_dec
    def part_del(self):
        sht = self.get_active_sheet()
        index = self.get_selection()
        p1 = self.get_part(sht, index)
        if ''.join(p1) == '':
            return
        part = [a.Text for a in self.cc]
        doc = self.desktop.getCurrentComponent()
        ctrlr = doc.CurrentController
        cell = sht.getCellByPosition(PART_ATTR_N, index)
        n = int(cell.getString())
        n -= int(part[PART_ATTR_N])
        if n > 0:
            cell.setString('%d' % n)
        else:
            name1 = ctrlr.ActiveSheet.Name
            index1 = list(doc.getSheets().ElementNames).index(name1)
            area = sht.getRangeAddress()
            arange = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
            arange.Sheet = index1
            arange.StartColumn = area.StartColumn
            arange.StartRow = area.StartRow
            arange.EndColumn = area.EndColumn
            arange.EndRow = area.EndRow
            sht.removeRange(arange, UP)

    def add_label(self):
        shtn = self.get_sheet_names()
        doc = self.desktop.getCurrentComponent()
        self.init_label1()
        labels = 'Labels'
        if labels not in shtn:
            doc.Sheets.insertNewByName(labels, len(shtn))
        self.msgbox('Add label')

    def cell_border(self, cell):
        border = cell.TopBorder
        border.InnerLineWidth = 50
        border.Color = 0
        cell.TopBorder = border
        cell.BottomBorder = border
        cell.LeftBorder = border
        cell.RightBorder = border
        cell.Size.Width *= 2
        cell.Size.Height *= 2
        cell.VertJustify = VJ_CENTER
        return cell

    def init_label1(self):
        shtn = self.get_sheet_names()
        doc = self.desktop.getCurrentComponent()
        label1 = 'Label1'
        if label1 not in shtn:
            doc.Sheets.insertNewByName(label1, len(shtn))
        sht1 = doc.Sheets.getByName(label1)
        #sht1.Rows.Height *= 2
        sht1.Columns.getByIndex(0).Width *= 1.2
        sht1.Columns.getByIndex(1).Width *= 1.5
        for i in range(0, PART_ATTR_N+1):
            sht1.Rows.getByIndex(i).Height *= 2
            cell = sht1.getCellByPosition(0, i)
            cell = self.cell_border(cell)
            cell.setString(PART_ATTR_LIST[i])
            cell = sht1.getCellByPosition(1, i)
            cell = self.cell_border(cell)
            cell.HoriJustify = HJ_CENTER
            cell.setString(self.cc[i].Text)
        cell = sht1.getCellByPosition(0, PART_ATTR_N+1)
        cell = self.cell_border(cell)
        cell = sht1.getCellByPosition(1, PART_ATTR_N+1)
        cell = self.cell_border(cell)

    def AddBorder(self):
        oDoc = XSCRIPTCONTEXT.getDocument()
        oSheets = oDoc.getSheets()
        oSheet =  oSheets.getByName('Label1')
        oCell = oSheet.getCellByPosition(3,4)
        Border = oCell.TopBorder
        Border.OuterLineWidth = 100
        Border.Color = 700
        oCell.TopBorder = Border
        oCell.BottomBorder = Border
        oCell.LeftBorder = Border
        oCell.RightBorder = Border

    def msgbox(self, msg='message', title='Message', btntype=1):
        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow()
        toolkit = window.getToolkit()
        mb = toolkit.createMessageBox(window, MESSAGEBOX, btntype, title, msg)
        mb.execute()

    def createUnoService(self, name):
        return self.ctx.getServiceManager().createInstance(name)

    def center(self, model):
        parentpos = self.desktop.getCurrentComponent().CurrentController.Frame.ContainerWindow.PosSize
        #self.msgbox('%s %s' % (parentpos.Width, parentpos.Height))
        model.PositionX = parentpos.Width/8
        model.PositionY = parentpos.Height/8
        #model.PositionX = 1
        #model.PositionY = 1

def myparts(*args):
    MyParts(XSCRIPTCONTEXT.getComponentContext()).execute()
