
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.PushButtonType import STANDARD, OK, CANCEL
from com.sun.star.awt import XActionListener, XTextListener, XMouseListener, XMenuListener, Rectangle
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.awt.MouseButton import RIGHT
from com.sun.star.awt.PopupMenuDirection import EXECUTE_DEFAULT
from com.sun.star.table import CellRangeAddress
from com.sun.star.sheet.CellDeleteMode import UP
from com.sun.star.table.CellVertJustify import CENTER as VJ_CENTER
from com.sun.star.table.CellHoriJustify import CENTER as HJ_CENTER
from uno import createUnoStruct
from unohelper import Base

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

class ButtonListener(Base, XActionListener):
    def __init__(self, cb):
        self.cb = cb
    def actionPerformed(self, evt):
        self.cb()

class LabelListener(Base, XMouseListener):
    def __init__(self, cb):
        self.cb = cb
    def mouseReleased(self, evt):
        self.cb(evt)

class ComboboxListener(Base, XTextListener):
    def __init__(self, cb):
        self.cb = cb
    def textChanged(self, evt):
        self.cb(evt.Source)

class MenuListener(Base, XMenuListener):
    def __init__(self, cb):
        self.cb = cb
    def itemSelected(self, evt):
        self.cb(evt.MenuId)

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
            mp.part_find_cb()
            mp.cc[PART_ATTR_N].setText(pn)
            return
        sht = mp.get_active_sheet()
        if not mp.part_cmp(sht, index, part):
            mp.part_find_cb()
            mp.cc[PART_ATTR_N].setText(pn)
            return
        p1 = mp.get_part(sht, index)
        if ''.join(p1) == '' and index > 0:
            if mp.part_find(sht, part, 0) != index:
                mp.part_find_cb()
                mp.cc[PART_ATTR_N].setText(pn)
                return
        return f(*args, **kwargs)
    return tmp

class MyParts(object):
    def __init__(self, ctx):
        self.ctx = ctx
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.doc = self.desktop.getCurrentComponent()
        if self.get_active_sheet().Name == 'Labels':
            return
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
        if not self.init_data(dlg):
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

    def dropdown_cb(self):
        self.msgbox('dropdown')

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
        listener = ButtonListener(self.part_find_cb)
        btn.addActionListener(listener)
        btn = self.init_button(dlg, model.Width/3, 0, 'Add')
        listener = ButtonListener(self.part_add_cb)
        btn.addActionListener(listener)
        btn = self.init_button(dlg, 2*model.Width/3, 0, 'Del')
        listener = ButtonListener(self.part_del_cb)
        btn.addActionListener(listener)
        self.move_to_btn = self.init_button(dlg, 0, 1, 'Move to...')
        listener = ButtonListener(self.part_move_cb)
        self.move_to_btn.addActionListener(listener)
        self.add_label_btn = self.init_button(dlg, model.Width/3, 1, 'Add label')
        listener = ButtonListener(self.add_label_cb)
        self.add_label_btn.addActionListener(listener)
        btn = self.init_button(dlg, 2*model.Width/3, 1, 'Close')
        listener = ButtonListener(self.close_cb)
        btn.addActionListener(listener)
        return dlg

    def init_data(self, dlg):
        sht = self.get_active_sheet()
        index = self.get_selection()
        if index == None:
            area = self.doc.CurrentController.getSelection().getRangeAddress()
            index = area.StartRow
            cell = sht.getCellByPosition(0, index)
            if index == 0:
                self.set_selection(index)
                if not self.get_part(sht, index):
                    return
            elif self.get_part(sht, index-1):
                self.set_selection(index)
                if not self.get_part(sht, index):
                    return
        part = self.get_part(sht, index)
        l = min(len(part), PART_ATTR_LEN)
        for i in range(0, l):
            self.cc[i].setText(part[i])
        return True

    def get_combo_itemlist(self, index):
        itemlist = PART_ATTR_ITEMS[index]
        '''
        if index >= PART_ATTR_N:
            return itemlist
        ii = set()
        for i in itemlist:
            ii.add(i)
        sht = self.get_active_sheet()
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
        '''
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
        valw = self.cc[1]
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
        fpwm = self.cc[2].getModel()
        fps = PART_ATTR_ITEMS[2]
        if devw.Text not in PART_ATTR_ITEMS[0]:
            fps1 = self.get_footprints(valw.Text)
            if fps:
                fps = fps1 + list(fps)
                fps = tuple(fps)
            fpwm.StringItemList = fps
        else:
            fpwm.StringItemList = fps

    def get_footprints(self, partn):
        shtn = self.get_sheet_names(['Labels'])
        ret = []
        for name in shtn:
            sht1 = self.doc.Sheets.getByName(name)
            i = 0
            while True:
                part = self.get_part(sht1, i)
                if not part:
                    break
                i += 1
                if part[1] == partn:
                    ret.append(part[2])
        return ret

    def execute(self):
        if hasattr(self, 'dlg'):
            self.dlg.execute()

    def get_active_sheet(self):
        return self.doc.CurrentController.ActiveSheet

    def get_selection(self):
        sel = self.doc.CurrentController.getSelection()
        area = sel.getRangeAddress()
        if area.StartRow != area.EndRow or area.StartColumn != 0 or area.EndColumn != PART_ATTR_LEN - 1:
            return
            #self.msgbox('%d %d %d %d %d' % (area.StartRow, area.EndRow, area.StartColumn, area.EndColumn, len(args)))
        return area.StartRow

    def set_selection(self, index):
        sht = self.get_active_sheet()
        sel = sht.getCellRangeByPosition(0, index, PART_ATTR_LEN - 1, index)
        self.doc.CurrentController.select(sel)

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

    def get_sheet_names(self, skip=[]):
        names = [self.doc.Sheets.getByIndex(i).getName() for i in range(0, self.doc.Sheets.getCount())]
        for s in skip:
            if s in names:
                names.remove(s)
        return names

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

    def part_find_cb(self):
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
        index = self.part_find(sht, part1, index)
        self.set_selection(index)
        part = self.get_part(sht, index)
        if not part:
            return
        for i in range(0, len(self.cc)):
            self.cc[i].setText(part[i])
            self.ll[i].setState(checks[i])

    def part_find(self, sht, part, start=0):
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

    def part_eq(self):
        part = [a.Text for a in self.cc]
        part[PART_ATTR_N] = ''
        index = self.get_selection()
        if index == None:
            return False
        sht = self.get_active_sheet()
        part1 = self.get_part(sht, index)
        if not part1:
            return False
        return self.part_cmp(sht, index, part)

    @part_add_del_dec
    def part_add_cb(self):
        self.part_add(self.get_active_sheet(), self.get_selection())

    def part_add(self, sht, index=None):
        part = [a.Text for a in self.cc]
        if index == None:
            qty = part[PART_ATTR_N]
            part[PART_ATTR_N] = ''
            index = self.part_find(sht, part, 0)
            part[PART_ATTR_N] = qty
        part1 = self.get_part(sht, index)
        if not part1:
            for i in range(0, PART_ATTR_LEN):
                cell = sht.getCellByPosition(i, index)
                cell.setString(part[i])
        else:
            cell = sht.getCellByPosition(PART_ATTR_N, index)
            qty = int(cell.getString())
            qty += int(part[PART_ATTR_N])
            cell.setString('%d' % qty)

    @part_add_del_dec
    def part_del_cb(self):
        sht = self.get_active_sheet()
        index = self.get_selection()
        p1 = self.get_part(sht, index)
        if not p1:
            return
        part = [a.Text for a in self.cc]
        self.part_del(sht, index, int(part[PART_ATTR_N]))

    def part_del(self, sht, index, qty):
        if sht == None:
            sht = self.get_active_sheet()
        part1 = self.get_part(sht, index)
        qty1 = int(part1[PART_ATTR_N])
        if qty < qty1:
            cell = sht.getCellByPosition(PART_ATTR_N, index)
            cell.setString('%d' % (qty1 - qty))
        else:
            index1 = self.get_sheet_names().index(sht.Name)
            sel = self.doc.CurrentController.getSelection()
            area = sel.getRangeAddress()
            arange = createUnoStruct('com.sun.star.table.CellRangeAddress')
            arange.Sheet = index1
            arange.StartColumn = area.StartColumn
            arange.StartRow = area.StartRow
            arange.EndColumn = area.EndColumn
            arange.EndRow = area.EndRow
            sht.removeRange(arange, UP)
            #self.msgbox('%d %d %d %d' % (area.StartColumn, area.StartRow, area.EndColumn, area.EndRow))

    def part_move_cb(self):
        if not self.part_eq():
            return
        sht = self.get_active_sheet()
        shtn = self.get_sheet_names(['Labels', sht.Name])
        if not shtn:
            return
        index = self.get_selection()
        if index == None:
            return
        menu = self.createUnoService('com.sun.star.awt.PopupMenu')
        for i in range(0, len(shtn)):
            menu.insertItem(i+1, shtn[i], 0, i)
        listener = MenuListener(self.part_move)
        menu.addMenuListener(listener)
        rect = Rectangle()
        btnpos = self.move_to_btn.PosSize
        rect.X = btnpos.X + btnpos.Width/2
        rect.Y = btnpos.Y + btnpos.Height/2
        menu.execute(self.dlg.Peer, rect, EXECUTE_DEFAULT)

    def part_move(self, menuid):
        sht = self.get_active_sheet()
        shtn = self.get_sheet_names(['Labels', sht.Name])
        sht1 = self.doc.Sheets.getByName(shtn[menuid-1])
        index = self.get_selection()
        part = self.get_part(sht, index)
        qty = int(part[PART_ATTR_N])
        part1 = [a.Text for a in self.cc]
        qty1 = int(part1[PART_ATTR_N])
        self.part_del(sht, index, qty1)
        self.part_add(sht1)

    def add_label_cb(self):
        self.init_labels()
        index = self.add_label()
        if type(index) == int:
            self.add_label_btn.Label = 'Add label %d' % (index + 1)

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

    def init_labels(self):
        shtn = self.get_sheet_names()
        labels = 'Labels'
        if labels not in shtn:
            self.doc.Sheets.insertNewByName(labels, len(shtn))
        else:
            return

    def format_cell_text(self, i, col):
        rc = self.cc[0].Text in ['RESISTOR', 'CAPACITOR']
        if col == 0:
            if i == 1 and not rc:
                return 'Part #'
            if i == 2:
                return 'Package'
            if i == 3:
                return 'Qty'
            return PART_ATTR_LIST[i]
        if col == 1:
            if i == 1:
                text = self.cc[1].Text
                for i in [PART_ATTR_N + 1, PART_ATTR_N + 2]:
                    texti = self.cc[i].Text
                    if texti:
                        text += '/' + texti
                return text
            return self.cc[i].Text

    def add_label(self, index=None):
        sht1 = self.doc.Sheets.getByName('Labels')
        for i in range(0, 21):
            row = 6*int(i / 3)
            col = 3*int(i % 3)
            cell = sht1.getCellByPosition(col, row)
            if not cell.getString():
                index = i
                break
        if index == None:
            return False
        k1 = 1
        k2 = 1.5
        if index == 0:
            sht1.Columns.getByIndex(0).Width *= k1
            sht1.Columns.getByIndex(1).Width *= k2
        if index == 1:
            sht1.Columns.getByIndex(2).Width /= 10
            sht1.Columns.getByIndex(3).Width *= k1
            sht1.Columns.getByIndex(4).Width *= k2
        if index == 2:
            sht1.Columns.getByIndex(5).Width /= 10
            sht1.Columns.getByIndex(6).Width *= k1
            sht1.Columns.getByIndex(7).Width *= k2
        row = 6*int(index / 3)
        col = 3*int(index % 3)
        rr = False
        #self.msgbox(str(index) + str(row) + str(col))
        if sht1.Rows.getByIndex(row).Height == sht1.Rows.getByIndex(row + PART_ATTR_N + 1).Height:
            rr = True
        for i in range(0, PART_ATTR_N+1):
            if col == 0 and rr:
                sht1.Rows.getByIndex(row + i).Height *= 1.5
                if i == 0 and row:
                    sht1.Rows.getByIndex(row - 1).Height /= 2
            cell = sht1.getCellByPosition(col, row + i)
            cell = self.cell_border(cell)
            cell.setString(self.format_cell_text(i, 0))
            cell = sht1.getCellByPosition(col + 1, row + i)
            cell = self.cell_border(cell)
            cell.HoriJustify = HJ_CENTER
            cell.setString(self.format_cell_text(i, 1))
        cell = sht1.getCellByPosition(col, row+PART_ATTR_N+1)
        cell = self.cell_border(cell)
        cell.CharHeight = 8
        if row or col:
            cell.Formula = "=A5"
        cell = sht1.getCellByPosition(col + 1, row+PART_ATTR_N+1)
        cell = self.cell_border(cell)
        cell.CharHeight = 8
        if row or col:
            cell.Formula = "=B5"
        cell.HoriJustify = HJ_CENTER
        return index

    def close_cb(self):
        self.dlg.endExecute()

    def msgbox(self, msg='message', title='Message', btntype=1):
        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow()
        toolkit = window.getToolkit()
        mb = toolkit.createMessageBox(window, MESSAGEBOX, btntype, title, msg)
        mb.execute()

    def createUnoService(self, name):
        return self.ctx.getServiceManager().createInstance(name)

    def center(self, model):
        parentpos = self.doc.CurrentController.Frame.ContainerWindow.PosSize
        #self.msgbox('%s %s' % (parentpos.Width, parentpos.Height))
        model.PositionX = parentpos.Width/8
        model.PositionY = parentpos.Height/8
        #model.PositionX = 1
        #model.PositionY = 1

def myparts(*args):
    MyParts(XSCRIPTCONTEXT.getComponentContext()).execute()

