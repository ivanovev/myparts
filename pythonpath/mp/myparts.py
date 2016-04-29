
from com.sun.star.awt import XActionListener, XKeyListener, XKeyHandler, XMouseListener, XMenuListener, XTextListener, Rectangle
from com.sun.star.awt.MouseButton import RIGHT
from com.sun.star.awt.PopupMenuDirection import EXECUTE_DEFAULT
from com.sun.star.awt.PosSize import POS
from com.sun.star.awt.PushButtonType import STANDARD
from com.sun.star.sheet.CellDeleteMode import UP
from com.sun.star.table.CellVertJustify import CENTER as VJ_CENTER
from com.sun.star.table.CellHoriJustify import CENTER as HJ_CENTER
from uno import createUnoStruct
from unohelper import Base

PART_ATTR_ITEMS = [
    ('RESISTOR', 'CAPACITOR', 'INDUCTOR', 'CONNECTOR', 'LED', 'XTAL'),
    ('1p', '1n', '1u', '1m', '1', '1k', '1M'),
    ('0402', '0603', '0805', '1206'),
    ('0', '1', '2', '3', '4', '5'),
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

class UpDownListener(Base, XKeyListener):
    def __init__(self, cb):
        self.cb = cb
    def keyPressed(self, evt):
        self.cb(evt)

class ComboboxListener(Base, XMouseListener):
    def __init__(self, cb):
        self.cb = cb
    def mousePressed(self, evt):
        if evt.Buttons == RIGHT:
            self.cb(evt.Source)

class ComboboxListener2(Base, XTextListener):
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
        mp1 = args[0]
        if not mp1.ll[ms.PART_ATTR_N].getState():
            return
        part = [a.Text for a in mp1.cc]
        pn = part[ms.PART_ATTR_N]
        part[ms.PART_ATTR_N] = ''
        index = mp1.get_selection()
        if index == None:
            mp1.part_find_cb()
            mp1.cc[ms.PART_ATTR_N].setText(pn)
            return
        sht = mp1.get_sheet()
        if not mp1.part_cmp(sht, index, part):
            mp1.part_find_cb()
            mp1.cc[ms.PART_ATTR_N].setText(pn)
            return
        p1 = mp1.get_part(sht, index)
        if ''.join(p1) == '' and index > 0:
            if mp1.part_find(sht, part, 0) != index:
                mp1.part_find_cb()
                mp1.cc[ms.PART_ATTR_N].setText(pn)
                return
        return f(*args, **kwargs)
    return tmp

class MyParts(object):
    def __init__(self, ctx, title='MyParts'):
        self.ll = []
        self.cc = []
        if not hasattr(self, 'x'):
            self.x = 100
            self.y = 100
            self.w = 120
            self.h = 150
        ms.MySearch.__init__(self, ctx, title=title)
        if title != 'MyParts':
            return
        if self.get_sheet().Name == 'Labels':
            return
        if not self.init_data():
            self.part_dlg_combo_upd_cb(self.cc[0])

    def init_dlg(self):
        model = self.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', self.ctx)
        model.Width = self.w
        model.Height = self.h
        model.Title = self.title
        ms.out(model, 'model')

        self.udlistener = UpDownListener(self.up_down_cb)

        dlg = self.createUnoService('com.sun.star.awt.UnoControlDialog')
        dlg.setPosSize(self.x, self.y, 0, 0, POS)
        dlg.setModel(model)
        toolkit = self.createUnoService('com.sun.star.awt.Toolkit')
        dlg.createPeer(toolkit, None)

        self.init_rows(dlg)
        self.init_buttons(dlg)
        ms.out(dlg, 'dlg')

        listener = ms.PositionListener(self.move_cb, self.resize_cb)
        dlg.addWindowListener(listener)

        listener = ms.DisposeListener(self.config_write)
        dlg.addEventListener(listener)

        dlg.addKeyListener(self.udlistener)

        #self.handler = UpDownHandler(self.up_down_cb)
        #self.ctrl.addKeyHandler(self.handler)
        #ms.out(dlg, 'dlg')

        self.dlg = dlg

    def up_down_cb(self, evt):
        index = self.get_selection()
        mod = evt.value.Modifiers
        code = evt.value.KeyCode
        if mod == 2 and code == 0x20F:
            if index > 0:
                index -= 1
        if mod == 2 and code == 0x20D:
            p = self.get_part(None, index)
            if p:
                index += 1
            else:
                index = None
        if index != None:
            self.set_selection(index)
        p = self.get_part(None, index)
        if p:
            self.set_mypart(p)

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
        for i in range(0, ms.PART_ATTR_LEN):
            label = ms.PART_ATTR_LIST[i]
            items = PART_ATTR_ITEMS[i]
            l,w = self.init_row(dlg, i*model.Height/(ms.PART_ATTR_LEN + 2), label, items)
            l.addKeyListener(self.udlistener)
            w.addKeyListener(self.udlistener)
            l.setState(1)
            w.Text = PART_ATTR_DFLT[i]
            listener = LabelListener(self.part_dlg_label_upd_cb)
            l.addMouseListener(listener)
            listener = ComboboxListener(self.part_dlg_combo_upd_cb)
            w.addMouseListener(listener)
            if i == 0:
                listener = ComboboxListener2(self.part_dlg_combo_upd_cb2)
                w.addTextListener(listener)
                ms.out(l, 'l')
                ms.out(w, 'w')

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
        btnmodel.FocusOnClick = 0
        model.insertByName(label, btnmodel)
        return dlg.getControl(label)

    def init_buttons(self, dlg):
        model = dlg.getModel()
        if 0:
            btn = self.init_button(dlg, 0, 0, 'Find')
            listener = ButtonListener(self.part_find_cb)
        else:
            btn = self.init_button(dlg, 0, 0, 'Search')
            listener = ButtonListener(self.part_search_cb)
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

    def init_data(self):
        sht = self.get_sheet()
        index = self.get_selection()
        if index == None:
            area = self.doc.CurrentController.getSelection().getRangeAddress()
            index = area.StartRow
            cell = sht.getCellByPosition(0, index)
            if not cell.getString():
                index = self.get_part_count(sht)
            self.set_selection(index)
        part1 = self.get_part(sht, index)
        if not part1:
            return False
        self.set_mypart(part1)
        return True

    def part_dlg_label_upd_cb(self, evt):
        if evt.Buttons == RIGHT:
            state = evt.Source.getState()
            for l in self.ll:
                l.setState(1 - state)

    def part_dlg_combo_upd_cb(self, source):
        if source not in self.cc:
            return
        i = self.cc.index(source)
        devw = self.cc[0]
        valw = self.cc[1]
        if source == devw:
            note1lm = self.ll[-2].getModel()
            note1wm = self.cc[-2].getModel()
            note2lm = self.ll[-1].getModel()
            note2wm = self.cc[-1].getModel()
            if devw.Text == 'RESISTOR':
                note1lm.Label = 'Tolerance'
                note1wm.StringItemList = ('1%', '5%')
                note1wm.Text = '1%'
                note2lm.Label = ms.PART_ATTR_LIST[-1]
                note2wm.StringItemList = PART_ATTR_ITEMS[-1]
                note2wm.Text = PART_ATTR_DFLT[-1]
                if valw.Text == '1u':
                    valw.Text = '1k'
            elif devw.Text == 'CAPACITOR':
                note1lm.Label = 'Voltage'
                note1wm.StringItemList = ('16V', '25V', '50V')
                note1wm.Text = '50V'
                note2lm.Label = 'Dielectric'
                note2wm.StringItemList = ('NPO', 'X5R', 'X7R', 'Y5V')
                note2wm.Text = 'NPO'
                if valw.Text == '1k':
                    valw.Text = '1u'
            else:
                note1lm.Label = ms.PART_ATTR_LIST[-2]
                note1wm.StringItemList = PART_ATTR_ITEMS[-2]
                note1wm.Text = PART_ATTR_DFLT[-2]
                note2lm.Label = ms.PART_ATTR_LIST[-1]
                note2wm.StringItemList = PART_ATTR_ITEMS[-1]
                note2wm.Text = PART_ATTR_DFLT[-1]
        fpw = self.cc[2]
        fpwm = fpw.getModel()
        fps = PART_ATTR_ITEMS[2]
        if source == fpw:
            if fps:
                fps = list(fps)
                fps1 = self.get_footprints(valw.Text, fps)
                fps = fps1 + fps
                fps = tuple(fps)
            fpwm.StringItemList = fps
        else:
            fpwm.StringItemList = fps

    def part_dlg_combo_upd_cb2(self, source):
        devw = self.cc[0]
        if devw.Text in PART_ATTR_ITEMS[0]:
            self.part_dlg_combo_upd_cb(source)

    def selection_cb(self):
        index = self.get_selection()
        if index == None:
            return
        part1 = self.get_part(None, index)
        self.set_mypart(part1)

    def part_find_cb(self):
        sht = self.get_sheet()
        part1 = self.get_mypart()
        checks = [l.getState() for l in self.ll]
        part1[ms.PART_ATTR_N] = ''
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

    def part_search_cb(self):
        part1 = self.get_mypart()
        if hasattr(self, 'search'):
            self.search.close_cb()
        self.search = ms.MySearch(self.ctx, part1, self.selection_cb)
        self.search.execute()

    def close_cb(self):
        if hasattr(self, 'search'):
            self.search.close_cb()
        self.dlg.endExecute()

    def get_footprints(self, partn, skip=[]):
        shtn = self.get_sheet_names(['Labels'])
        ret = []
        for name in shtn:
            sht1 = self.doc.getSheets().getByName(name)
            i = 0
            while True:
                part = self.get_part(sht1, i)
                if not part:
                    break
                i += 1
                if part[1] == partn and part[2] not in ret and part[2] not in skip:
                    ret.append(part[2])
        return ret

    def execute(self):
        if hasattr(self, 'dlg'):
            self.dlg.execute()

    def get_mypart(self):
        part1 = [a.Text for a in self.cc]
        checks = [l.getState() for l in self.ll]
        l = min(len(part1), len(checks))
        for i in range(0, l):
            if not checks[i]:
                part1[i] = ''
        return part1

    def set_mypart(self, part1):
        for i in range(0, min(ms.PART_ATTR_LEN, len(part1))):
            self.cc[i].setText(part1[i])

    def part_eq(self):
        part = [a.Text for a in self.cc]
        part[ms.PART_ATTR_N] = ''
        index = self.get_selection()
        if index == None:
            return False
        sht = self.get_sheet()
        part1 = self.get_part(sht, index)
        if not part1:
            return False
        return self.part_cmp(sht, index, part)

    @part_add_del_dec
    def part_add_cb(self):
        self.part_add(self.get_sheet(), self.get_selection())

    def part_add(self, sht, index=None):
        part = [a.Text for a in self.cc]
        if index == None:
            qty = part[ms.PART_ATTR_N]
            part[ms.PART_ATTR_N] = ''
            index = self.part_find(sht, part, 0)
            part[ms.PART_ATTR_N] = qty
        part1 = self.get_part(sht, index)
        if not part1:
            for i in range(0, ms.PART_ATTR_LEN):
                cell = sht.getCellByPosition(i, index)
                cell.setString(part[i])
        else:
            cell = sht.getCellByPosition(ms.PART_ATTR_N, index)
            qty = int(cell.getString())
            qty += int(part[ms.PART_ATTR_N])
            cell.setString('%d' % qty)

    @part_add_del_dec
    def part_del_cb(self):
        sht = self.get_sheet()
        index = self.get_selection()
        p1 = self.get_part(sht, index)
        if not p1:
            return
        part = [a.Text for a in self.cc]
        self.part_del(sht, index, int(part[ms.PART_ATTR_N]))
        if not self.init_data():
            self.part_dlg_combo_upd_cb(self.cc[0])

    def part_del(self, sht, index, qty):
        if sht == None:
            sht = self.get_sheet()
        part1 = self.get_part(sht, index)
        qty1 = int(part1[ms.PART_ATTR_N])
        if qty < qty1:
            cell = sht.getCellByPosition(ms.PART_ATTR_N, index)
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

    def part_move_cb(self):
        if not self.part_eq():
            return
        sht = self.get_sheet()
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
        sht = self.get_sheet()
        shtn = self.get_sheet_names(['Labels', sht.Name])
        sht1 = self.doc.Sheets.getByName(shtn[menuid-1])
        index = self.get_selection()
        part = self.get_part(sht, index)
        qty = int(part[ms.PART_ATTR_N])
        part1 = [a.Text for a in self.cc]
        qty1 = int(part1[ms.PART_ATTR_N])
        self.part_del(sht, index, qty1)
        self.part_add(sht1)
        if not self.init_data():
            self.part_dlg_combo_upd_cb(self.cc[0])

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
        labels = 'Labels'
        shtn = self.get_sheet_names(skip=[])
        if labels not in shtn:
            self.doc.getSheets().insertNewByName(labels, len(shtn))
        else:
            return

    def format_label_text(self, i, col):
        p = self.cc[0].Text
        rc = p in ['RESISTOR', 'CAPACITOR']
        x = p == 'XTAL'
        if col == 0:
            if i == 1:
                if x:
                    return 'Frequency'
                if not rc:
                    return 'Part #'
            if i == 2:
                return 'Package'
            if i == 3:
                return 'Notes'
            if i == 4:
                return 'Qty'
            return ms.PART_ATTR_LIST[i]
        if col == 1:
            text = self.cc[i].Text
            if i == 3:
                text = self.cc[ms.PART_ATTR_N + 1].Text
                note2 = self.cc[ms.PART_ATTR_N + 2].Text
                if note2:
                    text += '   ' + note2
            if i == 4:
                text = self.cc[ms.PART_ATTR_N].Text
            return text

    def add_label(self, index=None):
        sht1 = self.get_sheet('Labels')
        index = None
        for i in range(0, 21):
            row = 7*int(i / 3)
            col = 3*int(i % 3)
            cell = sht1.getCellByPosition(col, row)
            if not cell.getString():
                index = i
                break
        if index == None:
            return False
        k1 = 1.05
        k2 = 1.4
        if index == 0:
            sht1.Columns.getByIndex(0).Width *= k1
            sht1.Columns.getByIndex(1).Width *= k2
        if index == 1:
            sht1.Columns.getByIndex(2).Width /= 12
            sht1.Columns.getByIndex(3).Width *= k1
            sht1.Columns.getByIndex(4).Width *= k2
        if index == 2:
            sht1.Columns.getByIndex(5).Width /= 12
            sht1.Columns.getByIndex(6).Width *= k1
            sht1.Columns.getByIndex(7).Width *= k2
        row = 7*int(index / 3)
        col = 3*int(index % 3)
        rr = False
        if sht1.Rows.getByIndex(row).Height == sht1.Rows.getByIndex(row + ms.PART_ATTR_N + 2).Height:
            rr = True
        for i in range(0, ms.PART_ATTR_N+2):
            if col == 0 and rr:
                sht1.Rows.getByIndex(row + i).Height *= 1.25
                if i == 0 and row:
                    sht1.Rows.getByIndex(row - 1).Height /= 5
            cell = sht1.getCellByPosition(col, row + i)
            cell = self.cell_border(cell)
            cell.setString(self.format_label_text(i, 0))
            cell = sht1.getCellByPosition(col + 1, row + i)
            cell = self.cell_border(cell)
            cell.HoriJustify = HJ_CENTER
            cell.setString(self.format_label_text(i, 1))
        cell = sht1.getCellByPosition(col, row+ms.PART_ATTR_N+2)
        cell = self.cell_border(cell)
        cell.CharHeight = 8
        if row or col:
            cell.Formula = "=A6"
        cell = sht1.getCellByPosition(col + 1, row+ms.PART_ATTR_N+2)
        cell = self.cell_border(cell)
        cell.CharHeight = 8
        if row or col:
            cell.Formula = "=B6"
        cell.HoriJustify = HJ_CENTER
        return index

def myparts(*args):
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
    MP = type('MyParts', (MyParts, ms.MySearch), {})
    MP(XSCRIPTCONTEXT.getComponentContext()).execute()

