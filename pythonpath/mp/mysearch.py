
from com.sun.star.awt import XActionListener, XWindowListener
from com.sun.star.awt import Rectangle, WindowDescriptor
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.awt.WindowAttribute import BORDER, SHOW, SIZEABLE, MOVEABLE, CLOSEABLE
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.lang import XEventListener
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view.SelectionType import SINGLE
from com.sun.star.beans import PropertyValue
from uno import createUnoStruct
from unohelper import Base
from collections import OrderedDict as OD

PART_ATTR_LIST = ['Device', 'Value', 'Footprint', 'Quantity', 'Note1', 'Note2']
PART_ATTR_LEN = len(PART_ATTR_LIST)

def out(obj, name='out'):
    f = open('/tmp/'+name, 'w')
    for i in dir(obj):
        f.write(str(i))
        f.write('\n\n\r')
    f.close()

class ButtonListener(Base, XActionListener):
    def __init__(self, cb):
        self.cb = cb
    def actionPerformed(self, evt):
        self.cb()

class PositionListener(Base, XWindowListener):
    def __init__(self, cbm, cbr):
        self.cbm = cbm
        self.cbr = cbr
    def windowMoved(self, evt):
        self.cbm()
    def windowResized(self, evt):
        self.cbr()

class DisposeListener(Base, XEventListener):
    def __init__(self, cb):
        self.cb = cb
    def disposing(self, evt):
        self.cb()

class NodeListener(Base, XSelectionChangeListener):
    def __init__(self, cb):
        self.cb = cb
    def selectionChanged(self, evt):
        self.cb()

class MySearch(object):
    def __init__(self, ctx, part1=[], selection_cb=None, title='MySearch'):
        self.ctx = ctx
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.doc = self.desktop.getCurrentComponent()
        self.ctrl = self.doc.getCurrentController()
        self.smgr = self.ctx.getServiceManager()
        self.parent = self.desktop.getCurrentFrame().getContainerWindow()
        self.toolkit = self.parent.getToolkit()
        if self.get_sheet().Name == 'Labels':
            return
        if not hasattr(self, 'x'):
            self.x = 100
            self.y = 100
            self.w = 200
            self.h = 300
        pv = PropertyValue()
        pv.Name = 'nodepath'
        pv.Value = 'vnd.myparts.settings/%s' % title.lower()
        self.pv = pv
        self.config_read()
        self.title = title
        self.init_dlg()
        if title != 'MySearch':
            return
        self.result = OD()
        if part1:
            self.part_search(None, part1)
        self.display_results()
        self.selection_cb = selection_cb

    def new_search(self, ctx, part1=[]):
        return MySearch(ctx, part1)

    def init_dlg(self):
        rect = createUnoStruct("com.sun.star.awt.Rectangle")
        rect = Rectangle(self.x, self.y, self.w, self.h)

        desc = createUnoStruct("com.sun.star.awt.WindowDescriptor")
        frame = self.smgr.createInstanceWithContext('com.sun.star.frame.Frame', self.ctx)
        wnd = self.toolkit.createWindow(WindowDescriptor(SIMPLE, 'floatingwindow', self.parent, -1, rect, BORDER + SHOW + SIZEABLE + MOVEABLE + CLOSEABLE))
        frame.initialize(wnd)
        self.frame = frame

        cont = self.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlContainer', self.ctx)
        cont_model = self.smgr.createInstanceWithContext('com.sun.star.awt.UnoControlContainerModel', self.ctx)
        cont_model.BackgroundColor = 0xCCCCCC
        cont.setModel(cont_model)
        cont.createPeer(self.toolkit, wnd)
        cont.setPosSize(0, 0, self.w, self.h, POSSIZE)
        frame.setComponent(cont,None)
        self.cont = cont

        tree = self.smgr.createInstanceWithContext('com.sun.star.awt.tree.TreeControl', self.ctx)
        self.tree_model = self.smgr.createInstanceWithContext('com.sun.star.awt.tree.TreeControlModel', self.ctx)
        tree.setPosSize(0, 0, self.w, self.h, POSSIZE)
        tree.setModel(self.tree_model)
        cont.addControl('tree', tree)
        self.tree = tree

        self.tree_data = self.createUnoService("com.sun.star.awt.tree.MutableTreeDataModel")
        self.tree_model.DataModel = self.tree_data

        listener = PositionListener(self.move_cb, self.resize_cb)
        wnd.addWindowListener(listener)

        listener = DisposeListener(self.config_write)
        wnd.addEventListener(listener)

        listener = NodeListener(self.tree_selection_cb)
        self.tree.addSelectionChangeListener(listener)

        self.wnd = wnd

    def createUnoService(self, name):
        return self.smgr.createInstance(name)

    def close_cb(self):
        if hasattr(self, 'wnd'):
            self.config_write()
            self.wnd.setVisible(False)

    def move_cb(self):
        if hasattr(self, 'wnd'):
            self.x = self.wnd.PosSize.X
            self.y = self.wnd.PosSize.Y
        elif hasattr(self, 'dlg'):
            self.x = self.dlg.PosSize.X
            self.y = self.dlg.PosSize.Y

    def resize_cb(self):
        self.w = self.wnd.PosSize.Width
        self.h = self.wnd.PosSize.Height
        self.cont.setPosSize(0, 0, self.w, self.h, POSSIZE)
        self.tree.setPosSize(0, 0, self.w, self.h, POSSIZE)

    def tree_selection_cb(self):
        sel = self.tree.getSelection()
        if not sel:
            return
        selp = sel.getParent()
        name = sel.getDisplayValue()
        if selp != self.root_node:
            name = selp.getDisplayValue()
        sht1 = self.get_sheet(name)
        self.ctrl.setActiveSheet(sht1)
        if selp == self.root_node:
            index = self.get_part_count(sht1)
        else:
            part = sel.getDisplayValue()
            pp = part.split()
            index = int(pp.pop(0)) - 1
        self.set_selection(index)
        self.selection_cb()

    def part_cmp(self, sht, index, part):
        if sht == None:
            sht = self.get_sheet()
        if type(index) == int:
            part2 = self.get_part(sht, index)
        elif type(index) == list:
            part2 = index
        else:
            return False
        psz = min(len(part), len(part2))
        for i in range(0, psz):
            if part[i] == '':
                continue
            if part2[i] == '':
                continue
            if part[i] != part2[i]:
                return False
        return True

    def part_search(self, shtn, part):
        if shtn:
            names = shtn
        else:
            names = self.get_sheet_names()
        for name in names:
            sht1 = self.doc.Sheets.getByName(name)
            index = 0
            while True:
                index = self.part_find(sht1, part, index)
                part1 = self.get_part(sht1, index)
                if not part1:
                    break
                if name not in self.result:
                    self.result[name] = OD()
                self.result[name][index] = part1
                index += 1

    def part_find(self, sht, part, start=0):
        if sht == None:
            sht = self.get_sheet()
        index = start
        while True:
            part2 = self.get_part(sht, index)
            if not part2:
                return index
            if self.part_cmp(sht, index, part):
                return index
            index += 1

    def get_selection(self):
        sel = self.ctrl.getSelection()
        area = sel.getRangeAddress()
        if area.StartRow != area.EndRow or area.StartColumn != 0 or area.EndColumn != PART_ATTR_LEN - 1:
            return
        return area.StartRow

    def set_selection(self, index):
        sht = self.get_sheet()
        sel = sht.getCellRangeByPosition(0, index, PART_ATTR_LEN - 1, index)
        self.ctrl.select(sel)

    def display_results(self):
        root_node = self.tree_data.createNode("Results", True)
        self.tree_data.setRoot(root_node)
        n3 = None
        name1 = self.get_sheet().getName()
        if not self.result:
            names = self.get_sheet_names(skip=[])
            for n in names:
                n1 = self.tree_data.createNode(n, False)
                root_node.appendChild(n1)
                if n == name1:
                    n3 = n1
                sht1 = self.get_sheet(n)
                index = 0
                while True:
                    part1 = self.get_part(sht1, index)
                    if not part1:
                        break
                    k = ' '.join(part1)
                    k = '%d %s' % (index+1, k)
                    n2 = self.tree_data.createNode(k, False)
                    n1.appendChild(n2)
                    index = index + 1
        else:
            for k1,v1 in self.result.items():
                n1 = self.tree_data.createNode(k1, False)
                root_node.appendChild(n1)
                for k2, v2 in v1.items():
                    k3 = ' '.join(v2)
                    k3 = '%d %s' % (k2+1, k3)
                    n2 = self.tree_data.createNode(k3, False)
                    n1.appendChild(n2)
                if k1 == name1:
                    n3 = n1
        self.tree.expandNode(root_node)
        self.root_node = root_node
        # next line must be called after tree populated
        self.tree_model.setPropertyValues(('SelectionType', 'RootDisplayed', 'ShowsHandles', 'ShowsRootHandles', 'Editable'), (SINGLE, False, False, False, False, False))
        if n3:
            self.tree.select(n3)
            self.tree.expandNode(n3)

    def execute(self):
        if hasattr(self, 'wnd'):
            self.wnd.setVisible(True)

    def get_part(self, sht, row):
        if not sht:
            sht = self.get_sheet()
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

    def set_part(self, sht, row, p, psz):
        if not sht:
            sht = self.get_active_sheet()
        for i in range(0, psz):
            cell = sht.getCellByPosition(i, row)
            if i < len(p):
                cell.setString(p[i])
            else:
                cell.setString('')

    def get_part_count(self, sht):
        if not sht:
            sht = self.get_sheet()
        index = 0
        while True:
            cell = sht.getCellByPosition(0, index)
            if not cell.getString():
                return index
            index += 1

    def get_sheet(self, name=''):
        if not name:
            return self.ctrl.getActiveSheet()
        else:
            return self.doc.getSheets().getByName(name)

    def get_sheet_names(self, skip=['Labels']):
        names = [self.doc.Sheets.getByIndex(i).getName() for i in range(0, self.doc.Sheets.getCount())]
        for s in skip:
            if s in names:
                names.remove(s)
        return names

    def msgbox(self, msg='message', title='Message', btntype=1):
        mb = self.toolkit.createMessageBox(self.parent, MESSAGEBOX, btntype, title, msg)
        mb.execute()

    def config_read(self):
        try:
            config_provider = self.smgr.createInstanceWithContext(
                    'com.sun.star.configuration.ConfigurationProvider', self.ctx)
            config_reader = config_provider.createInstanceWithArguments(
                    'com.sun.star.configuration.ConfigurationAccess', (self.pv,))
            if hasattr(config_reader, 'x'):
                self.x = config_reader.x
            if hasattr(config_reader, 'y'):
                self.y = config_reader.y
            if hasattr(config_reader, 'w'):
                self.w = config_reader.w
            if hasattr(config_reader, 'h'):
                self.h = config_reader.h
        except:
            pass

    def config_write(self):
        try:
            config_provider = self.smgr.createInstanceWithContext(
                    'com.sun.star.configuration.ConfigurationProvider', self.ctx)
            config_writer = config_provider.createInstanceWithArguments(
                    'com.sun.star.configuration.ConfigurationUpdateAccess', (self.pv,))
            if hasattr(config_writer, 'x'):
                config_writer.x = self.x
            if hasattr(config_writer, 'y'):
                config_writer.y = self.y
            if hasattr(config_writer, 'w'):
                config_writer.w = self.w
            if hasattr(config_writer, 'h'):
                config_writer.h = self.h
            config_writer.commitChanges()
        except:
            pass

def mysearch(*args):
    MySearch(XSCRIPTCONTEXT.getComponentContext()).execute()

