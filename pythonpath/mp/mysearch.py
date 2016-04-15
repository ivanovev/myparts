
from com.sun.star.awt import XActionListener, XWindowListener
from com.sun.star.awt import Rectangle, WindowDescriptor
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.awt.PushButtonType import STANDARD
from com.sun.star.awt.ScrollBarOrientation import VERTICAL
from com.sun.star.awt.WindowAttribute import BORDER, SHOW, SIZEABLE, MOVEABLE, CLOSEABLE
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.lang import XEventListener
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

class TreeListener(Base, XWindowListener):
    def __init__(self, cbm, cbr):
        self.cbm = cbm
        self.cbr = cbr
    def windowMoved(self, evt):
        self.cbm()
    def windowResized(self, evt):
        self.cbr()

class TreeListener2(Base, XEventListener):
    def __init__(self, cb):
        self.cb = cb
    def disposing(self, evt):
        self.cb()

class MySearch(object):
    def __init__(self, ctx, part1=[]):
        self.ctx = ctx
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.doc = self.desktop.getCurrentComponent()
        self.smgr = self.ctx.getServiceManager()
        self.parent = self.desktop.getCurrentFrame().getContainerWindow()
        self.toolkit = self.parent.getToolkit()
        if self.get_active_sheet().Name == 'Labels':
            return
        self.x = 100
        self.y = 100
        self.w = 200
        self.h = 300
        pv = PropertyValue()
        pv.Name = 'nodepath'
        pv.Value = 'vnd.myparts.settings/mysearch'
        self.pv = pv
        self.config_read()
        self.init_dlg()
        self.result = OD()
        if part1:
            self.part_search(None, part1)
        self.display_results()

    def new_search(self, ctx, part1=[]):
        out(MySearch)
        return MySearch(ctx, part1)

    def init_dlg(self):
        rect = createUnoStruct("com.sun.star.awt.Rectangle")
        rect = Rectangle(self.x, self.y, self.w, self.h)

        desc = createUnoStruct("com.sun.star.awt.WindowDescriptor")
        frame = self.smgr.createInstanceWithContext('com.sun.star.frame.Frame', self.ctx)
        frame.Title = 'MySearch'
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

        listener = TreeListener(self.move_cb, self.resize_cb)
        wnd.addWindowListener(listener)

        listener = TreeListener2(self.config_write)
        wnd.addEventListener(listener)

        self.wnd = wnd

    def createUnoService(self, name):
        return self.ctx.getServiceManager().createInstance(name)

    def move_cb(self):
        self.x = self.wnd.PosSize.X
        self.y = self.wnd.PosSize.Y

    def resize_cb(self):
        self.w = self.wnd.PosSize.Width
        self.h = self.wnd.PosSize.Height
        self.cont.setPosSize(0, 0, self.w, self.h, POSSIZE)
        self.tree.setPosSize(0, 0, self.w, self.h, POSSIZE)

    def part_cmp(self, sht, index, part):
        if sht == None:
            sht = self.get_active_sheet()
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
            sht = self.get_active_sheet()
        index = start
        while True:
            part2 = self.get_part(sht, index)
            if not part2:
                return index
            if self.part_cmp(sht, index, part):
                return index
            index += 1

    def display_results(self):
        root_node = self.tree_data.createNode("Results", True)
        self.tree_data.setRoot(root_node)
        n3 = None
        if not self.result:
            for i in range(0, 20):
                n1 = self.tree_data.createNode("Node_1", False)
                root_node.appendChild(n1)
            self.tree.select(n1)
        else:
            name1 = self.get_active_sheet().getName()
            for k1,v1 in self.result.items():
                n1 = self.tree_data.createNode(k1, False)
                root_node.appendChild(n1)
                for k2, v2 in v1.items():
                    k3 = ' '.join(v2)
                    n2 = self.tree_data.createNode(k3, False)
                    n1.appendChild(n2)
                if k1 == name1:
                    n3 = n1
        self.tree.expandNode(root_node)
        # next line must be called after tree populated
        self.tree_model.setPropertyValues(('SelectionType', 'RootDisplayed', 'ShowsHandles', 'ShowsRootHandles', 'Editable'), (SINGLE, False, False, False, False, False))
        if n3:
            self.tree.select(n3)
            self.tree.expandNode(n3)

    def execute(self):
        if hasattr(self, 'wnd'):
            self.wnd.setVisible(True)

    def close_cb(self):
        if hasattr(self, 'wnd'):
            self.config_write()
            self.wnd.setVisible(False)

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

    def get_active_sheet(self):
        return self.doc.CurrentController.ActiveSheet

    def get_sheet_names(self, skip=['Labels']):
        names = [self.doc.Sheets.getByIndex(i).getName() for i in range(0, self.doc.Sheets.getCount())]
        for s in skip:
            if s in names:
                names.remove(s)
        return names

    def msgbox(self, msg='message', title='Message', btntype=1):
        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow()
        toolkit = window.getToolkit()
        mb = toolkit.createMessageBox(window, MESSAGEBOX, btntype, title, msg)
        mb.execute()

    def config_read(self):
        try:
            config_provider = self.smgr.createInstanceWithContext(
                    'com.sun.star.configuration.ConfigurationProvider', self.ctx)
            config_reader = config_provider.createInstanceWithArguments(
                    'com.sun.star.configuration.ConfigurationAccess', (self.pv,))
            self.x = config_reader.x
            self.y = config_reader.y
            self.w = config_reader.w
            self.h = config_reader.h
        except:
            pass

    def config_write(self):
        try:
            config_provider = self.smgr.createInstanceWithContext(
                    'com.sun.star.configuration.ConfigurationProvider', self.ctx)
            config_writer = config_provider.createInstanceWithArguments(
                    'com.sun.star.configuration.ConfigurationUpdateAccess', (self.pv,))
            config_writer.x = self.x
            config_writer.y = self.y
            config_writer.w = self.w
            config_writer.h = self.h
            config_writer.commitChanges()
        except:
            pass

def mysearch(*args):
    MySearch(XSCRIPTCONTEXT.getComponentContext()).execute()

