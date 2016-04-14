
from com.sun.star.awt import XActionListener, XWindowListener
from com.sun.star.awt import Rectangle, WindowDescriptor
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.awt.PushButtonType import STANDARD
from com.sun.star.awt.ScrollBarOrientation import VERTICAL
from com.sun.star.awt.WindowAttribute import BORDER, SHOW, SIZEABLE, MOVEABLE, CLOSEABLE
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.view.SelectionType import SINGLE
from com.sun.star.beans import PropertyValue
from uno import createUnoStruct
from unohelper import Base

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
    def __init__(self, cb):
        self.cb = cb
    def windowResized(self, evt):
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
        tree_model = self.smgr.createInstanceWithContext('com.sun.star.awt.tree.TreeControlModel', self.ctx)
        tree_model.SelectionType = SINGLE
        tree.setPosSize(0, 0, self.w, self.h, POSSIZE)
        tree.setModel(tree_model)
        cont.addControl('tree', tree)
        self.tree = tree

        tree_data = self.createUnoService("com.sun.star.awt.tree.MutableTreeDataModel")
        root_node = tree_data.createNode("Root", True)
        tree_data.setRoot(root_node)
        tree_model.DataModel = tree_data
        tree_model.RootDisplayed = True

        for i in range(0, 20):
            n1 = tree_data.createNode("Node_1", False)
            root_node.appendChild(n1)

        listener = TreeListener(self.resize_cb)
        wnd.addWindowListener(listener)

        self.wnd = wnd

    def resize_cb(self):
        w = self.wnd.PosSize.Width
        h = self.wnd.PosSize.Height
        self.cont.setPosSize(0, 0, w, h, POSSIZE)
        self.tree.setPosSize(0, 0, w, h, POSSIZE)

    def createUnoService(self, name):
        return self.ctx.getServiceManager().createInstance(name)

    def execute(self):
        if hasattr(self, 'wnd'):
            self.wnd.setVisible(True)

    def close_cb(self):
        if hasattr(self, 'wnd'):
            self.config_write()
            self.wnd.setVisible(False)

    def init_buttons(self, dlg):
        model = dlg.getModel()
        btn = self.init_button(dlg, model.Width/3, 1, 'Close')
        listener = ButtonListener(self.close_cb)
        btn.addActionListener(listener)
        return dlg

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

    def get_active_sheet(self):
        return self.doc.CurrentController.ActiveSheet

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
    import os, sys
    prefix = 'file://'
    fname = __file__
    if fname.find(prefix) == 0:
        fname = __file__.replace(prefix, '', 1)
    dirname = os.path.dirname(fname)
    if dirname not in sys.path:
        sys.path.append(os.path.dirname(fname))
    #global mp
    #mp = __import__('myparts')
    MS = type('MySearch', (MySearch,), {})
    MS(XSCRIPTCONTEXT.getComponentContext()).execute()

