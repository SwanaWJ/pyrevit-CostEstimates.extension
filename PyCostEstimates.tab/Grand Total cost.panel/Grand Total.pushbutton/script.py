# -*- coding: utf-8 -*-
import clr
import datetime

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window
from pyrevit import revit, DB
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

doc = revit.doc


# ---------------------------------------------------------------------
# External Event Handler (SAFE Revit API access)
# ---------------------------------------------------------------------
class RefreshTotalHandler(IExternalEventHandler):
    def __init__(self, window):
        self.window = window

    def Execute(self, uiapp):
        total = 0.0
        doc = uiapp.ActiveUIDocument.Document

        collector = DB.FilteredElementCollector(doc).WhereElementIsElementType()
        for elem in collector:
            p = elem.LookupParameter("Cost")
            if p and not p.IsReadOnly:
                try:
                    total += p.AsDouble()
                except:
                    pass

        self.window.TotalText.Text = "ZMW {:,.2f}".format(total)
        self.window.UpdatedText.Text = "Last updated: {}".format(
            datetime.datetime.now().strftime("%H:%M:%S")
        )

    def GetName(self):
        return "Refresh Grand Total"


# ---------------------------------------------------------------------
# Modeless Window
# ---------------------------------------------------------------------
class GrandTotalWindow(Window):
    def __init__(self):
        xaml_path = __file__.replace("script.py", "ui.xaml")
        clr.AddReference("IronPython.Wpf")
        import wpf
        wpf.LoadComponent(self, xaml_path)

        self.handler = RefreshTotalHandler(self)
        self.ex_event = ExternalEvent.Create(self.handler)

        # Initial populate
        self.ex_event.Raise()

    def OnRefresh(self, sender, args):
        self.ex_event.Raise()

    def OnClose(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------
# Launch window
# ---------------------------------------------------------------------
win = GrandTotalWindow()
win.Show()
