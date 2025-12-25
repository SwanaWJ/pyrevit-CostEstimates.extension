# -*- coding: utf-8 -*-
import clr
import datetime

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window
from pyrevit import revit, DB

doc = revit.doc


class GrandTotalWindow(Window):
    def __init__(self):
        xaml_path = __file__.replace("script.py", "ui.xaml")
        clr.AddReference("IronPython.Wpf")
        import wpf
        wpf.LoadComponent(self, xaml_path)

        self.update_total()

    def update_total(self):
        total = 0.0

        # Sum COST parameter of all element TYPES
        collector = DB.FilteredElementCollector(doc).WhereElementIsElementType()
        for elem in collector:
            p = elem.LookupParameter("Cost")
            if p and not p.IsReadOnly:
                try:
                    total += p.AsDouble()
                except:
                    pass

        self.TotalText.Text = "ZMW {:,.2f}".format(total)
        self.UpdatedText.Text = "Last updated: {}".format(
            datetime.datetime.now().strftime("%H:%M:%S")
        )

    def OnRefresh(self, sender, args):
        self.update_total()

    def OnClose(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------
# Launch modeless window
# ---------------------------------------------------------------------
win = GrandTotalWindow()
win.Show()
