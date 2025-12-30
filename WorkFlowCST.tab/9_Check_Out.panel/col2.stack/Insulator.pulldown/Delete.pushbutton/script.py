# -*- coding: utf-8 -*-

__title__ = "Delete" 
__author__ = "Juan Achenbach"
__version__ = 'Version: 1.0'
__doc__ = """Version: 1.0
_____________________________________________________________________
Description:

Delete all insulation

_____________________________________________________________________
How-to:

Run the script to erase all insulation in the model

_____________________________________________________________________

Author: Juan Manuel Achenbach Anguita & ChatGPT"""

from pyrevit import revit, DB
from System.Collections.Generic import List
import System
from System.Windows.Forms import Form, Label, Timer


doc = revit.doc


def get_all_pipe_insulations(document):
    """Return all elements in the Pipe Insulations category."""
    return (DB.FilteredElementCollector(document)
            .OfCategory(DB.BuiltInCategory.OST_PipeInsulations)
            .WhereElementIsNotElementType()
            .ToElements())


class AutoClosePopup(Form):
    def __init__(self, message, duration_ms=2000):
        self.Text = "Info"
        self.Width = 300
        self.Height = 120

        label = Label()
        label.Text = message
        label.Dock = System.Windows.Forms.DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        self.Controls.Add(label)

        timer = Timer()
        timer.Interval = duration_ms
        timer.Tick += self.close_popup
        timer.Start()

    def close_popup(self, sender, args):
        self.Close()


def main():
    insulations = list(get_all_pipe_insulations(doc))

    if not insulations:
        popup = AutoClosePopup(u"No hi ha aïllaments company!")
        popup.ShowDialog()
        return

    ids_to_delete = List[DB.ElementId]([i.Id for i in insulations])

    with revit.Transaction("Delete Pipe Insulations"):
        deleted = doc.Delete(ids_to_delete)

    popup = AutoClosePopup(u"S'han esborrat: {} aïllaments".format(len(deleted)))
    popup.ShowDialog()


if __name__ == "__main__":
    main()
