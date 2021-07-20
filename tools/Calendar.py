import tkinter.simpledialog as sd
import tkcalendar

class CalendarDialog(sd.Dialog):
    """Dialog box that displays a calendar and returns the selected date"""
    def body(self, master):

        self.calendar = tkcalendar.Calendar(master)
        self.calendar.pack()

    def apply(self):
        self.result = self.calendar.selection_get()