import win32com.client as win32
from services import clocks as cl
import time, random as rn
from services import subscriber as sub


class Reporter:
    _xlApp = None
    _xlOpened = False

    def __init__(self, file_path, sheets):
        self._file_path = file_path
        self._wb = None
        self._sheets = sheets
        self._intervals = [0, 0.025, 0.1, 0, 0.2, 0.25, 0]

    def start_client(self):
        if Reporter._xlOpened is False:
            import pythoncom
            Reporter._xlApp = win32.Dispatch('Excel.Application', pythoncom.CoInitialize())
            Reporter._xlApp.Visible = True
            self._wb = Reporter._xlApp.Workbooks.Open(self._file_path)
        return self

    def load_real_data(self, end_point_data: list):
        rn.shuffle(self._sheets)
        timer = cl.Clock(data_list=end_point_data, callback=lambda: self._fill_cells())
        timer.start()

    def _fill_cells(self):
        self.start_client()
        alive = True
        intervals = [5]
        Reporter._xlOpened = True
        while alive:
            cl.Clock.moments = sub.Subscriber.on_connect_to_fetch().get_response()
            try:
                self._set_activity()
                print("Showing dashboard....")
            except Exception as ex:
                print(ex)
            time.sleep(rn.choice(intervals))

    def _set_activity(self):
        for active in self._sheets:
            self._processor(active)

    def _processor(self, active: str):
        cell_index = 2
        for data in cl.Clock.moments:
            if data['CURRENCY'] == active:
                sheet = self._wb.Worksheets(active)
                sheet.Range('A{}'.format(cell_index)).Value = data['REAL TIME']
                sheet.Range('B{}'.format(cell_index)).Value = (float) (data['POWER'].replace('\x00', ''))
                sheet.Range('C{}'.format(cell_index)).Value = (float) (data['SMA6'].replace('\x00', ''))
                sheet.Range('D{}'.format(cell_index)).Value = (float) (data['SMA30'].replace('\x00', ''))
                sec = rn.choice(self._intervals)
                if sec != 0:
                    time.sleep(sec)
                cell_index += 1

    def get_intervals(self):
        return self._intervals
