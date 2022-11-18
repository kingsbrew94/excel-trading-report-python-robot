import openpyxl as xl
import win32com.client as win32
import logging as log
import time, random as rn

log.basicConfig(format='%(asctime)s | %(message)s', level=log.INFO)


class View:
    _xlApp = None
    _xWb   = None

    def __init__(self, app_name: str, sheet: str, data_list: list):
        self._app_name = app_name  # Workbook file path
        self._target_sheet = sheet
        self._active_sheet = None
        self._wb = None
        self._data_list = data_list
        self._intervals = [0, 0.025, 0.1, 0, 0.2, 0.25, 0]
        self._sheet_cell_heads = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB']

    def init_interface_client(self):
        log.info("LAUNCHING APPLICATION ...")
        import pythoncom
        View._xlApp = win32.Dispatch('Excel.Application', pythoncom.CoInitialize())
        View._xlApp.Visible = True
        View._xWb = View._xlApp.Workbooks.Open(self._app_name)
        log.info("APPLICATION LAUNCHED!!!")
        return self

    def get_client_interface(self):
        log.info("REASSIGNING INTERFACE ...")
        return View._xWb

    def _set_rows(self,data_row,index,sheet):
        i: int = 0
        data_size = len(data_row)
        for data in data_row:
            # print(str(data)," ds: ", data_size," --- ",str(i))
            sheet.Range('{}{}'.format(self._sheet_cell_heads[i], index)).Value = data \
                if ((data_size - 1) == i) else float(data)
            i += 1

    def _set_header(self,data_row,index,sheet):
        i: int = 0
        for data in data_row:
            sheet.Range('{}{}'.format(self._sheet_cell_heads[i], index)).Value = data
            i += 1

    def reassign_interface(self, cli, data_list):
        sec = rn.choice(self._intervals)
        actual_data = data_list
        if sec != 0:
            log.info("UPDATING PRICE...")
            time.sleep(sec)
        sheet = cli.Worksheets(self._target_sheet)

        self._set_header(data_row=data_list[0], index=1, sheet=sheet)
        del actual_data[0]
        index: int = 2
        for data_row in data_list:
            self._set_rows(data_row=data_row,index=index,sheet=sheet)
            time.sleep(sec)
            log.info("{}-{} AFFECTED!".format("ROW", index))

            index += 1
        if index >= 50:
            cli.Save()
        return True

    def load_initial_data(self):
        try:
            self._wb = xl.load_workbook(self._app_name)
            log.info("APPLICATION ACCESS SUCCESSFUL!!!")
            self._wb.remove_sheet(self._wb.get_sheet_by_name(self._target_sheet))
            self._wb.create_sheet(self._target_sheet)
            log.info(self._target_sheet)
            self._active_sheet = self._wb.get_sheet_by_name(self._target_sheet)
            log.info("INITIALIZING APPLICATION >> "+self._app_name)
            for data_row in self._data_list:
                self._active_sheet.append(data_row)
            log.info("APPLICATION INITIALIZED!!!")
            self._wb.save(self._app_name)
        except Exception as ex:
            log.info(ex)
