from services import subscriber as sb
from services import Path as p
from Interface import view as vw
import logging as log


class DashboardPricer:

    def __init__(self, data_list: list, sheet: str):
        self._data_list = data_list
        self._sheet_name = sheet
        self._file_path = ('{}\\app\\'.format(p.Path.get_path())) + sb.Subscriber.get_dashboard_name()
        log.info("ACCESSING APPLICATION >> "+self._file_path)
        self._vw = vw.View(app_name=self._file_path, sheet=sheet, data_list=data_list)

    def initialize_data(self):
        self._vw.load_initial_data()
        return self

    def get_client_view(self):
       return self._vw.init_interface_client()

    def update_data(self,client_interface, data_list):
       return self._vw.reassign_interface(client_interface, data_list)
