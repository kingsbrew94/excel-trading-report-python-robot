from services import pricer as pz


class MainController:

    def __init__(self, record):
        self._record = record

    def create_price_at(self, sheet_name: str):
        return pz.DashboardPricer(self._record, sheet_name)

