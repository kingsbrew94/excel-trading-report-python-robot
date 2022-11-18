from services import helpers as hp


class Subscriber:
    _service_response = None

    @staticmethod
    def get_env():
        return hp.get_env()

    @staticmethod
    def get_dashboard_name():
        """

        :rtype: object
        """
        return Subscriber.get_env()['OUTPUT_WORKBOOK']

    @staticmethod
    def get_response():
        return Subscriber._service_response

    @staticmethod
    def on_connect_to_fetch():
        Subscriber._service_response = hp.connect_data_service()
        return Subscriber
