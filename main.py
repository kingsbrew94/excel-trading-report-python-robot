import logging as log
import threading
from station import mainComtroller as mc
from services import Path as p
from services import subscriber as sub
from services import globals as G
import os, time


def price_action(prs,sub, name):
    time.sleep(2)
    alive = True
    log.info("---- Running Thread "+name)
    client_view = prs.get_client_view()
    cli = client_view.get_client_interface()
    while alive:
        try:
            prs.update_data(cli,sub.Subscriber.on_connect_to_fetch().get_response())
        except Exception as ex:
            log.info("Re-Executing Process...")
            log.error(ex)


if __name__ == '__main__':
    p.Path.set_path(os.path.join(os.getcwd()))
    passive_data = sub.Subscriber.on_connect_to_fetch().get_response()
    prs = mc.MainController(passive_data). \
        create_price_at(G.Sheet.CURRENCY_PRICE_SHEET)
    price_chart_bot = threading.Thread(target=price_action,args=(prs,sub,"CURRENCY_PRICE_SHEET_BOT"))
    price_chart_bot.start()
    log.info("MAIN THREAD EXECUTED")

