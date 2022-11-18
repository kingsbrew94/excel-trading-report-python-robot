import json, logging as log


def get_env():
    with open('env.json', 'r') as file:
        return json.loads("".join(list(file.readlines())))


def get_time_interval():
    return get_env()["TIME_INTERVAL"]


def connect_data_service():
    try:
        env: list = get_env()
        endpoint: str = env["DATA_SERVICE_LOCATION_ID"] + env["DATA_SERVICE_URL"].format(env["DATA_STORE"])
        with open(endpoint) as file:
            content = file.readlines()
        head = content[:1]
        rows = content[1:]
        row_size = len(rows) - 1;
        rows.reverse()
        price_rows = []
        header = ("ID,{}".format(head[0].replace("\'", "").replace("\x00", "").replace("\t", "").replace("\n", ""))).split(",")
        price_rows.append(header)
        for i in range(0, env["DATA_LIMIT"]):
            if i > row_size:
                break
            row = (str(i+1)+","+rows[i].replace("\'", "").replace("\x00", "").replace("\t", "").replace("\n", "")).split(",")
            rwc = [rw for rw in row]
            price_rows.append(rwc)
        return price_rows
    except Exception as ex:
        log.info(ex)
        return None
