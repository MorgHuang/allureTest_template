import requests, json, yaml, os, datetime, logging, allure


# reference: https://mms-order-dev.hkmpcl.com.hk/order/swagger-ui/index.html?configUrl=/order/v3/api-docs/swagger-config#/
class ele:

    def domain():
        dataPath = os.path.join(os.path.abspath(os.getcwd()), "domain.yaml")
        with open(dataPath, encoding="utf-8") as f:
            config = yaml.load(f.read(), Loader=yaml.SafeLoader)
        domain = config['domain']['order']
        # print(domain)
        return domain

    def accessToken():
        url = "https://mms-user-staging.hkmpcl.com.hk/user/login/merchantAppLogin2"

        payload = json.dumps({
            "username": "test",
            "password": "qwe123"
        })
        headers = {
            'accept': 'application/json',
            'Content-Type': 'application/json'
        }

        token = requests.request("POST", url, headers=headers, data=payload)

        # 把 token 撈出來
        accessToken = token.json()["accessToken"]
        # print(accessToken)
        return accessToken

    def updateDate():
        # 取得當前時間
        now = datetime.datetime.now() + datetime.timedelta(seconds=60)

        # 將當前時間格式化
        formatted_time = now.strftime(r"%Y%m%d %H%M%S")
        # print(formatted_time)
        return formatted_time

# ele.domain()
# ele.accessToken()
# ele.updateDate()
