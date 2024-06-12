from Elements import *

case = os.path.basename(__file__)

class Test:
    @allure.title(case)
    def test(self):
        url = ele.domain()+"/order/HKTV/consignmentOrderSummary"

        payload = json.dumps({
        "deliveryEnd": "",
        "deliveryStart": "",
        "endDate": "2023-01-19",
        "pickUpEnd": "",
        "pickUpStart": "",
        "startDate": "2023-01-19",
        "overseasWaybill": "",
        "deliveryMethod": "STANDARD_DELIVERY",
        "warehouseList": [
            "H036700501",
            "H036700508",
            "H036700509",
            "H036700599"
        ],
        "storeCode": "H0367005",
        "todayVersion": True
        })

        headers = {
        'Authorization': 'Bearer ' + ele.accessToken(),
        'Content-Type': 'application/json'
        }

        response = requests.request("POST", url, headers=headers, data=payload)

        logging.info(response.text)
