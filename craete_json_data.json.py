# encoding=utf-8
import json
import os


if __name__ == '__main__':
    request_data = {
        "actorId": "1001264",
        "androidId": "467fcaa9",
        "checkNo": "PC01000091",
        "orgId": "A931",
    }
    items = []
    index = 1
    while index <= 5000:
        items.append({
            "itemId": str(index),
            "itemName": "康麦斯牌美康宁(褪黑素)片_18g(300mg*60片)",
            "itemSpec": "18G(300MG*60片)",
            "itemUnit": "瓶",
            "producer": "美国康龙集团(KangLong Group Corp U.S.A)",
            "itemZjm": "KMSPMKNTHSP",
            "barcode": "763052881286",
            "makeNo": str(index),
            "validDate": "2020-06-20",
            "actDate": "2019-07-17 11:28:59",
            "actorId": "1001264",
            "checkNo": "PC01000091",
            "checkQty": index,
            "orgId": "A931",
            "stockQty": "2.0000"
        })
        index += 1
        print(index)

    request_data["list"] = items

    file = open('json.json', mode='w+', encoding='utf-8')
    file.write(str(json.dumps(request_data, ensure_ascii=False)))
    file.write('\n')
    file.close()

