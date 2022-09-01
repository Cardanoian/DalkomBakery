import openpyxl
import pandas as pd
from datetime import datetime
import re


def make_shop(file):
    now = datetime.now()
    try:
        data = pd.read_csv(file, encoding="utf-8")
    except UnicodeDecodeError:
        data = pd.read_csv(file, encoding="cp949")
    result = []
    for i in range(len(data)):
        row = data.iloc[i]
        site = "메이크샵"
        name = row["수령인명"].strip()
        zip_code = re.findall("[0-9]+", row["수령인 우편번호(XXXXXX)"])[0]
        address = row["수령인 주소"].strip()
        phone = row["수령인 핸드폰"]
        home_phone = row["수령인 전화"]
        try:
            deliver_date = re.findall("[0-9]+", row["상품옵션"].split("수령 희망일: ")[1].split(" / ")[0])
            print(deliver_date)
            for j in range(len(deliver_date)):
                print(deliver_date[j])
                if len(deliver_date[j]) == 1:
                    deliver_date[j] = "0" + deliver_date[j]
            print(deliver_date)
            if len(deliver_date) < 2 or len(deliver_date) > 6:
                raise Exception("!")
            elif len(deliver_date) < 4:
                if len(deliver_date) == 3:
                    del deliver_date[0]
                deliver_date = "".join(deliver_date).strip()
            else:
                if len(deliver_date) == 4:
                    first = ''.join(deliver_date[:2])
                    last = ''.join(deliver_date[2:])
                else:
                    first = ''.join(deliver_date[:3])
                    last = ''.join(deliver_date[3:])
                while int(first[0]) > 1:
                    first = first[1:]
                while int(last[0]) > 1:
                    last = last[1:]
                deliver_date = f"{first} ~ {last}"
        except:
            deliver_date = ""
        order = f"{row['주문상품명'].strip()} ({str(row['주문품목 수량'])})"
        order_detail = {"출처": site, "수신자": name, "수신자우편번호": zip_code, "수신자주소": address,
                        "수신자전화번호": home_phone, "수신자휴대전화": phone, "상품명(수량)": order, "수령희망일": deliver_date,
                        "전달글": "", "모델명": ""}
        result.append(order_detail)
    # print(result)
    return result


def naver_store(file):
    now = datetime.now()
    data = openpyxl.load_workbook(filename=file).active
    data = pd.DataFrame(list(data.values)[2:], columns=list(data.values)[1])
    # print(data, type(data))
    result = []
    site = "스마트스토어" if "판매채널" in data.columns else "네이버페이"
    for i in range(len(data)):
        row = data.iloc[i]
        # print(row, type(row))
        name = row["수취인명"].strip()
        zip_code = re.findall("[0-9]+", row["우편번호"])[0]
        address = row["통합배송지"].strip() if site == "스마트스토어" else row["배송지"].strip()
        phone = row["수취인연락처1"]
        home_phone = row["수취인연락처2"]
        try:
            deliver_date = re.findall("[0-9]+", row["옵션정보"].split("수령 희망일: ")[1].split(" / ")[0])
            print(deliver_date)
            for j in range(len(deliver_date)):
                print(deliver_date[j])
                if len(deliver_date[j]) == 1:
                    deliver_date[j] = "0" + deliver_date[j]
            print(deliver_date)
            if len(deliver_date) < 2 or len(deliver_date) > 6:
                raise Exception("!")
            elif len(deliver_date) < 4:
                if len(deliver_date) == 3:
                    del deliver_date[0]
                deliver_date = "".join(deliver_date).strip()
            else:
                if len(deliver_date) == 4:
                    first = ''.join(deliver_date[:2])
                    last = ''.join(deliver_date[2:])
                else:
                    first = ''.join(deliver_date[:3])
                    last = ''.join(deliver_date[3:])
                while int(first[0]) > 1:
                    first = first[1:]
                while int(last[0]) > 1:
                    last = last[1:]
                deliver_date = f"{first} ~ {last}"
        except:
            deliver_date = ""

        order = f"{row['상품명'].strip()} ({str(row['수량'])})"
        order_detail = {"출처": site, "수신자": name, "수신자우편번호": zip_code, "수신자주소": address,
                        "수신자전화번호": home_phone, "수신자휴대전화": phone, "상품명(수량)": order, "수령희망일": deliver_date,
                        "전달글": "", "모델명": ""}
        result.append(order_detail)
    # print(result)
    return result
