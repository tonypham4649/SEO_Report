import requests
import json
import pandas as pd

in_data = '1033061★1033065★1030850★3099316★3883467★4594902★4952431★5001088★5001089★5045890★5213321★5373216★5873054'

the_headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8'
}


the_cookies = {
    '_ga': 'GA1.3.644156606.1597801793',
    '_ga': 'GA1.4.644156606.1597801793',
    '_td': '2411c80e-959a-4183-8a42-a05c2915ad74',
    '_gid': 'GA1.3.1258029100.1648027081',
    'auto_logindefault': '81d278cfe138cbede19eae3f1a0413a35468a479',
    'cwssid': '6mr2iffepstj0oe20su73o2mkj',
    '_gat': '1'
}

form_data = "pdata=%7B%22_t%22%3A%22b843acd40c90c5948d549176e0d6c30003647a99623bb61cac5c1%22%7D"
big_df = pd.DataFrame()
for acc in in_data.split('★'):
    sub_df = pd.DataFrame()

    the_url = 'https://kcw.kddi.ne.jp/gateway/get_detail_account_info.php?myid=1376387&_v=1.80a&_av=5&ln=ja&aid={}&get_priv_setting=false'.format(
        acc)

    response = requests.post(
        the_url, headers=the_headers, cookies=the_cookies, data=form_data)
    response_js = json.loads(response.content.decode('unicode-escape'))

    sub_df = pd.json_normalize(response_js)

    if len(big_df) == 0:
        big_df = sub_df
    else:
        big_df = big_df.append(sub_df)
