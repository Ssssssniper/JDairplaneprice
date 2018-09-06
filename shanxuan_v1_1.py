import requests
from openpyxl import Workbook
# 实现了写入excel功能，但是没实现循环写入同一个excel。

def Get_Plan_Money(depCity, arrCity, depDate):
    global money_dict
    headers = {
        'Host': "jipiao.jd.com",
        'Accept-Language': "zh-CN,zh;q=0.9",
        'Accept-Encoding': "gzip, deflate, br",
        'X-Requested-With': 'XMLHttpRequest',
        'Connection': "keep-alive",
        'User-Agent': "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"
    }
    # proxies = {
    #      'http': 'http://1.31.23.209:80',
    #      'http': 'http://110.243.166.164',
    # }
    roon_session = requests.session()
    payload = {'depCity': depCity, 'arrCity': arrCity, 'depDate': depDate, 'arrDate': depCity, 'queryModule': '1',
               'lineType': 'OW', 'queryType': 'listquery'}
    url = 'https://jipiao.jd.com/search/queryFlight.action?'
    page = roon_session.get(url,headers=headers, params=payload).json()

    money_dict = {}
    print('%s From%sTo%s:' % (depDate, depCity, arrCity))
    print(page)
    for i in range(4):
        if page['data']['flights'] is None:
            print('Already request{} times'.format(i+1))
            page = roon_session.get(url, headers=headers, params=payload).json()
        else:
            for plan in page['data']['flights']:
                if plan['airways'] == 'MU':
                    print('%s-%slowest price：%s' % (plan['airwaysCn'], plan['flightNo'], plan['bingoLeastClassInfo']['price']))
                    money_dict[plan['flightNo']] = plan['bingoLeastClassInfo']['price']
            # print(money_dict.items())
    print('请休息一会。')
    return money_dict


while True:
    form_list = input('please input depcity,arrcity,date:（Date Formate：xxxx-xx-xx）').split()
    Get_Plan_Money(form_list[0], form_list[1], form_list[2])
    wb = Workbook()
    ws1 = wb.create_sheet('{}'.format(form_list[2]))
    for item in money_dict.items():
        ws1.append(item)
    wb.save('{}-{}.xlsx'.format(form_list[0], form_list[1]))
    cur = input('Continue??-----Y/N:')
    if cur in ('Y', 'y'):
        continue
    else:
        break
