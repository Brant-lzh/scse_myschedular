import requests
import re
import xlwt
import os
from scse_myschedular.execl_style import set_style

class MyScse_login(object):
    url = 'http://class.sise.com.cn:7001/sise/'
    toLogin_url = 'http://class.sise.com.cn:7001/sise/login_check_login.jsp'
    JSESSIONID = ''
    random = ''
    post_key = ''
    post_value = ''
    username=''
    password=''
    def __init__(self,username,password):
        self.username = username
        self.password = password

    def get_values(self):
        request = requests.get(self.url)
        html = request.content.decode('GBK')
        self.JSESSIONID = request.cookies.get('JSESSIONID')
        self.random = re.findall('<input id="random"   type="hidden"  value="(.*?)"  name="random" />',html,re.S)[0]
        values = re.findall('<input type="hidden"(.*?)>',html,re.S)[0]
        self.post_key = re.findall('name="(.*?)"',values,re.S)[0]
        self.post_value = re.findall('value="(.*?)"',values,re.S)[0]


    def to_login(self):
        self.get_values()
        headers = {
            'Host': 'class.sise.com.cn:7001',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': '172',
            'Origin': 'http://class.sise.com.cn:7001',
            'Connection': 'close',
            'Referer': 'http://class.sise.com.cn:7001/sise/',
            'Cookie': 'JSESSIONID='+self.JSESSIONID,
            'Upgrade-Insecure-Requests': '1',
        }
        data = {
            self.post_key:self.post_value,
            'random':self.random,
            'username':self.username,
            'password':self.password,
            # 'token': '3105193660B8D7D3A2912127973C499F3DA7CFE3334BB',
        }
        result = requests.post(self.toLogin_url,headers=headers,data=data).content.decode('GBK')
        if result.find('<script>top.location.href=\'/sise/index.jsp\'</script>') == -1:
            return False
        else:
            return True
    def student_schedular(self):
        headers = {
            'Host': 'class.sise.com.cn:7001',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'close',
            'Referer': 'http://class.sise.com.cn:7001/sise/module/student_states/student_select_class/main.jsp',
            'Cookie': 'JSESSIONID='+self.JSESSIONID,
            'Upgrade-Insecure-Requests': '1',
        }
        url = 'http://class.sise.com.cn:7001/sise/module/student_schedular/student_schedular.jsp'
        student_class_html = requests.post(url,headers=headers).content.decode('GBK')
        MySchedular_dict = re.findall("class='font12'>(.*?)</td>",student_class_html,re.S)

        filename = 'MySchedular.xls'
        if os.path.exists(filename):
            os.remove(filename)

        wbk = xlwt.Workbook(encoding='utf-8')
        sheet = wbk.add_sheet('课程表', cell_overwrite_ok=True)  # 第二参数用于确认同一个cell单元是否可以重设值。

        sheet.write_merge(0, 0, 0, 7, '课程表', set_style(bold=True,FontHeight=13))
        title = ['', '星期一', '星期二', '星期三', '星期四', '星期五','星期六','星期日']
        for item in title:
            sheet.write(1, title.index(item), item, set_style(FontColor='white', bold=True, bgColor='blue'))
        i = 0
        j = 2
        for item in MySchedular_dict:
            content = ''
            if item != '&nbsp;':
                content = item
            if content.find('<br>') != -1:
                content = content.replace('<br>','\n')
            sheet.write(j, i, content, set_style(FontColor='white', bgColor='green'))
            i += 1
            if (i % 8 == 0):
                j+=1
                i=0
        wbk.save(filename)



