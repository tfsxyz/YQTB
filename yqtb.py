from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import time
import csv
import smtplib
from email.mime.text import MIMEText


# 单一用户的填报
class Fill:
    def __init__(self):
        self.opt = Options()
        self.opt.add_argument("--headless")
        self.opt.add_argument("--disbale-gpu")
        # self.web = Chrome()  # 有头
        self.web = Chrome(options=self.opt)  # 把参数配置设置到浏览器中 无头
        self.username = '2018301817'
        self.password = "A565656a"  # 这里可以改为表单
        self.state = True
        self.name = ''

    def open_url(self):
        self.web.get("http://yqtb.nwpu.edu.cn/wx/xg/yz-mobile/index.jsp")

    def login(self):
        # 输入用户名和密码
        self.web.find_element_by_xpath('//*[@id="username"]').send_keys(self.username)
        self.web.find_element_by_xpath('//*[@id="password"]').send_keys(self.password)
        login = self.web.find_element_by_xpath('//*[@id="fm1"]/div[4]/div/input[5]')
        login.click()  # 点击事件
        try:
            everyday = self.web.find_element_by_xpath('/html/body/div/div[5]/ul/li[1]/a/i')
            everyday.click()
        except NoSuchElementException:
            print('The username or password is not available')
            self.state = False

    def fill(self):
        # 一系列点击，必须默认其他的都是对的
        check = self.web.find_element_by_xpath('//*[@id="rbxx_div"]/div[3]/label[2]/div[1]/p')
        check.click()
        put_first = self.web.find_element_by_xpath('//*[@id="rbxx_div"]/div[16]/div/a')
        put_first.click()
        know = self.web.find_element_by_xpath('//*[@id="qrxx_div"]/div[2]/div[11]/label/div[1]/i')
        know.click()
        put_second = self.web.find_element_by_xpath('//*[@id="save_div"]')
        put_second.click()

    def verify(self):
        try:
            verify = self.web.find_element_by_xpath('//*[@id="rbxx_div"]/div[1]/h1/i').text
        except NoSuchElementException:
            print('No success')  # 这一块只能用异常了
            self.state = False
        else:
            if '已提交' in verify:
                print('Success')
                self.state = True
            else:
                print('Unknown Error')
                self.state = False

    def obtain_name(self):
        basic_information = self.web.find_element_by_xpath('//*[@id="form1"]/div[4]/div/div[1]/div[3]')
        basic_information.click()
        self.name = self.web.find_element_by_xpath('//*[@id="form1"]/div[5]/div[2]/span').text
        today = self.web.find_element_by_xpath('//*[@id="form1"]/div[3]/div/div[1]/div[1]')
        today.click()

    def main(self):
        print('The username is ' + self.username)
        self.open_url()
        self.login()
        if self.state:
            self.fill()
            self.obtain_name()
            # time.sleep(1)
            self.verify()
        self.web.close()
        return self.state

    def test(self):
        print("Start the test")
        self.open_url()
        self.login()
        # self.fill()
        # self.verify()
        self.obtain_name()
        print(self.name)
        time.sleep(5000)


class YQTB:
    def __init__(self):
        self.user_csv = open("user.csv", mode="r")
        self.email_csv = open("email.csv", mode="r")
        self.id_list = []
        self.i = 0
        self.state = False
        self.name = ''
        self.mail_host = ''
        self.mail_user = ''
        self.mail_pass = ''
        self.sender = ''

    def read_csv(self):
        reader = csv.reader(self.user_csv)
        for line in reader:
            self.id_list.append({'username': line[0], 'password': line[1], 'email': line[2]})

    def fill(self):
        single_fill = Fill()
        single_fill.username = self.id_list[self.i]['username']
        single_fill.password = self.id_list[self.i]['password']
        try:
            single_fill.main()
        except NoSuchElementException:
            print('Unknown Error Out')
        else:
            self.state = single_fill.state  # 默认为假，异常后还是假
            self.name = single_fill.name  # 获取姓名，不过有可能也获取不到

    def email_login(self):
        reader = csv.reader(self.email_csv)
        for line in reader:  # 这写的有点傻
            self.mail_host = line[0]
            self.mail_user = line[1]
            self.mail_pass = line[2]
            self.sender = line[3]
            print(line)

        try:
            self.smtpObj = smtplib.SMTP()
            # 连接到服务器
            self.smtpObj.connect(self.mail_host, 25)
            # 登录到服务器
            self.smtpObj.login(self.mail_user, self.mail_pass)
            print('Email login success')
        except smtplib.SMTPException as e:
            print('Error email login', e)  # 打印错误

    def email_send(self):
        if self.state:
            message_text = '尊敬的' + self.name + '同学:\n' + '         您今日的疫情填报已经完成，感谢您使用本软件。'
            message = MIMEText(message_text, 'plain', 'utf-8')
            # 邮件主题
            message['Subject'] = '疫情填报通知'
            # 发送方信息
            message['From'] = 'cjmnxc@163.com'
            # 接受方信息
            message['To'] = self.id_list[self.i]['email']
        else:
            message_text = '尊敬的' + self.name + '同学:\n' + '         您今日的疫情填报未完成，请联系管理员。'
            message = MIMEText(message_text, 'plain', 'utf-8')
            # 邮件主题
            message['Subject'] = '疫情填报通知'
            # 发送方信息
            message['From'] = 'cjmnxc@163.com'
            # 接受方信息
            message['To'] = self.id_list[self.i]['email']

        self.smtpObj.sendmail(
            self.sender, self.id_list[self.i]['email'], message.as_string())
        print('Send email successfully')

    def main(self):
        print('Start the progress')
        self.state = False
        self.read_csv()
        self.email_login()
        while self.i <= len(self.id_list) - 1:
            self.fill()
            # 失败二次尝试
            if not self.state:
                self.fill()

            self.email_send()
            self.i += 1
        self.smtpObj.quit()
        print('Finish the progress')

    def test(self):
        print('Start test')
        self.read_csv()
        self.email_login()
        # print(self.id_list)e
        self.fill()
        self.email_send()


if __name__ == '__main__':
    # fill = Fill()
    # fill.main()
    yqtb = YQTB()
    yqtb.main()
