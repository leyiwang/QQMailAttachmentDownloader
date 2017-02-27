#coding=utf-8
'''
  Title: QQ Mail Attachment Downloader
  Version: V0.3
  Author: Leyi Wang
  Date: Last update 2016-12-4
  Email: leyiwang.cn@gmail.com
'''

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import os, time, getpass, sys, logging, datetime

logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level = logging.INFO)
class QQMailAttachmentDownloader(object):
    CHROMEDRIVE = r'C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe'
    def __init__(self, page_list, download_dir=None):
        self.chrome_init(download_dir)
        self.username = raw_input("Please input the QQ number:")
        self.password = getpass.getpass("Please Input the password:")
        self.browser = webdriver.Chrome(self.CHROMEDRIVE, chrome_options=self.options)
        self.browser.maximize_window()
        self.page_list = page_list
        self.run()

    def chrome_init(self, download_dir):
        if not os.path.exists(self.CHROMEDRIVE):
            logging.error(self.CHROMEDRIVE + ' is not exisit, please download chromedriver.exe')
            sys.exit()
        self.download_dir = self.check_download_dir(download_dir)
        self.options = webdriver.ChromeOptions()
        prefs = {'profile.default_content_settings.popups': 0, r'download.default_directory':self.download_dir}
        self.options.add_experimental_option('prefs', prefs)
        
    @staticmethod
    def check_download_dir(download_dir):
        if download_dir == None:
            download_dir = os.getcwd() + os.sep + r'download'
        if not os.path.exists(download_dir):
            os.mkdir(download_dir)
        return download_dir
    
    def downloader(self):
        time.sleep(3)
        WebDriverWait(self.browser, 10).until(lambda the_driver: the_driver.find_element_by_id('readmailbtn_link').is_displayed())
        self.browser.find_element_by_xpath("//a[@id='readmailbtn_link']").click()#进入收件箱
        self.__switch_to_iframe('mainFrame')
        maillisturl_base_list = self.browser.execute_script('return document.location.href').split('page=0')
        for page_index in self.page_list:
            if page_index!=0:
                maillist_url = ('page='+ str(page_index)).join(maillisturl_base_list)
                self.browser.get(maillist_url)
                self.__switch_to_iframe('mainFrame')
            mail_id_lst = [item.get_attribute("value") for item in self.browser.find_elements_by_xpath("//td[@class='cx']/input")]
            logging.info(u"正在读取收件箱第"  + str(page_index+1) + u'页,该页一共' + str(len(mail_id_lst)) + u'封邮件')
            for index, mail_id in enumerate(mail_id_lst):
                if index==0:
                    self.browser.find_element_by_xpath("//td[@class='l']").click()
                    mailurl_base_list = self.browser.execute_script('return document.location.href').split(mail_id)
                    self.browser.back()
                mail_url = mail_id.join(mailurl_base_list)
                self.browser.get(mail_url)
                self.__switch_to_iframe('mainFrame')
                logging.info(u"正在读取第" + str(index+1) + u"封邮件的附件下载地址...")
                time.sleep(3)
                filename_lst = self.browser.find_elements_by_xpath("//div[@class='name_big']/span")
                download_href = self.browser.find_elements_by_xpath("//div[@class='down_big']/a[@unset='true']")
                logging.info(u"该邮件一共有" + str(len(download_href)) + u"个附件")
                for k in range(len(download_href)):
                    filename = filename_lst[k*2].text
                    logging.info(u"正在下载第" + str(k+1) + u'个附件......' + filename)
                    download_href[k].click()
                    time.sleep(5)
            logging.info(u'当前页面所有附件下载完成.')
        logging.info(u'所有页面附件下载完成.')
    
    def login(self):
        self.browser.get('https://mail.qq.com/')
        self.browser.switch_to.frame('login_frame')
        self.browser.find_element_by_xpath("//div/a[@id='switcher_plogin']").click()
        self.browser.find_element_by_xpath("//input[@id='u']").send_keys(self.username)
        self.browser.find_element_by_xpath("//input[@id='p']").send_keys(self.password)
        time.sleep(5)
        self.browser.find_element_by_xpath("//input[@id='login_button']").click()

    def run(self):
        self.login()
        self.downloader()

    def __switch_to_iframe(self, iframe_name, timeout=10):
        WebDriverWait(self.browser, timeout).until(lambda driver: driver.find_element_by_id(iframe_name).is_displayed())
        self.browser.switch_to.frame(iframe_name)

if __name__ == '__main__':
    start_time = datetime.datetime.now()
    downloader = QQMailAttachmentDownloader(range(9))
    end_time = datetime.datetime.now()
    logging.info("\nDone! Seconds cost:"+str((end_time - start_time).seconds))
