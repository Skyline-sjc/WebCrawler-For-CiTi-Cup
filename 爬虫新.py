from faulthandler import is_enabled
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from seleniumwire import webdriver as webdriver_wire
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from fake_useragent import UserAgent    
import pandas as pd
import random
import time
import shutil
import threading
import copy
import os
sem = threading.Semaphore(8)
if os.path.exists(r'C:\Users\user\Desktop\花旗\专利结果'):
    print("已存在文件夹")
else:
    os.mkdir(r'C:\Users\user\Desktop\花旗\专利结果')

try:
    os.system("taskkill /f /im chrome.exe /t")
except:
    pass
try:
    os.system("taskkill /f /im chromedriver.exe /t")
except:
    pass
wait_time_lst=[2**i for i in range(-1,10)]
company_df=pd.read_excel('./需要返工的公司.xlsx')
cnt_list=list(company_df['专利数'])
code_list=list(company_df['证券代码'])
name_list=list(company_df['公司名称'])
class IP: ##创建IP池
    def __init__(self,n=30):
        self.ip_list=[]
        self.main_driver=None
        self.init_url="https://s.wanfangdata.com.cn/patent?q=%E7%94%B3%E8%AF%B7%E4%BA%BA%2F%E4%B8%93%E5%88%A9%E6%9D%83%E4%BA%BA%3A%22%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22&p=19"
        self.ip_num=n
    def is_element_present(self,driver,by, value): ##测试网页元素是否存在
        try:
            element = driver.find_element(by=by, value=value)
        except NoSuchElementException as e:
            return False
        return True
    
    def test_ip(self,ip): ##测试ip是否可用
        options = webdriver_wire.ChromeOptions()
        options.add_argument(ip)
        options.add_argument('--headless')
        driver = webdriver_wire.Chrome(chrome_options= options)
        driver.get(self.init_url)
        if self.is_element_present(driver,By.XPATH,'//*[@id="main-message"]/h1/span'):
            driver.quit()
            return False
        else:
            driver.quit()
            return True
    
    def get_ip(self): ##生成可用的ip池
        print("正在获取ip池...")
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(chrome_options= options)
        driver.get("https://www.kuaidaili.com/free/inha/"+str(random.randint(1,50)))
        result=[]
        while len(result)<=self.ip_num: 
            j=1
            while j<=15:
                if len(result)>self.ip_num:
                    break
                ip=driver.find_element(by=By.XPATH,value='//*[@id="list"]/table/tbody/tr['+str(j)+']/td[1]').text
                port=driver.find_element(by=By.XPATH,value='//*[@id="list"]/table/tbody/tr['+str(j)+']/td[2]').text
                type=driver.find_element(by=By.XPATH,value='//*[@id="list"]/table/tbody/tr['+str(j)+']/td[4]').text.lower()
                new_ip='--proxy-server='+type+'://'+ip+':'+port
                if self.test_ip(new_ip) == True and new_ip not in result:
                    result.append(new_ip)
                    print(len(result))
                j+=1
                print("正在爬取第j个")
        button=driver.find_element(by=By.XPATH,value='//*[@id="listnav"]/ul/li[10]/a').click()
        self.ip_list=result
        driver.quit()
        return
    def create_ip_lst(self):
        self.get_ip()
        return self.ip_list

class GetPatent:
    def __init__(self,ip_lst=[]):
        self.ip_list=ip_lst
        self.main_driver=None
        self.init_url="https://s.wanfangdata.com.cn/patent?q=%E7%94%B3%E8%AF%B7%E4%BA%BA%2F%E4%B8%93%E5%88%A9%E6%9D%83%E4%BA%BA%3A%22%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22&p=19"
        company_df=pd.read_excel(r"C:\Users\user\Desktop\花旗\公司名称.xlsx")
        self.code=None
        self.name=None
        self.result_df=pd.DataFrame(columns=['专利名称','专利摘要','申请/专利权人','公开日期','专利编号'])
        self.whole_num=0
        self.save_path=r'C:\Users\user\Desktop\花旗\专利结果\\'
        self.wait_time_lst=[2**i for i in range(-1,3)]
        self.current_page=0
        self.last_page=0
        self.url="https://s.wanfangdata.com.cn/patent?q=%E7%94%B3%E8%AF%B7%E4%BA%BA%2F%E4%B8%93%E5%88%A9%E6%9D%83%E4%BA%BA%3A%22%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22&p=19"
        self.j=0
        self.whole_num=0
        self.dominar_lst=[]
        self.date_list=[2000+i for i in range(0,23)]
    def clean_process(self):
        try:
            os.system("taskkill /f /im chrome.exe /t")
        except:
            pass
        return
    def is_element_present(self,driver,by, value): ##测试网页元素是否存在
        try:
            driver.find_element(by=by, value=value)
        except NoSuchElementException as e:
            return False
        return True
    def close_driver(self):
        if self.is_element_present(self.main_driver,By.XPATH,'/html/body/div[5]/div/div[1]/div[2]/div/div'):
            return True
        else:
            return False
    def make_driver(self,url,date): ##生成driver
        self.clean_process()
        with open('stealth.min.js', 'r') as f:
            js = f.read()
       # proxy=random.choice(self.ip_list)
        options = webdriver_wire.ChromeOptions()
        #options.add_argument(proxy)
        options.add_argument('-ignore-certificate-errors')
        options.add_argument("--auto-open-devtools-for-tabs")
        options.add_argument('-ignore -ssl-errors')
        options.add_argument('--incognito')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox') # 必要！！
        options.add_argument('--disable-dev-shm-usage') # 必要！！
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver_wire.Chrome(chrome_options= options)   
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {'source': js})
        driver.get(url)
        if date ==2022:
            text='(专利权人:('+self.name+')) and Date:'+str(2022)+'-*'
        else:
            text='(专利权人:('+self.name+')) and Date:'+str(date)+'-'+str(date+1)
        pos=driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[1]/div/div/div/div[1]/div[2]/input')
        ActionChains(driver).click(pos).perform()
        driver.switch_to.active_element.send_keys(Keys.CONTROL+'a')
        driver.switch_to.active_element.send_keys(Keys.BACK_SPACE)
        driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[1]/div/div/div/div[1]/div[2]/input').send_keys(text)##文本内容
        driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[1]/div/div/div/div[1]/div[2]/div/div[2]').click() ## 点击搜索
        time.sleep(random.choice(self.wait_time_lst)*2)
        if self.is_element_present(driver,By.XPATH,'/html/body/div[5]/div/div[1]/div[2]/div/div') ==True:
            print(self.name+'未找到专利！')
            self.main_driver=driver
            return False
        if self.is_element_present(driver,By.XPATH,'/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[4]/div[1]/div/div[3]/div[3]/div/img[2]'):
            pos=driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[4]/div[1]/div/div[3]/div[3]/div/img[2]')
            ActionChains(driver).click(pos).perform()
        pos=driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[1]/div[5]/div[2]/div[1]/span')
        ActionChains(driver).click(pos).perform()
        pos=driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[1]/div[5]/div[2]/div[2]/div[3]')
        ActionChains(driver).click(pos).perform()
        if self.is_element_present(driver,By.XPATH,'/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[4]/div[1]/div/div[3]/div[3]/div/img[2]'):
            pos=driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[4]/div[1]/div/div[3]/div[3]/div/img[2]')
            ActionChains(driver).click(pos).perform()
        self.main_driver=driver
        whole_num=int(driver.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/span/span[2]').text)
        self.whole_num=whole_num
        self.last_page=int(whole_num/50)+1
        return True
    
    def get_current_page_patent(self):
        driver=self.main_driver
        whole_num=self.whole_num
        current_j = str(WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[2]/span[1]')).text)
        self.j=int(current_j[:-1]) - 1
        j=copy.deepcopy(self.j)
        print('爬取第'+str(self.current_page+1)+'页')
        front_path='/html/body/div[5]/div/div[2]/div[2]/div/div[2]/div[2]/div[3]/div['
        
        while True:
            i=1
            self.url=driver.current_url
            if int((j+1)/50)==int(whole_num/50):
                limit=whole_num%50
            else:
                limit=50
            while i<=limit :
                patents_name=WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value=front_path+str(i)+']/div[1]/div[2]/span[2]')).text
                
                if self.is_element_present(driver,By.XPATH,front_path+str(i)+']/div[2]/span[last()]') ==True:
                    
                    patents_public_time=WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value=front_path+str(i)+']/div[2]/span[last()]')).text
                else:
                    patents_public_time=''

                if self.is_element_present(driver,By.XPATH,front_path+str(i)+']/div[3]/span[2]') ==True:
                    patents_abstract=WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value=front_path+str(i)+']/div[3]/span[2]')).text
                else:
                    patents_abstract=''
                    
                if self.is_element_present(driver,By.XPATH,front_path+str(i)+']/div[2]/span[4]/span') ==True:
                    patents_owner=WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value=front_path+str(i)+']/div[2]/span[4]/span')).text
                else:
                    patents_owner=self.name

                if self.is_element_present(driver,By.XPATH,front_path+str(i)+']/div[2]/span[3]') == True:
                    patents_code=WebDriverWait(driver,10,random.random(),(NoSuchElementException,StaleElementReferenceException)).until(lambda x:x.find_element(by=By.XPATH,value=front_path+str(i)+']/div[2]/span[3]')).text
                else:
                    patents_code=''
                self.result_df.loc[j+1]=[patents_name,patents_abstract,patents_owner,patents_public_time,patents_code]
                print(str(j+1)+'/'+str(whole_num)) 
                i+=1
                j+=1    
            self.current_page=int(j/50)+1
            if int(j-1) % 1000 == 0:
                driver.delete_all_cookies()
            if j >= 6000:
                return True
            if int((j-1)/50)!=int(whole_num/50):
                pos=driver.find_element(by=By.CLASS_NAME,value='next')
                ActionChains(driver).click(pos).perform()
                time.sleep(random.choice(self.wait_time_lst))
                print('爬取第'+str(int(j/50)+1)+'页')
            else:
                print('爬取完成！')
                break
        return False
    
    def remake_driver(self,url):
        self.main_driver.quit()
        print("正在清理进程...")
        self.clean_process()
        with open('stealth.min.js', 'r') as f:
            js = f.read()
        #proxy=random.choice(self.ip_list)
        options = webdriver_wire.ChromeOptions()
        options.add_argument('-ignore-certificate-errors')
        options.add_argument("--auto-open-devtools-for-tabs")
        options.add_argument('-ignore -ssl-errors')
        options.add_argument('--incognito')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox') # 必要！！
        options.add_argument('--disable-dev-shm-usage') # 必要！！
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver_wire.Chrome(chrome_options= options)   
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {'source': js})
        driver.get(url)
        time.sleep(random.choice(self.wait_time_lst))      
        self.main_driver=driver
        if self.close_driver():
            return True
        return False

    def run(self,i):
        self.code=list(company_df['证券代码'])[i]
        self.name=list(company_df['公司名称'])[i]
        if os.path.exists(self.save_path+str(self.code)):
            shutil.rmtree(self.save_path+str(self.code))
        os.mkdir(self.save_path+str(self.code))
        self.save_path=self.save_path+str(self.code)
        for date in self.date_list:
            self.current_page=0
            flag_make=False
            while flag_make == False:
                try:
                    flag=self.make_driver(self.url,date)
                    flag_make = True
                except:
                    continue

            if flag:
                while self.current_page !=self.last_page:
                    try:
                        flag=self.get_current_page_patent()
                        if flag==False and self.current_page ==self.last_page:
                            break
                        if flag==True:
                            break
                    except:
                        print("捕获异常,正在重启driver...")
                    if self.remake_driver(self.url) == False:
                        continue
                    else:
                        break
            
            self.result_df=self.result_df.drop_duplicates()
            self.result_df.to_csv(self.save_path+'\\'+self.code+'-'+str(date)+'.csv',index=False)
            self.main_driver.quit()
        return 
    def get_save_path(self):
        return self.save_path
def merge_csv(path,filename):
    name=os.listdir(path)
    name=sorted(name)
    template=pd.read_csv(path+'//'+name[0])
    result_df=pd.DataFrame(columns=template.columns)
    for i in range(len(name)):
        template=pd.read_csv(path+'//'+name[i])
        result_df=pd.concat([result_df,template])
    result_df=result_df.drop_duplicates()
    result_df.to_csv(r'C:\Users\user\Desktop\花旗\新专利结果\\'+filename+'.csv',index=False)
    return

for i in range(len(code_list)):
    if cnt_list[i] >=30000:
        p=GetPatent()
        p.run(i)
        path=p.get_save_path()
        name=code_list[i]
        merge_csv(path,name)

