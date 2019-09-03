import pandas as pd
from pandas import DataFrame
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import re

driver = webdriver.Chrome('chromedriver')

class CleanText():
    "Cleaning texts"
    
    '''cleaning a applicant number'''
    def clean_text_AppNum(self,text):
        cleaned_text = re.sub(' ','',text)
        return cleaned_text[:2] + '-' + cleaned_text[2:6] + '-' + cleaned_text[6:]

    '''cleaning a inventor'''
    def clean_text_nameOfInvent(self,text):
        try:
            n = text.index('(')
            return text[:n]
        except:
            return text   
    
    '''cleaning a ipc'''
    def clean_text_ipc(self,text):
        return text[0]     

    '''cleaning for finding page number'''
    def clean_text_page(self,text):
        cleaned_text = re.sub(' ','',text)
        cleaned_text = re.sub('\(','',cleaned_text)
        cleaned_text = re.sub('\)','',cleaned_text)
        cleaned_text = re.sub('[a-zA-Z]','',cleaned_text)
        return cleaned_text       

    '''cleaning for finding a number of applicant'''
    def clean_text_numofApp(self,text):
        cleaned_text = re.sub(' ','',text)
        return cleaned_text[:2] + '-' + cleaned_text[2:6] + '-' + cleaned_text[6:]

    '''cleaning a name of invent'''
    def clean_text_nameOfInvent(self,text):
        try:
            n = text.index('(')
            return text[:n]
        except:
            return text

    '''cleaning for finding a IPC'''
    def clean_text_ipc(self,text):
        return text[0]

class ReadyforScrape():
    "Open the page & search"

    def __init__(self, search_words):
        self.url = "http://kpat.kipris.or.kr/kpat/searchLogina.do?next=MainSearch"
        self.search_words = search_words

        # kipris 접속
        driver.get(self.url)
        

    '''input the search words'''
    def Searching(self):
        
        # 키워드 입력
        driver.find_element_by_id('queryText').send_keys(self.search_words)

        # 검색
        driver.find_element_by_xpath('//*[@id="SearchPara"]/fieldset/span[1]/a').click()

class FindByXpath():
    "Click the button by xpath"

    global driver

    def __init__(self, xpath):
        self.xpath = xpath
    
    def ClickButton(self):
        driver.find_element_by_xpath(self.xpath).click()

    def GetText(self):
        txt = driver.find_element_by_xpath(self.xpath).text
        return txt
        
class Scraping():
    "Get Id List from the Current Page"

    def __init__(self):
        self.article_id = []
        self.ct = CleanText()

    '''Get Pages'''
    def CountPages(self):
        
        # 총 페이지 수 긁어오기
        FBX = FindByXpath('//*[@id="divMainArticle"]/form/section/div[1]/p')
        pageNum = self.ct.clean_text_page(FBX.GetText())
        n = pageNum.index('/')

        endPageNum = int(pageNum[n+1:])

        print("Lasg Page: ", endPageNum)

        return endPageNum

    
    '''Get Ids from Article for First Time'''
    def GetIdList_FirstTime(self):
     
        while True:

            # Article ID 수집
            html = driver.page_source
            soup = BeautifulSoup(html, 'lxml')
            lines = soup.select('section.search_section > article')

            for line in lines:
                self.article_id.append(line.get('id'))

            # id 전부 수집했는지 확인
            if len(self.article_id) == 90:
                self.CountPages()
                return self.article_id
            else:
                self.article_id.clear()


    '''Get Ids from Article not for Frist Time '''
    def GetIdList_NotFirst(self):

        temp_id = self.article_id[len(self.article_id)-1]   
        article_id_temp = [] 
        ch = 1

        while True:

            article_id_temp.clear()
    
            # Article ID 수집
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            lines = soup.select('section.search_section > article')

            for line in lines:
                article_id_temp.append(line.get('id'))

            if article_id_temp[len(article_id_temp)-1] != temp_id:
                self.article_id.extend(article_id_temp)
                return 0

            ch = ch + 1
            if ch >= 20:
                return ch

    '''Get Data about Applicant from Pages''' 
    def GetData(self, i):

        # 출원인
        applicant_xpath = '//*[@id="' + self.article_id[i] + '"]/div[2]/ul/li[4]/a'
        
            # 페이지 전환될 때까지 기다리기
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, applicant_xpath)))
            applicant_data = driver.find_element_by_xpath(applicant_xpath).text
        except:
            print("출원인 정보 미오픈 전 스크랩 시도")
            applicant_data = '-'
        
        
        # 출원번호
        numOfApp_xpath = '//*[@id="' + self.article_id[i] + '"]/div[2]/ul/li[2]/span[2]/a'
        numOfApp_data = self.ct.clean_text_numofApp(driver.find_element_by_xpath(numOfApp_xpath).text)

        # 발명의 명칭
        nameOfInvent_xpath = '//*[@id="' + self.article_id[i] + '"]/div[1]/h1/a[2]'
        
            # 페이지 전환될 때까지 기다리기
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, nameOfInvent_xpath)))
            nameOfInvent_data = self.ct.clean_text_nameOfInvent(driver.find_element_by_xpath(nameOfInvent_xpath).text)
        except:
            print("발명의 명칭 정보 미오픈 전 스크랩 시도")
            nameOfInvent_data = '-'
        
         
        # 요약
        summary_xpath = '//*[@id="' + self.article_id[i] + '"]/div[2]/div[2]/div'
        summary_data = driver.find_element_by_xpath(summary_xpath).text

        # IPC
        ipc_xpath = '//*[@id="' + self.article_id[i] + '"]/div[2]/ul/li[1]/span[2]/a'
        ipc_data = driver.find_element_by_xpath(ipc_xpath).text
        # ipc_data = self.ct.clean_text_ipc(driver.find_element_by_xpath(ipc_xpath).text)
  
        return applicant_data, numOfApp_data, nameOfInvent_data, summary_data, ipc_data

class ChangePage():
    "Over the Page"


    '''처음 페이지를 넘기는 경우'''
    def FirstTime(self,nextPage):

        page_xpath = '//*[@id="divMainArticle"]/form/section/div[2]/span/a[' + str(nextPage-1) + ']'
        driver.find_element_by_xpath(page_xpath).click()

        print("Go to Next Page : next page =", nextPage, "p")


    '''처음이 아닌 경우'''
    def NotFirstTime(self,nextPage):

        remainNum = int(nextPage % 10)
        
        if remainNum == 1:
            page_xpath = '//*[@id="divMainArticle"]/form/section/div[2]/span/a[11]'
            
        elif remainNum == 0:
            page_xpath = '//*[@id="divMainArticle"]/form/section/div[2]/span/a[10]'
            
        else: 
            page_xpath = '//*[@id="divMainArticle"]/form/section/div[2]/span/a[' + str(remainNum) + ']'
            
        # 검색 완료될 때까지 기다리기
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, page_xpath)))
            driver.find_element_by_xpath(page_xpath).click()
        except:
            print("검색 전 스크랩 시도")

        print("Go to Next Page : nextPage =", nextPage, "p")
              
class ExcelWriter():
    "Write on Excel File"

    def __init__(self, excel_file_name):
        self.writer = pd.ExcelWriter(excel_file_name)


    '''Columns: ['출원번호', '발명의 명칭', '출원인', '요약', 'IPC']'''
    def writeOnExcel(self, numofApp, nameOfInvent, applicant, summary, ipc):
        self.df = DataFrame({'출원번호':numofApp, '발명의 명칭':nameOfInvent, '출원인':applicant, '요약':summary, 'IPC':ipc})
        self.df.to_excel(self.writer, sheet_name='Sheet')
        self.writer.save()
        print("Write Data on Excel")


'''Start Scrping'''
if __name__ == '__main__':

    # 검색식 입력
    SearchFormula = input("검색식을 입력하세요: ")    #TL=[shackle] or RTU*모니터   etc...
    
    # 키프리스 검색
    Ready = ReadyforScrape(SearchFormula)
    Ready.Searching()

    # 검색 완료될 때까지 기다리기
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "divBoardPager")))
    except:
        print("검색 전 스크랩 시도")

    # 모드 전환 - 페이지당 90개 / 요약함께보기 
    FindByXpath('//*[@id="opt28"]/option[3]').ClickButton()
    FindByXpath('//*[@id="pageSel"]/a').ClickButton()
    FindByXpath('//*[@id="btnTextView"]/button').ClickButton()

    # 모드 전환될 때까지 기다리기
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "divBoardPager")))
    except:
        print("모드 전환 전 스크랩 시도")

    # ID 수집
    scrape = Scraping()
    IDs = scrape.GetIdList_FirstTime()
    
    # 전체 페이지 체크
    EndPage = scrape.CountPages()

    # 데이터 변수 생성
    Applicant =[] 
    AppNum = [] 
    InventName = [] 
    Summary = [] 
    IPC = [] 

    try:
        
        for nowPage in range(1, EndPage+1):
            length = len(IDs)
            start_index = (nowPage - 1) * 90

            for i in range(start_index, length):
                # 각 ID마다 Data 가져오기
                a,b,c,d,e = scrape.GetData(i)
                # 가져온 Data 스택
                Applicant.append(a)
                AppNum.append(b)
                InventName.append(c)
                Summary.append(d)
                IPC.append(e)
                # 현재 번호 출력
                print("i:", i+1)
    
            # 마지막 페이지인지 확인
            if nowPage == EndPage: break
            
            # 다음 페이지로 이동
            n = int(nowPage)
            n += 1
            print("next page: ", n)

            if n < 12 : 
                ChangePage().FirstTime(n)
            else :
                ChangePage().NotFirstTime(n)

            # 페이지가 완전히 넘어갈 때까지 대기
            ch = scrape.GetIdList_NotFirst()
            if ch != 0: break

    finally:
        # 엑셀 파일 이름 변수
        FileName = "Shackle.xlsx"
        
        # 엑셀에 데이터 저장
        writer = ExcelWriter(FileName)
        writer.writeOnExcel(AppNum, InventName, Applicant, Summary, IPC)


    
