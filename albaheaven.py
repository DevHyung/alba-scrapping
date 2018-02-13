from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import time
import random
def saveExcel(data):
    """
    :param query: naver 쇼핑물에 query 한걸로 파일명 을 만들꺼
    :param data:  크롤링한 결과물
    :return:  NONE
    """
    # 엑셀시트 header 설정 및, 열의 넓이 설정
    header1 = ['상호명','담당자','이메일','TEL','HP','url']
    wb = Workbook()
    ws1 = wb.worksheets[0]
    ws1.column_dimensions['A'].width = 50
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 30
    ws1.column_dimensions['D'].width = 30
    ws1.column_dimensions['E'].width = 30
    ws1.append(header1)
    # 데이터 삽입
    # itemlist 가 [품명,최저가,링크] 이런식으로 온걸
    # openpyxl 객체 ws1 에 append 시키면 들어감
    for itemlist in data:
        ws1.append(itemlist)
    wb.save(query+"_알바천국.xlsx")

if __name__ == "__main__": # 직접실행시키는 경우
    driver = webdriver.Chrome('./chromedriver')
    driver.maximize_window()
    with open("option.txt") as f:
        lines = f.readlines()
        query = lines[5].split(':')[-1].strip()
        standDate = lines[6].split(':')[-1].strip()
        print(">>> 설정 파일 불러오기 성공 ")
    baseurl = 'http://www.alba.co.kr'
    driver.get('http://www.alba.co.kr/search/Search.asp?WsSrchWord=&wsSrchWordarea=&Section=0&Page=&hidschContainText=&hidWsearchInOut=&hidGroupKeyJobArea=&hidGroupKeyJobHotplace=&hidGroupKeyJobJobKind=&hidGroupKeyResumeArea=&hidGroupKeyResumeJobKind=&hidGroupKeyPay=&hidGroupKeyWorkWeek=&hidGroupKeyWorkPeriod=&hidGroupKeyOpt=&hidGroupKeyGender=&hidGroupKeyAge=&hidGroupKeyCareer=&hidGroupKeyLicense=&hidGroupKeyEduData=&hidGroupKeyWorkTime=&hidGroupKeyWorkState=&hidGroupKeyJobCareer=&hidSort=&hidSortOrder=1&hidSortDate=&hidSortCnt=&hidSortFilter=&hidArea=&area=&hidJobKind=&jobkind=&gendercd=C03&ageconst=G01&agelimitmin=&agelimitmax=&workperiod=&workweek=')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="wsSrchWord"]').send_keys(query)
    driver.find_element_by_xpath('//*[@id="wsSchForm"]/div/button').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="SearchJob"]/div/p[1]/span/a[2]').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="pagesize"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="pagesize"]/option[3]').click()
    time.sleep(3)
    pageidx = 1
    IsGo = True
    bs4 = BeautifulSoup(driver.page_source, 'lxml')
    total = bs4.find('div', id='SearchJob').find('em').get_text().strip()
    print(">>> 총 ", total, "개 게시글 존재 ")
    while IsGo:
        btns = driver.find_elements_by_class_name('summaryBtn')
        for btn in btns:
            while True:
                try:
                    btn.click()
                    break
                except:
                    time.sleep(0.3)
            # 자이제 날짜맞춰서 가져오는부분
            # 페이지 이동해서 가져오는부분
            # 요약정보따오고
            # 페이지들어가서 가져올 필요가잇는지
        bs4 = BeautifulSoup(driver.page_source, 'lxml')
        lis = bs4.find('ul',id='jobNormal').find_all('li')
        smbs = bs4.find_all('div',class_='summaryView')
        summaryidx = 0
        alist = []
        titlelist = []
        for li in lis:
            title = li.find('span',class_='company').get_text().strip()
            href = baseurl+li.find('a')['href']
            dateday = li.find('span',class_='regDate').get_text().split(':')[1].strip()
            try:
                print(smbs[summaryidx].find('iframe')['src'])
                print(baseurl + smbs[summaryidx].find('iframe')['src'])
                html2 = requests.get(baseurl+smbs[summaryidx].find('iframe')['src'])
                bs42 = BeautifulSoup(html2.text, 'lxml')
                print(bs42.prettify())
                #telhps = smbs[summaryidx].find('td',class_='telHtel').find_all('em')
            except:
                print(summaryidx)
                print(smbs[summaryidx].find('td',class_='telHtel').find_all('em'))
            tel = ''
            hp = ''
            for tmp in telhps:
                if 'HP' in tmp.get_text().strip():
                    hp = tmp.get_text().split('HP.')[1].strip()
                else:
                    tel = tmp.get_text().split('Tel.')[1].strip()
            if title not in titlelist:
                if dateday == '' or dateday >= standDate:
                    print(title, href, dateday)
                    print("hp:",hp,"tel:",tel)
                    alist.append(href)
                    titlelist.append(title)
                else:
                    IsGo = False
                    break
            summaryidx +=1

        pageidx += 1
        driver.find_element_by_xpath('//*[@id="SearchJob"]/div/div/span[1]/a['+str(pageidx)+']').click()
        time.sleep(3)
        bs4 = BeautifulSoup(driver.page_source, 'lxml')
    #'//*[@id="SearchJob"]/div/div/span[3]/a'
    print(">>> 조건에 맞는 ", len(alist), "개 게시글 필터링 ")
    alistidx = 1
    datalist = []
    for a in alist:
        tmplist = []
        driver.get(a)
        time.sleep(random.randint(5,8))
        try:
            driver.find_element_by_xpath('//*[@id="allcontent"]/div[2]/div[5]/div[1]/div/a').click()
        except:
            pass
        print(">>> ",alistidx,'번째 추출중..')
        bs42 = BeautifulSoup(driver.page_source,'lxml')
        title = bs42.find('span',class_='companyName').get_text().strip()
        try:
            driver.switch_to_frame(driver.find_elements_by_tag_name('iframe')[3])
            name = driver.find_element_by_xpath('/html/body/div/div[1]/span[2]').text
            email = driver.find_element_by_xpath('/html/body/div/div[2]/span[2]/a').text
            phone = driver.find_element_by_xpath('/html/body/div/div[3]/span[2]/div/span').text
        except:
            input("!!! 캡챠 발생 정지 !!! 캡챠를 푼후 엔터키를 눌러주세요 ::")
            driver.switch_to_frame(driver.find_elements_by_tag_name('iframe')[3])
            name = driver.find_element_by_xpath('/html/body/div/div[1]/span[2]').text
            email = driver.find_element_by_xpath('/html/body/div/div[2]/span[2]/a').text
            phone = driver.find_element_by_xpath('/html/body/div/div[3]/span[2]/div/span').text
        alistidx+=1
        datalist.append([title,name,email,phone,a])
    saveExcel(datalist)