from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import time
import random
now = time.localtime()
nowdate = "%04d.%02d.%02d" % (now.tm_year, now.tm_mon, now.tm_mday)
s = "%04d.%02d.%02d_%02d%02d%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
def saveExcel(data):
    """
    :param query: naver 쇼핑물에 query 한걸로 파일명 을 만들꺼
    :param data:  크롤링한 결과물
    """
    # 엑셀시트 header 설정 및, 열의 넓이 설정
    header1 = ['상호명','TEL','HP','url']
    wb = Workbook()
    ws1 = wb.worksheets[0]
    ws1.column_dimensions['A'].width = 50
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 30
    ws1.column_dimensions['D'].width = 30
    ws1.append(header1)
    # 데이터 삽입
    # itemlist 가 [품명,최저가,링크] 이런식으로 온걸
    # openpyxl 객체 ws1 에 append 시키면 들어감
    for itemlist in data:
        ws1.append(itemlist)
    wb.save(s+query+"_알바천국.xlsx")

if __name__ == "__main__": # 직접실행시키는 경우
    try:
        driver = webdriver.Chrome('./chromedriver')
        driver.maximize_window()
        with open("option.txt") as f:
            lines = f.readlines()
            query = lines[7].split(':')[-1].strip()
            standDate = lines[8].split(':')[-1].strip()
            endDate = lines[9].split(':')[-1].strip()
            delay = int(lines[10].split(':')[-1].strip())
            print(">>> 설정 파일 불러오기 성공 ")

        last = int(input(">>> *만약 임시저장으로 끊겼으면 엑셀마지막왼쪽의 행번호를 입력 (처음시작은 1):"))
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
        IsGo = False
        IsWhileGo = True
        bs4 = BeautifulSoup(driver.page_source, 'lxml')
        total = bs4.find('div', id='SearchJob').find('em').get_text().strip()
        print(">>> 총 ", total, "개 게시글 존재 ")
        datalist = []
        datalist.clear()
        listidx = 1
        while IsWhileGo:
            btns = driver.find_elements_by_class_name('summaryBtn')
            for btn in btns:
                while True:
                    try:
                        time.sleep(0.5)
                        btn.click()
                        break
                    except:
                        time.sleep(0.5)
            time.sleep(2)
            bs4 = BeautifulSoup(driver.page_source, 'lxml')
            lis = bs4.find('ul',id='jobNormal').find_all('li')
            smbs = bs4.find_all('div',class_='summaryView')
            summaryidx = 0
            alist = []
            alist.clear()
            titlelist = []
            titlelist.clear()
            for li in lis:
                title = li.find('span',class_='company').get_text().strip()
                href = baseurl+li.find('a')['href']
                dateday = li.find('span',class_='regDate').get_text().split(':')[1].strip()
                if title not in titlelist:
                    if dateday == '':
                        dateday = nowdate
                    if endDate >= dateday and dateday >= standDate:
                        try:
                            html2 = requests.get(baseurl + smbs[summaryidx].find('iframe')['src'])
                            bs42 = BeautifulSoup(html2.text, 'lxml')
                            telhps = bs42.find('td', class_='telHtel').find_all('em')
                        except:
                            pass
                        tel = ''
                        hp = ''
                        for tmp in telhps:
                            try:
                                if 'HP' in tmp.get_text().strip():
                                    hp = tmp.get_text().split('HP.')[1].strip()
                                else:
                                    tel = tmp.get_text().split('Tel.')[1].strip()
                            except:
                                print("연락처 가져오기 에러",baseurl + smbs[summaryidx].find('iframe')['src'])
                                pass
                        datalist.append([title,tel,hp,baseurl + smbs[summaryidx].find('iframe')['src']])
                        print(listidx, ' 개 추출중..')
                        listidx+=1
                        print(baseurl + smbs[summaryidx].find('iframe')['src'])
                        summaryidx += 1
                        alist.append(href)
                        titlelist.append(title)
                        IsGo = True #한번들어오면
                    else:
                        if IsGo: #한번이라도 들어간상태에였으면
                            IsWhileGo = False
                            break
            pageidx += 1
            if pageidx == 11:
                try:
                    driver.find_element_by_xpath('//*[@id="SearchJob"]/div/div/span[3]/a').click()
                except:
                    driver.find_element_by_xpath('//*[@id="SearchJob"]/div/div/span[4]/a').click()
                pageidx = 1
            else:
               try:
                    driver.find_element_by_xpath('//*[@id="SearchJob"]/div/div/span[1]/a['+str(pageidx)+']').click()
               except:
                   driver.find_element_by_xpath('//*[@id="SearchJob"]/div/div/span[2]/a[' + str(pageidx) + ']').click()
            time.sleep(3)
            bs4 = BeautifulSoup(driver.page_source, 'lxml')
        print(">>> 조건에 맞는 게시글 추출완료 1")
        saveExcel(datalist)
    except:
        print(">>> 조건에 맞는 게시글 추출완료 2")
        saveExcel(datalist)
    finally:
        driver.quit()