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
    header1 = ['상호명','담당자','이메일','전화번호','url']
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
    wb.save(query+"_알바몬.xlsx")

if __name__ == "__main__": # 직접실행시키는 경우
    driver = webdriver.Chrome('./chromedriver')
    driver.maximize_window()
    with open("option.txt") as f:
        lines = f.readlines()
        query = lines[1].split(':')[-1].strip()
        standDate = lines[2].split(':')[-1].strip()
        print(">>> 설정 파일 불러오기 성공 ")
    baseurl = 'http://www.albamon.com'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    html = requests.get('http://www.albamon.com/search/Recruit?Keyword='+query+'&IsExcludeDuplication=True&PageSize=1000&OrderType=1', headers=headers)
    bs4 = BeautifulSoup(html.text,'lxml')
    total = bs4.find('span',class_='total').find('em').get_text().strip()
    print(">>> 총 ",total,"개 게시글 존재 ")
    divs = bs4.find_all('div',class_='booth')
    alist = []
    for div in divs:
        href  = 'http://www.albamon.com'+div.find('dt').find('a')['href']
        dateday = div.find('span',class_='regTime').get_text().strip()
        if dateday == '' or dateday >= standDate:
            alist.append(href)
        else:
            break
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