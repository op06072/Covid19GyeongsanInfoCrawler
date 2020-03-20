from urllib.request import urlopen
from bs4 import BeautifulSoup
import openpyxl as xl
import os
import firebase_admin
from firebase_admin import credentials
from firebase_admin import db
from selenium import webdriver
import requests
import re

path = os.path.dirname(__file__)
wb = xl.load_workbook(path+"/코로나데이터.xlsx", data_only=True)
driver = webdriver.Chrome(path+"/chromedriver")

status = wb["발생 현황"]
route = wb["이동 경로"]
clinic = wb["선별진료소"]
mask = wb["마스크 판매 약국"]
xy_url = "http://maps.googleapis.com/maps/api/geocode/xml?address="
# 카카오맵 api주소
building = 'https://dapi.kakao.com/v2/local/search/keyword.json?query=' # 건물명 검색
address = 'https://dapi.kakao.com/v2/local/search/address.json?query=' # 주소 검색
headers = { "Authorization": "KakaoAK 8d63652d146dc5a3a958048e957ba061"}


""" for i in bsObject[3::2]:
    Where = []
    When = {}
    I=i.tbody.contents
    info = I[1].contents
    Info = {"인적사항" : info[5].contents[0]} #인적사항 : 성별(거주하는 동, 나이)
    day = info[9].contents[0].split(".")
    Info["확진일자"] = day[0]+"월 "+day[1]+"일"
    Info["입원기관"] = info[11].contents[0]
    Info["접촉자 수(격리조치 중 접촉자 수)"] = info[13].contents[0]
    number = int(info[1].contents[0])
    where = I[3].contents[1].contents
    for j in where[1::2]:
        j = j.contents[1].contents[0]
        if j == "경로 확인중":
            When["None"] = ["경로 확인 중"]
            Info["이동 경로"] = When
            Patients[number] = Info
            break
        elif j == "주요동선(직장,병의원,약국)없음":
            When["None"] = [j]
            Info[""]

        if j[0] is "※": # 여기 작업중
            when = j.split("/")
            month = int(when[0])
            when = when[1].split(" ")
            date = int(when[0])
            do = when[1] """

def addresslocation(PLACE): # 주소로 좌표획득
    if "," in PLACE:
        PLACE = PLACE.split(",")[0]
    location = requests.get(address+PLACE, headers = headers).json()['documents'][0]
    return [float(location['x']), float(location['y'])]

def buildinglocation(PLACE): # 건물명으로 좌표획득
    if "농산물직판장" in PLACE:
        PLACE = PLACE.split("직판장")[0]
    PLACE = "".join(PLACE.split(" "))
    location = requests.get(building+PLACE, headers = headers).json()['documents'][0]
    return [float(location['x']), float(location['y'])]

def routeaddress(string):
    strlist = " ".join(re.findall("\s+", string))
    #ban = [':','(',')','월','화','수','목','금','토','일',' ']
    if "동)" in strlist:
        strlist = strlist.split("동)")[0]+"동)"
    elif "읍)" in strlist:
        strlist = strlist.split("읍)")[0]+"읍)"
    elif "면)" in strlist:
        strlist = strlist.split("면)")[0]+"면)"
    elif "방문" in strlist:
        strlist = strlist.split("방문")[0]
    elif "점" in strlist:
        strlist = strlist.split("점")[0]+"점"
    searchstr = '경산 ' + strlist
    return buildinglocation(searchstr)

# 확진자 이동경로 크롤링
def movingroute():
    driver.get("http://www.gbgs.go.kr/programs/coronaMove/coronaMove.do")
    #html = driver.find_element_by_xpath('html/body').find_elements_by_xpath('div')[2].find_element_by_xpath('div/div/div').find_elements_by_xpath('table')
    html = driver.page_source
    bsObject = BeautifulSoup(html, "html.parser").find_all("table")
    Patients = {}
    row = 1
    for i in bsObject:
        I = i.tbody
        Info=I.tr
        info = []
        I = I.contents
        for j in Info:
            t=str(type(j))
            if "bs4.element.NavigableString" not in t and "bs4.element.Comment" not in t:
                info.append(j)
        route.cell(row, 1, int(info[0].contents[0])) # 경산시 내부 번호
        if info[1].contents[0] == 0:
            route.cell(row, 2, "미정") # 확진 번호
        else:
            route.cell(row, 2, int(info[1].contents[0])) # 확진 번호
        route.cell(row, 3, info[2].contents[0]) # 인적사항 : 성별(거주하는 동, 나이)
        day = info[3].contents[0].split(".")
        route.cell(row, 4, day[0]+"월 "+day[1]+"일")
        if len(info[4]) == 0:
            route.cell(row, 5, "배정요청")
        elif len(info[4].contents) > 1:
            center = []
            for j in info[4].contents:
                t=str(type(j))
                if "bs4.element.NavigableString" in t or "bs4.element.Comment" in t:
                    center.append(j)
            route.cell(row, 5, "\n".join(center))
        else:
            route.cell(row, 5, info[4].contents[0])
        route.cell(row, 6, info[5].contents[0])
        where = I[3].find_all("li")
        extra = []
        for j in where:
            if j.contents[0] != "\n":
                j = j.contents[0]
            else:
                continue
            if j == "경로 확인중":
                route.cell(row, 8, "경로 확인중")
            elif j == "주요동선(직장,병의원,약국)없음":
                route.cell(row, 8, "주요동선(직장,병의원,약국)없음")
            else:
                if "※" in j:
                    if j[0] == "※":
                        extra.append(j)
                        continue
                    else:
                        j = j.split("※")
                        extra.append("※"+j[1])
                        j = j[0]
                jlist = [j]
                if "\n" in j:
                    jlist = j.split("\n")
                for k in jlist:
                    if k == "":
                        continue
                    row += 1
                    if "퇴원" not in k and "사망" not in k and "격리" not in k and "검사" not in k and "자택" not in k and "신천지" not in k and "직장" not in k and "양성" not in k:
                        locate = routeaddress(k)
                        route.cell(row, 9, float(locate[0]))
                        route.cell(row, 10, float(locate[1]))
                    route.cell(row, 8, k)
                    #route.cell(row, 8, Route[0])
                    #route.cell(row, 9, Route[1])
                for k in extra:
                    row += 1
                    route.cell(row, 11, k)
        row += 1
    wb.save(path+"/코로나데이터.xlsx")

def occurrence():
    html = urlopen("http://www.gbgs.go.kr/programs/corona/corona.do")
    gs = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.contents # 경산데이터
    status.cell(1, 1, gs[5].contents[0]) # 업데이트 시간
    info = gs[7].contents
    patients = info[1].ul.contents # 확진자 데이터
    diagnosis = info[3].ul.contents # 의심환자 데이터
    status.cell(2, 1, int("".join(patients[1].span.contents[0].split(",")))) # 확진자 수
    status.cell(3, 1, int("".join(diagnosis[3].span.contents[0].split(",")))) # 검사중인 환자 수
    status.cell(4, 1, int("".join(patients[5].span.contents[0].split(",")))) # 퇴원자 수
    status.cell(5, 1, int("".join(diagnosis[5].span.contents[0].split(",")))) # 음성인 환자 수
    wb.save(path+"/코로나데이터.xlsx")

def maskinfo():
    html = urlopen("http://www.gbgs.go.kr/design/health/COVID19/COVID19_05_02.html") # 경산시 마스크 공적판매처
    OfficialMask = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.div.contents
    OfficialMask =[OfficialMask[9].table.tbody, OfficialMask[15].table.tbody]
    Dong = []
    dong = OfficialMask[0].contents
    for i in dong:
        t=str(type(i))
        if "bs4.element.NavigableString" not in t and "bs4.element.Comment" not in t:
            Dong.append(i)
    EupMyeon = []
    eupmyeon = OfficialMask[1].contents
    for i in eupmyeon:
        t=str(type(i))
        if "bs4.element.NavigableString" not in t and "bs4.element.Comment" not in t:
            EupMyeon.append(i)
    row = 1
    for i in Dong: # 동지역 판매처
        I = []
        for j in i:
            t=str(type(j))
            if "bs4.element.NavigableString" not in t and "bs4.element.Comment" not in t:
                I.append(j)
        if len(I) == 3:
            mask.cell(row, 1, I[0].contents[0].split("\\")[0]) # 분류
            mask.cell(row, 2, I[1].contents[0]) # 판매처 이름
            mask.cell(row, 3, I[2].contents[0]) # 전화번호
            mask.cell(row, 4, "공적판매처") # 공사구분
            mask.cell(row, 5, "동") # 동 읍/면 구분
        else:
            mask.cell(row, 2, I[0].contents[0]) # 판매처 이름
            mask.cell(row, 3, I[1].contents[0]) # 전화번호
            mask.cell(row, 4, "공적판매처") # 공사구분
            mask.cell(row, 5, "동") # 동 읍/면 구분
        row += 1
    for i in EupMyeon: # 읍면지역 판매처
        I = []
        for j in i:
            t=str(type(j))
            if "bs4.element.NavigableString" not in t and "bs4.element.Comment" not in t:
                I.append(j)
        if len(I) == 3:
            mask.cell(row, 1, I[0].contents[0].split("\\")[0]) # 분류
            place = I[1].contents[0] # 판매처 이름
            mask.cell(row, 3, I[2].contents[0]) # 전화번호
        else:
            place = I[0].contents[0] # 판매처 이름
            mask.cell(row, 3, I[1].contents[0]) # 전화번호
        if "우체국" in place:
            mask.cell(row, 2, "경산"+place)
        elif "농협" in place:
            if "하나로마트" in place:
                Place = place.split("하나로마트")
                if Place[1] == "":
                    if " " in Place[0]:
                        Place[0] = Place[0].split(" ")
                        if Place[0][0] != "":
                            Place[0] = Place[0][0]
                        else:
                            Place[0] = Place[0][1]
                    if "농협" in Place[0]:
                        Place = Place[0].split("농협")
                    if "지점" in Place[1]:
                        Place[1] = Place[1].split("지점")[0]+"점"
                    mask.cell(row, 2, Place[0]+"농협 하나로마트 "+Place[1])
                elif Place[0] == "":
                    if " " in Place[1]:
                        Place[1] = Place[1].split(" ")
                        if Place[1][0] != "":
                            Place[1] = Place[1][0]
                        else:
                            Place[1] = Place[1][1]
                    if "농협" in Place[1]:
                        Place = Place[1].split("농협")
                    if "지점" in Place[1]:
                        Place[1] = Place[1].split("지점")[0]+"점"
                    mask.cell(row, 2, Place[0]+"농협 하나로마트 "+Place[1])
                else:
                    if " " in Place[0]:
                        Place[0] = Place[0].split(" ")
                        if Place[0][0] != "":
                            Place[0] = Place[0][0]
                        else:
                            Place[0] = Place[0][1]
                    if "농협" in Place[0]:
                        Place1 = Place[0].split("농협")
                    if "지점" in Place[1]:
                        Place[1] = Place[1].split("지점")[0]+"점"
                    mask.cell(row, 2, Place1[0]+"농협 하나로마트 "+Place[1])
            else:
                mask.cell(row, 2, place)
        mask.cell(row, 4, "공적판매처") # 공사구분
        mask.cell(row, 5, "읍/면") # 동 읍/면 구분
        row += 1
    html = urlopen("http://www.gbgs.go.kr/design/health/COVID19/COVID19_05_03.html") # 경산시 마스크 판매 약국
    pharmacy = BeautifulSoup(html, "html.parser").body.contents[5].div.div.table.tbody.contents
    for i in pharmacy[1::2]:
        i = i.contents
        mask.cell(row, 1, i[3].contents[0]) # 읍면동
        mask.cell(row, 2, i[5].contents[0]) # 판매처이름
        mask.cell(row, 3, i[9].contents[0]) # 전화번호
        mask.cell(row, 4, "약국") # 공사구분
        mask.cell(row, 5, i[7].contents[0]) # 주소
        row += 1
    driver.get("http://www.gbgs.go.kr/design/health/COVID19/COVID19_05_05.html") # 경산시 판매 현황
    html = driver.page_source
    stock = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.div.table.tbody.contents
    for i in stock[1::]:
        row = 1
        i = i.contents
        name = "".join(i[0].contents[0].split("본점")) # 판매처 이름
        name = "".join(name.split(" "))
        if mask.cell(row, 2).value != None:
            Name = "".join(mask.cell(row, 2).value.split(" "))
        else:
            Name = None
        while Name not in [None,name]: # 판매처 인덱스 찾기
            row += 1
            if mask.cell(row, 2).value != None:
                Name = "".join(mask.cell(row, 2).value.split(" "))
            else:
                row -= 1
                Name = None
        if len(i[2].contents) == 0:
            Stock = "없음"
        else:
            Stock = i[2].contents[0]
        if "(" in Stock:
            Stock = Stock.split(" (")
            Stock[1] = "(" + Stock[1]
        else:
            Stock = [Stock, ""]
        mask.cell(row, 6, Stock[0]) # 재고 수준
        mask.cell(row, 7, Stock[1]) # 재고량
        if len(i[4].contents):
            renewaltime = i[4].contents[0]
        mask.cell(row, 8, renewaltime) # 재고량 갱신 시간
    row = 0
    for i in mask.rows:
        row += 1
        if i[1].value == None:
            break
        elif i[5].value == None:
            i[5].value = "집계중"
            i[7].value = renewaltime
        if i[3].value == "공적판매처":
            place = i[1].value
            #if "하나로마트" in place:
                #if place[-1] != "점":
                #    place += "본점"
            locate = buildinglocation(place)
            mask.cell(row, 9, locate[0])
            mask.cell(row, 10, locate[1])

        elif i[3].value == "약국":
            place = "경산시 "+i[4].value
            locate = addresslocation(place)
            mask.cell(row, 9, locate[0])
            mask.cell(row, 10, locate[1])

    wb.save(path+"/코로나데이터.xlsx")

def clinicinfo(): # 선별진료소 데이터 크롤링
    html = urlopen("http://www.gbgs.go.kr/design/health/COVID19/COVID19_04.html") # 선별 진료소 현황
    Clinic = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.table.tbody.contents
    row = 1
    for i in Clinic[1::2]:
        i = i.contents
        clinic.cell(row, 1, i[1].contents[0]) # 기관
        clinic.cell(row, 2, i[3].contents[0]) # 주소
        call = []
        for j in i[5].contents:
            t=str(type(j))
            if "bs4.element.NavigableString" in t or "bs4.element.Comment" in t:
                j = j.split("\r")
                j = j[-1].split("\n")
                j = j[-1].split("\t")[-1]
                call.append(j)
        clinic.cell(row, 3, " \r\n".join(call)) # 전화번호
        time = []
        for j in i[7].contents:
            t=str(type(j))
            if "bs4.element.NavigableString" in t or "bs4.element.Comment" in t:
                j = j.split("\r")
                j = j[-1].split("\n")
                j = j[-1].split("\t")[-1]
                time.append(j)
        clinic.cell(row, 4, " \r\n".join(time)) # 진료시간
        if len(i) > 9:
            extra = []
            I = i[9].contents
            if len(I) == 1 and "bs4.element.Tag" in str(type(I[0])):
                I = I[0].contents
            for j in I:
                t=str(type(j))
                if "bs4.element.NavigableString" in t or "bs4.element.Comment" in t:
                    j = j.split("\r")
                    j = j[-1].split("\n")
                    j = j[-1].split("\t")[-1]
                    extra.append(j)
            clinic.cell(row, 5, " \r\n".join(extra))
        else:
            extra = clinic.cell(row-1, 5).value
            clinic.cell(row, 5, extra)
        locate = addresslocation("경산시 " + clinic.cell(row, 2).value)
        clinic.cell(row, 6, locate[0])
        clinic.cell(row, 7, locate[1])
        row += 1
    wb.save(path+"/코로나데이터.xlsx")

def crawler():
    movingroute()
    occurrence()
    clinicinfo()
    maskinfo()

crawler()