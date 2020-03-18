from urllib.request import urlopen
from bs4 import BeautifulSoup
import openpyxl as xl
import os

path = os.path.dirname(__file__)
wb = xl.load_workbook(path+"/코로나데이터.xlsx", data_only=True)

status = wb["발생 현황"]
route = wb["이동 경로"]
clinic = wb["선별진료소"]
mask = wb["마스크 판매 약국"]

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

# 확진자 이동경로 크롤링
def movingroute():
    html = urlopen("http://www.gbgs.go.kr/programs/coronaMove/coronaMove.do")
    bsObject = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.contents
    Patients = {}
    row = 1
    for i in bsObject[3::2]:
        I = i.tbody.contents
        info=I[1].contents
        route.cell(row, 1, int(info[1].contents[0])) # 경산시 내부 번호
        route.cell(row, 2, int(info[3].contents[0])) # 확진 번호
        route.cell(row, 3, info[5].contents[0]) # 인적사항 : 성별(거주하는 동, 나이)
        day = info[9].contents[0].split(".")
        route.cell(row, 4, day[0]+"월 "+day[1]+"일")
        route.cell(row, 5, info[11].contents[0])
        route.cell(row, 6, info[13].contents[0])
        where = I[3].contents
        extra = []
        for j in where[1::2]:
            if j.contents[0] != "\n":
                j = j.contents[0].contents[0]
            else:
                j = j.contents[1].contents[0]
            if j == "경로 확인중":
                route.cell(row, 9, "경로 확인중")
            elif j == "주요동선(직장,병의원,약국)없음":
                route.cell(row, 9, "주요동선(직장,병의원,약국)없음")
            else:
                jlist = [j]
                if "※" in j:
                    if "\n※" in j:
                        Route = j.split("\n※ ")
                    elif " ※" in j:
                        Route = j.split(" ※ ")
                    else:
                        Route = j.split("※ ")
                    extra.append(Route[1])
                    j = Route[0]
                if "\n" in j:
                    jlist = j.split("\n")
                for k in jlist:
                    row += 1
                    if k != jlist[0] and ":" not in k:
                        route.cell(row, 9, k)
                        continue
                    if "사망" in k:
                        if "-" in k:
                            Route = k.split(" - ")
                        elif "–" in k:
                            Route = k.split("–")
                        else:
                            Route = k.split(" ")
                    elif ")~" in k:
                        Route = k.split("~")
                    elif " ~ " in k:
                        Route = k.split(" ~ ")
                    elif "~" in k:
                        Route = k.split("~")
                    else:
                        Route = ["",j]
                    route.cell(row, 7, Route[0])
                    if "-" in Route[1]:
                        Route = Route[1].split(" - ")
                    elif "–" in Route[1]:
                        Route = Route[1].split("–")
                    elif " " in Route[1]:
                        Route = Route[1].split(" ")
                    else:
                        Route = ["", Route[0]]
                    route.cell(row, 8, Route[0])
                    route.cell(row, 9, Route[1])
                for k in extra:
                    row += 1
                    route.cell(row, 10, k)
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
            mask.cell(row, 3, [1].contents[0]) # 전화번호
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
            mask.cell(row, 2, I[1].contents[0]) # 판매처 이름
            mask.cell(row, 3, I[2].contents[0]) # 전화번호
            mask.cell(row, 4, "공적판매처") # 공사구분
            mask.cell(row, 5, "읍/면") # 동 읍/면 구분
        else:
            mask.cell(row, 2, I[0].contents[0]) # 판매처 이름
            mask.cell(row, 3, I[1].contents[0]) # 전화번호
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
    html = urlopen("http://www.gbgs.go.kr/design/health/COVID19/COVID19_05_05.html") # 경산시 판매 현황
    stock = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.div.table.tbody.contents
    for i in stock[1::2]:
        row = 1
        i = i.contents
        name = i[1].contents[0] # 판매처 이름
        while mask.cell(row, 2).value not in [None, name]: # 판매처 인덱스 찾기
            row += 1
        Stock = i[5].contents[0]
        if "(" in Stock:
            Stock = Stock.split(" (")
            Stock[1] = "(" + Stock[1]
        else:
            Stock = [Stock, ""]
        mask.cell(row, 6, Stock[0]) # 재고 수준
        mask.cell(row, 7, Stock[1]) # 재고량
        mask.cell(row, 8, i[9].contents[0]) # 재고량 갱신 시간
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
        row += 1
    wb.save(path+"/코로나데이터.xlsx")

occurrence()
clinicinfo()
maskinfo()
