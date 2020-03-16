from urllib.request import urlopen
from bs4 import BeautifulSoup
import openpyxl as xl
import os

path = os.path.dirname(__file__)
wb = xl.load_workbook(path+"/코로나데이터.xlsx", data_only=True)
status = wb.create_sheet("발생 현황")
route = wb.create_sheet("이동 경로")
clinic = wb.create_sheet("선별진료소")
mask = wb.create_sheet("마스크 판매 약국")

html = urlopen("http://www.gbgs.go.kr/programs/coronaMove/coronaMove.do")
bsObject = BeautifulSoup(html, "html.parser").body.contents[5].div.div.div.contents
Patients = {}


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

print(bsObject)