from urllib.request import urlopen
from bs4 import BeautifulSoup

html = urlopen("http://www.gbgs.go.kr/programs/coronaMove/coronaMove.do")
bsObject = BeautifulSoup(html, "html.parser").body.contents[5].contents[1].contents[1].contents[1].contents
Patients = {}

for i in bsObject[3::2]:
    Where = []
    When = {}
    I=i.tbody.contents
    info = I[1].contents
    Info = {"인적사항" : info[5].contents[0]} #인적사항 : 성별(거주하는 동, 나이)
    day = info[9].contents[0].split(".")
    Info["확진일자"] = day[0]+"월 "+day[1]+"일"
    Info["입원기관"] = info[11].contents[0]
    Info["접촉자 수(격리조치 중 접촉자 수)"] = info[13].contents[0]
    Patients[int(info[1].contents[0])] = Info
    where = I[3].contents[1].contents
    for j in where[1::2]:
        j = j.contents[1].contents[0]
        if j == "경로 확인중":
            When["None"] = ["경로 확인 중"]
            Patients["이동 경로"] = When
            break

        if j[0] is "※":
            when = j.split("/")
            month = int(when[0])
            when = when[1].split(" ")
            date = int(when[0])
            do = when[1]


print(bsObject)