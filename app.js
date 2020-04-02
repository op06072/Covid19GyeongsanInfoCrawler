//import db from './firebaseInit'

const firebase = require('firebase');
const firestore = require('firestore');
require('firebase/firestore');
//export default firebaseApp.firestore()

const firebaseConfig = {
    apiKey: "AIzaSyAkGx878CnZNySeYntTL0Ly1Y6rc64aOaA",
    authDomain: "covid19gyeongsan.firebaseapp.com",
    databaseURL: "https://covid19gyeongsan.firebaseio.com",
    projectId: "covid19gyeongsan",
    storageBucket: "covid19gyeongsan.appspot.com",
    messagingSenderId: "1042161209413",
    appId: "1:1042161209413:web:5d24f2d8a00a6ec610132c",
    measurementId: "G-0HW8JR3B7C"
  };
var firebaseApp = firebase.initializeApp(firebaseConfig);
var db = firebaseApp.firestore();
const fs = require('fs');
const dataBuffer = fs.readFileSync('./corona.json');
const datajson = dataBuffer.toString();
const data = JSON.parse(datajson);

for (key in data.선별진료소) {
    var clinic = data.선별진료소[key];
    var locate = new firebase.firestore.GeoPoint(clinic.좌표.y, clinic.좌표.x)
    
    db.collection('선별진료소').doc(key).set({
        '주소' : clinic.주소,
        '전화번호' : clinic.전화번호,
        '좌표' : locate,
        '진료시간' : clinic.진료시간,
        '비고' : clinic.비고
    });
}

db.collection('발생동향').doc('기본정보').set(data.발생동향.기본정보)
db.collection('발생동향').doc('확진자수').set(data.발생동향.확진자수)
for (key in data.확진자동선) {
    var route = data.확진자동선[key]
    var move = {}

    if ('이동경로' in route) {
        for (i in route.이동경로) {
            move[i] = { '경로' : route.이동경로[i].경로}
            if ('좌표' in route.이동경로[i]) {
                move[i]['좌표'] = new firebase.firestore.GeoPoint(route.이동경로[i].좌표.y, route.이동경로[i].좌표.x)
            }
        }
    }

    var routedata = {
        '거주지' : route.거주지,
        '확진번호' : route.확진번호,
        '인적사항' : route.인적사항,
        '접촉자수(격리조치중)' : route['접촉자수(격리조치중)'],
        '입원기관' : route.입원기관,
        '확진일자' : route.확진일자,
        '이동경로' : move
    }

    if ('비고' in route) routedata.비고 = route.비고

    db.collection('확진자동선').doc(key).set(routedata)

}

let localmask = {}
for (i in data.공적마스크) {
    var mask = data.공적마스크[i]
    var locallist = []
    if (i != "축협") {
        localmask[i] = []   
    }
    if ('공적판매처' in mask) {
        for (j in mask.공적판매처) {
            var sell = mask.공적판매처[j]
            var maskInfo = {
                '갱신일시' : sell.갱신일시,
                '재고수준' : sell.재고수준,
                '전화번호' : sell.전화번호,
                '좌표' : new firebase.firestore.GeoPoint(sell.좌표.y, sell.좌표.x)
            }
            if ('재고량' in sell) maskInfo['재고량'] = sell.재고량
            if (i != "축협") localmask[i].push(j)
            db.collection('공적마스크').doc(i).collection('공적판매처').doc(j).set(maskInfo)
        }
    }
    if ('약국' in mask) {
        for (j in mask.약국) {
            var sell = mask.약국[j]
            var maskInfo = {
                '갱신일시' : sell.갱신일시,
                '재고수준' : sell.재고수준,
                '전화번호' : sell.전화번호,
                '주소' : sell.주소,
                '좌표' : new firebase.firestore.GeoPoint(sell.좌표.y, sell.좌표.x)
            }
            if ('재고량' in sell) maskInfo['재고량'] = sell.재고량
            localmask[i].push(j)
            db.collection('공적마스크').doc(i).collection('약국').doc(j).set(maskInfo)
        }
    }
}
db.collection('공적마스크').doc('판매지역').set(localmask)

for (key in data.발생동향.확진자수) {
    var patientnum = {}
    if (key == "전체") {
        for (i in data.확진자동선) {
            patientnum[i] = true
        }
    }
    else {
        for (i in data.확진자동선) {
            if (key == data.확진자동선[i].거주지) {
                patientnum[i] = true
            }
        }
    }
    db.collection('지역별확진자번호').doc(key).set(patientnum)
}

console.log('upload finish!')