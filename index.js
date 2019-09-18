const cheerio = require('cheerio');
const superagent = require('superagent');
const data = require('./data');
const schoolLink = [];
const schoolList = [];
const express = require('express');
const app = express();
const XLSX = require('xlsx');

let sever = app.listen(3000, function(){});


app.get('/', function(req, res) {
    res.send(schoolList);
})


getSchool();

function promiseAgent (url) {
    return  new Promise(function(reslove, rejecet) {
         superagent.get(url).end((err,res)=>{
            if (err){
                rejecet('请求失败!');
                return;
            }
            if (res && res.status === 200) {
                let $ = cheerio.load(res.text);
                let school = {};
                let englishName = $('table tr td span#MainContentPlaceHolder_lblInfoEnglishName').text();
                    school.englishName = englishName;
        
                let chineseName = $('table tr td span#MainContentPlaceHolder_lblInfoChineseName').text();
                    school.chineseName = chineseName;
        
                let schoolNumber = $('table tr td span#MainContentPlaceHolder_lblInfoSchoolNumber').text();
                    school.schoolNumber = schoolNumber;
        
                let telephone = $('table tr td span#MainContentPlaceHolder_lblInfoTelephone').text();
                    school.telephone = telephone;
        
                let fax = $('table tr td span#MainContentPlaceHolder_lblInfoFax').text();
                    school.fax = fax;
        
                let gender = $('table tr td span#MainContentPlaceHolder_lblInfoGender').text();
                    school.studentGender = gender;
        
                let district = $('table tr td span#MainContentPlaceHolder_lblInfoDistrict').text();
                    school.district = district;
                    
                let level = $('table tr td span#MainContentPlaceHolder_lblInfoSchoolLevel').text();
                    school.schoolLevel = level;
        
                let type = $('table tr td span#MainContentPlaceHolder_lblInfoFinanceType').text();
                    school.schoolType = type;
        
                let website = $('table tr td span#MainContentPlaceHolder_lnkInfoSchoolWebsite1').text();
                    school.schoolWebsite = website;
        
                let address = $('table tr td span#MainContentPlaceHolder_lblInfoSchoolAddress').text();
                    school.schoolAddress = address;
                
                reslove(school);
            }
        })
    }) 
}
async function getSchoolDetail () {
    let $ = cheerio.load(data.data.text);

   await $('table.tbl-three-column tr td a').each((idx, ele) => {
        schoolLink.push($(ele).attr('href'));
    });


    for (let i=0; i<schoolLink.length; i++){
            let url = 'https://applications.edb.gov.hk/schoolsearch/' + schoolLink[i] ;
        
            await promiseAgent(url).then(function(school){
                    schoolList.push(school);
            }).catch(function(err){
                console.log(err);
            })
    
    }
}

async function exportExcel(schools) {
    const headers = ['englishName', 'chineseName', 'schoolNumber', 'telephone', 'fax', 'studentGender', 'district', 'schoolLevel', 'schoolType', 'schoolWebsite', 'schoolAddress'];
    var _headers = headers.map(
        (v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1})).reduce((prev, next)=> Object.assign({}, prev, {[next.position]: {v: next.v}}),{});
        
    var _data = schools.map((v, i) => headers.map((k,j) => Object.assign({}, {v: v[k], position: String.fromCharCode(65+j) + (i+2)}))).reduce((prev, next)=> prev.concat(next)).reduce((prev, next)=> Object.assign({}, prev, {[next.position]: {v: next.v}}), {});

    var output = Object.assign({}, _headers, _data);

    let outputs = Object.keys(output);
    let ref = outputs[0] + ':' + outputs[outputs.length - 1];
    let wb = {
        SheetNames: ['ShoolList'],
        Sheets: {
            'ShoolList': Object.assign({}, output, {'!ref': ref})
        }
    }

    XLSX.writeFile(wb, 'Schools.xlsx');
}
async function getSchool() {
    await getSchoolDetail();
    console.log('----------------------------')
    exportExcel(schoolList);
}

