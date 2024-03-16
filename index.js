const  XLSX = require('xlsx') 

const normalWorkbook = XLSX.readFile('./files/常规迭代.xlsx') 
const sheetName = normalWorkbook.SheetNames[0]
const worksheet = normalWorkbook.Sheets[sheetName]
const data = XLSX.utils.sheet_to_json(worksheet)
const peopleMap = {} 

for (let i = 0; i < data.length; i++) {
    const item = data[i]
    const curPeople = item.__EMPTY_5
    if (peopleMap[curPeople])  {
        peopleMap[curPeople].score = peopleMap[curPeople].score + item.__EMPTY_4
    } else {
        peopleMap[curPeople] = {
            score:  item.__EMPTY_4
        }
    }
}
delete peopleMap['开发人']

console.log(peopleMap)



