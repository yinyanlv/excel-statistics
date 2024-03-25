const  XLSX = require('xlsx') 

function readExcel(options) {
    const workbook = XLSX.readFile(options.path) 
    const peopleMap = {} 

    for (let i = 0; i < workbook.SheetNames.length; i++) {
        const sheetName = workbook.SheetNames[i]
        const worksheet = workbook.Sheets[sheetName]
        const data = XLSX.utils.sheet_to_json(worksheet)
    
        // 冒烟失败、 线上故障
        if (options.isCountMode) {
            for (let j = 0; j < data.length; j++) {
                const item = data[j]
                const curPeople = item[options.peopleKey]
                if (peopleMap[curPeople])  {
                    peopleMap[curPeople][options.fieldName] = peopleMap[curPeople][options.fieldName] + 1 
                } else {
                    peopleMap[curPeople] = {
                        [options.fieldName]: 1
                    }
                }
            }       
        } else {
            for (let j = 0; j < data.length; j++) {
                const item = data[j]
                const curPeople = item[options.peopleKey]
                if (peopleMap[curPeople])  {
                    peopleMap[curPeople][options.fieldName] = peopleMap[curPeople][options.fieldName] + item[options.numberKey]
                } else {
                    peopleMap[curPeople] = {
                        [options.fieldName]:  item[options.numberKey]
                    }
                }
            }
        }
    }
    if (options.deleteList) {
        options.deleteList.forEach((key) => {
            delete peopleMap[key]
        })
    }

    return peopleMap 
}

function getStatisticsData() {
    const categoryMap = {
        '常规迭代': getNormalWorkData(),
        '紧急需求': getEmergentWorkData(),
        '技术调研': getSurveyData(),
        '技术分享': getShareData(),
        '线上问题解决记录': getSolveProblemData(),
        '冒烟失败': getTestData(),
        '线上故障': getFaultData()
    }

    return categoryMap
}

function getNormalWorkData() {
    return readExcel({
        path: './files/常规迭代.xlsx',
        fieldName: 'score',
        peopleKey: '__EMPTY_5',
        numberKey: '__EMPTY_4',
        deleteList: ['开发人'] 
    })
}

function getEmergentWorkData() {
    return readExcel({
        path: './files/紧急需求.xlsx',
        fieldName: 'score',
        peopleKey: '__EMPTY_5',
        numberKey: '__EMPTY_4',
        deleteList: ['开发人', undefined] 
    })
}

function getSurveyData() {
    return readExcel({
        path: './files/技术调研.xlsx',
        fieldName: 'score',
        peopleKey: '__EMPTY_2',
        numberKey: '__EMPTY_1',
        deleteList: ['调研人'] 
    })
}

function getShareData() {
    return readExcel({
        path: './files/技术分享.xlsx',
        fieldName: 'score',
        peopleKey: '__EMPTY_1',
        numberKey: '__EMPTY_3',
        deleteList: ['分享人'] 
    })
}

function getSolveProblemData() {
    return readExcel({
        path: './files/线上问题解决记录.xlsx',
        fieldName: 'score',
        peopleKey: '__EMPTY_4',
        numberKey: '__EMPTY_3',
        deleteList: ['解决人'] 
    })
}

function getTestData() {
    return readExcel({
        isCountMode: true,
        path: './files/冒烟失败.xlsx',
        fieldName: 'count',
        peopleKey: '__EMPTY_3',
        deleteList: ['开发人'] 
    })
}

function getFaultData() {
    return readExcel({
        isCountMode: true,
        path: './files/线上故障.xlsx',
        fieldName: 'count',
        peopleKey: '__EMPTY_4',
        deleteList: ['责任人'] 
    })
}

const data = getStatisticsData()

console.table(data)
