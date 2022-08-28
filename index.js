const express = require("express");
const { google } = require("googleapis");
const bodyParser = require('body-parser')

const app = express();

//set app view engine
app.set('views', './views')
app.set("view engine", "ejs");
app.use(express.json());
app.use(bodyParser.urlencoded({
    extended: true
}));
app.use('/assets', express.static('assets'));

const port = process.env.PORT || 3000;
app.listen(port, ()=>{
    console.log(`server started on ${port}`)
});

const auth = new google.auth.GoogleAuth({
    keyFile: "keys.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
});

const NAME_INDEX = 1
const PHONE_INDEX = 2
app.get('/', (req, res) => {
    return res.render('index')
})

app.post("/", async (req, res) => {
    try {
        const { sourceSpreadsheetId, destinationSpreadsheetId } = req.body || {};
        if (!sourceSpreadsheetId || !destinationSpreadsheetId) {
            throw new Error('Заповніть усі поля форми')
        }

        const authClientObject = await auth.getClient();
        const googleSheetsInstance = await google.sheets({ version: "v4", auth: authClientObject })
        const sourceDataListsRequest = await getSourceData(googleSheetsInstance, sourceSpreadsheetId)
        if (!sourceDataListsRequest.length) {
            throw new Error('ID джерела порожній')
        }

        const sourceData = prepareSourceList(sourceDataListsRequest)

        const destinationData = await googleSheetsInstance.spreadsheets.values.get({
            auth,
            spreadsheetId: destinationSpreadsheetId,
            range: "A:H",
        })

        const existingData = getExistingDestinationData(destinationData, sourceData)
        await updateDestinationSource(googleSheetsInstance, destinationSpreadsheetId, existingData)

        return res.render('result', {
            isSuccess: true,
            message: 'Дані успішно опрацьовані'
        })
    } catch (e) {
        return res.render('result', {
            isSuccess: false,
            message: e.message
        })
    }
})

const prepareSourceList = (data = []) => {
    const result = []

    data.forEach(v => {
        if (v.data && v.data.values && v.data.values.length) {
            v.data.values.forEach(item => {
                result.push({
                    item: item,
                    name: new Set(String(item[NAME_INDEX]).trim().split(' ').map(v => String(v).toLowerCase())),
                    phone: item[PHONE_INDEX] ? preparePhone(String(item[PHONE_INDEX]).trim()) : null
                })
            })
        }
    })

    return result
}

const preparePhone = (phone) => {
    if (!phone) return null
    phone = String(phone).replaceAll(/\D/ig, '')

    let result = ''
    switch (phone.length) {
        case 13:
            result = phone
            break
        case 12:
            result = `+${phone}`
            break
        case 11:
            result = `+3${phone}`
            break
        case 10:
            result = `+38${phone}`
            break
        case 9:
            result = `+380${phone}`
            break
        default:
            result = phone
    }

    return result
}

const getSourceData = async (googleSheetsInstance, spreadsheetId) => {
    const sheets = await googleSheetsInstance.spreadsheets.get({
        auth,
        spreadsheetId,
    })
    console.log(sheets.data.sheets[0].sheetId);

    if (!sheets.data || !sheets.data.sheets || !sheets.data.sheets.length) {
        throw new Error('Джерело таблиці порожне')
    }

    return Promise.all(sheets.data.sheets.map(v => googleSheetsInstance.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: `${v.properties.title}!A:H`,
    })))
}

const getDestinationSheetId = async (googleSheetsInstance, spreadsheetId) => {
    const sheets = await googleSheetsInstance.spreadsheets.get({
        auth,
        spreadsheetId,
    })

    if (!sheets.data || !sheets.data.sheets || !sheets.data.sheets.length) {
        throw new Error('Форма таблиці порожня')
    }

    return sheets.data.sheets[0].properties.sheetId
}

const getExistingDestinationData = (destinationData, sourceData) => {
    const result = []
    if (destinationData.data && destinationData.data.values && destinationData.data.values.length) {
        destinationData.data.values.forEach((destinationItem, index) => {
            if (index < 1) return

            const user = destinationItem[NAME_INDEX] ? String(destinationItem[NAME_INDEX]).toLowerCase().trim().split(' ') : null
            const phone = destinationItem[PHONE_INDEX] ? preparePhone(String(destinationItem[PHONE_INDEX]).trim()) : null

            let isNameEqual = false
            let isPhoneEqual = false

            const isFindUser = sourceData.some(source => {
                let isFindByName = true
                if (user && user.length) {
                    user.forEach(name => {
                        if (!source.name.has(name)) {
                            isFindByName = false
                        }
                    })
                } else {
                    isFindByName = false
                }

                const isFindByPhone = source.phone === phone

                isNameEqual = isFindByName
                isPhoneEqual = isFindByPhone

                return isFindByName || isFindByPhone
            })

            if (isFindUser) {
                result.push({
                    index,
                    isPhoneEqual,
                    isNameEqual
                })
            }
        })
    }

    return result
}

const updateDestinationSource = async (googleSheetsInstance, spreadsheetId, data) => {
    const sheetId = await getDestinationSheetId(googleSheetsInstance, spreadsheetId)

    const requests = data.map(v => ({
        updateCells: {
            start: {
                sheetId: sheetId,
                rowIndex: v.index,
                columnIndex: v.isNameEqual ? NAME_INDEX : (v.isPhoneEqual ? PHONE_INDEX : 0),
            },
            fields: 'userEnteredFormat',
            rows: [
                {
                    values: [
                        {
                            userEnteredFormat: {
                                backgroundColor: {
                                    red: 1,
                                    green: 0,
                                    blue: 0,
                                    alpha: 0.5,
                                },
                            },
                        },
                    ],
                },
            ],
        },
    }))
    return googleSheetsInstance.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
            requests
        },
    })
}