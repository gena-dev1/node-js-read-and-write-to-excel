const reader = require('xlsx')
const xl = require('excel4node')
// Read from file
const file = reader.readFile('./all.ods', { sheetStubs: true })

// words to filter
const nottt = ["bot", "yandex.net", "google"]

// name of columns
const headingColumnNames = [
    "country",
    "ip",
    "day",
    "?",
    "domen",
    "link",
]//Write Column Title in Excel file


const sheets = file.SheetNames
const wb = new xl.Workbook();
let data = []
for(let i = 0; i < sheets.length; i++)
{
    const temp = reader.utils.sheet_to_json( file.Sheets[file.SheetNames[0]], { header: 0, defval: "" })
    temp.forEach((res) => {
        res?.domen === undefined || res?.domen === null || res?.domen == ""
        ?
        data.push(res)
        :
        nottt.every( (nott) => {
            if (res.domen.toLowerCase().includes(nott.toLowerCase())) {
                return false
            } else {
                if(nott === nottt[nottt.length-1]){
                    data.push(res)
                }
            }
            return true
        })
        
    })

    const ws = wb.addWorksheet(file.SheetNames[i])
    let headingColumnIndex = 1;
    headingColumnNames.forEach(heading => {
        ws.cell(1, headingColumnIndex++).string(heading)
    });//Write Data in Excel file
    let rowIndex = 2;
    data.forEach( record => {
        let columnIndex = 1;
        Object.keys(record ).forEach(columnName =>{
            ws.cell(rowIndex,columnIndex++)
                .string(record [columnName])
        });
        rowIndex++;
    })

    // new excel file to write
    wb.write('data.xlsx');    
    data = []
}
