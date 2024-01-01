# excel-mate
#### Node.js library designed to simplify the process of creating Excel files on top of the powerful "excel4node" library.

### Usage
```javascript

const { Workbook, Sheet } = require("excel-mate")

const wb = new Workbook();              // excel4node Workbook

const sheet1 = new Sheet(wb, {
    name: "sheet1",
})

sheet1.addColumns([
    {
        label: "Names",
        name: "names",      
        style: null,                    // excel4node styles for registers cells
        styleHeader: wb.createStyle({   // excel4node styles for headers cells
            alignment: {
                horizontal: "center",
            }
        }),
        children: [
            {
                label: "First Name",
                name: "first_name",
                width: 40
            },
            {
                label: "Last Name",
                name: "last_name",
                width: 30,
                style: wb.createStyle({
                    font: {
                        bold: true,
                    },
                }),       
            }
        ]
    },
    {
        label: "Score",
        name: "score",      
        style: wb.createStyle({
            alignment: {
                horizontal: "center",
            },
            fill: {
                type: 'pattern',
                patternType: 'solid',
                bgColor: '#00ff00',
                fgColor: '#00ff00',
            },
            font: {
                bold: true,
                color: '#ffffff',
            },
        }),
        styleHeader: wb.createStyle({
            alignment: {
                horizontal: "center",
                vertical: "center"
            },
        }),
        width: 20,
    },
])

sheet1.registers = [
    { first_name: "Lionel", last_name: "Messi" , score: 22},
    { first_name: "Maria", last_name: "Smith" , score: 31},
    { first_name: "Alex", last_name: "Hansen" , score: 14},
    { first_name: "Anna", last_name: "Garcia" , score: 23}
]

sheet1.write()

wb.write('filename.xlsx', async function (err) {});

```