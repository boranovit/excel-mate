const { Workbook } = require("excel4node")
const Column = require("./column")

class Sheet {
    constructor(wb, { name, registers, columns, styles, wsCount }) {
        this.name = this.formatWs(name, wsCount) || ""
        this.registers = registers || []
        this.columns = columns || []
        this.wb = wb || new Workbook()
        this.ws = this.wb.addWorksheet(this.name);
        this.styles = styles || this.wb.createStyle({
            alignment: {
                horizontal: "center",
                wrapText: false,
                vertical: 'center'
            },
        }),
        this.freezedHeaders = true
        this.registersRow = 3
        this.colNum = 0
    }
    
    init(){
        return
    }

    write(){
        this.init();
        this.registersRow = countHeadersRows(this.columns) + 1;
        this.setWidth(this.columns);
        this.writeHeaders(this.columns, 1);
        if (this.freezedHeaders) this.freezeHeaders();
        this.writeRegisters();
    }
    
    addColumns(columns){
        columns.forEach(c=>{
            this.columns.push(new Column({
                label: c.label,
                name: c.name,
                num: ++this.colNum,
                style: c.style || this.styles,
                styleHeader: c.styleHeader || this.styles,
                width: c.width || 20,
                children: c.children && c.children.map((child, i)=> new Column({
                    label: child.label,
                    name: child.name,
                    num: i == 0 ? this.colNum : ++this.colNum,
                    style: child.style || this.styles,
                    styleHeader: child.styleHeader || this.styles,
                    width: child.width
                })),
            }))
        })
    }

    setWidth(columns){
        let childs = []
        if(columns && columns.length){
            columns.forEach((col) => {
                if (col.children && col.children.length) childs = [...childs, ...col.children]
                this.ws.column(col.num).setWidth(col.width); 
            })
        } 
        if(childs.length) this.setWidth(childs)
    }

    writeHeaders(columns, row) {
        let childs = []
        if(columns && columns.length){
            columns.forEach((col) => {
                if (col.children && col.children.length) {
                    this.ws.cell(row, col.num, row, (col.num + countColumnsParent([col], 0)), true)
                            .string(col.label || col.name).style(col.styleHeader || this.styles);
                    childs = [...childs, ...col.children]
                } else {
                    this.ws.cell(row, col.num, this.registersRow - 1, col.num, true)
                            .string(col.label || col.name).style(col.styleHeader || this.styles);
                }
            });
        } 
        if(childs.length) this.writeHeaders(childs, ++row);
      }
    
    freezeHeaders(){
        for (let i = 1 ; i < this.registersRow; i++){
            this.ws.row(i).freeze();
        }
    }
    
    writeRegister(register, columns, styleRegister, row){
        let childs = [];
        if(columns && columns.length){
            columns.forEach((col) => {
                if (col.children && col.children.length) childs = [...childs, ...col.children];
                else this.ws.cell(row, col.num)
                            .string(this.formatText(register[col.name] || "")).style(styleRegister[col.name] ||col.style);
            })
        }
        if(childs.length) this.writeRegister(register, childs, styleRegister, row);
    }
    
    writeRegisters() {
        let row = this.registersRow;
        this.registers.forEach((register) => {
            let styleRegister = {};
            if(this.answerStyle) styleRegister = this.setAnswerStyle(register);
            this.writeRegister(register, this.columns, styleRegister, row);
            row++;
        });
    }

    formatWs(str, i) {
        return str
    }
    formatText(str) {
        return str.toString()
    }
}



function countColumnsParent(columns, total){
    let childs = [];
    columns.forEach(col=>{
        if(col.children && col.children.length){
            total += col.children.length - 1;
            childs = [...childs, ...col.children]
        }
    })
    if(childs.length) return countColumnsParent(childs, total)
    return total;
}

function countNestedChildren(column) {
    if (!column || !column.children) return 1;
    const childCounts = column.children.map(child=> child && countNestedChildren(child));
    return 1 + Math.max(...childCounts);
}
function countHeadersRows(columns){
    let maxRowCount = 0;
    columns.forEach((column) => {
        const rowCount = countNestedChildren(column);
        if (rowCount > maxRowCount) maxRowCount = rowCount;
    });
    return maxRowCount;
}

module.exports = Sheet