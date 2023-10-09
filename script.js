let table = document.getElementsByClassName("sheet-body")[0],
    rows = document.getElementsByClassName("rows")[0],
    columns = document.getElementsByClassName("columns")[0]
tableExists = false

const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    table.innerHTML = ""
    for (let i = 0; i < rowsNumber; i++) {
        var tableRow = ""
        for (let j = 0; j < columnsNumber; j++) {
            tableRow += `<td contenteditable></td>`
        }
        table.innerHTML += tableRow
    }
    if (rowsNumber > 0 && columnsNumber > 0) {
        tableExists = true;

    }
    //This condition to check if fields of rows and columns have correct value or not 
    else if (isNaN(rowsNumber) || isNaN(columnsNumber)) {
        return (
            alert("____!!!You should fill fields first____")
        );
    }
    else if(rowsNumber <0 ||columnsNumber <0 ){
            alert("_____!!! You should enter  correct number of ROWS && COLUMNS _____")
    }
}

const ExportToExcel = (type, fn, dl) => {
    if (!tableExists) {
        // This alert to sure YOU generate your sheet or not
        return (
            alert("____!!!Can't do EXPORT without doing GENERATE ____")
        );
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}
