import ExcelJS from 'exceljs';
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile("/Users/shubham/Downloads/Music_player/borehole/coveredBoreholes.xlsx").then(()=>{
    const worksheet=workbook.getWorksheet(1);
    worksheet.eachRow((row,rowNumber)=>{
        // console.log(row.getCell(6).value)
        if(row.getCell(6).value==="SEAM-XI"){
            console.log(row.values)
        }
    })
})
.catch((err)=>{
    console.log(err)
})