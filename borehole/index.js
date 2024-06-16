import ExcelJS from 'exceljs';
const coveredBoreHoles = [
    'CMTU-001', 'CMTU-014', 'CMTU-015', 'CMTU-017',
    'CMTU-019', 'CMTU-061', 'CMTU-067', 'CMTU-081', 'CMTU-084',
    'CMTU-088', 'CMTU-126', 'CMTU-130', 'CMTU-132', 'CMTU-133',
    'CMTU-135', 'CMTU-136', 'CMTU-150', 'CMTU-158', 'CMTU-161',
    'CMTU-164', 'CMTU-165', 'CMTU-229', 'CMTU-232', 'CMTU-233',
    'CMTU-234', 'CMTU-236', 'CMTU-237', 'CMTU-238', 'CMTU-244',
    'CMTU-245', 'CMTU-250', 'CMTU-252', 'CMTU-254', 'CMTU-257',
    'CMTU-258', 'CMTU-259', 'CMTU-260', 'CMTU-261', 'CMTU-262',
    'CMTU-265', 'CMTU-266', 'UT-010'
];
const set=new Set();
const workbook = new ExcelJS.Workbook();
const newWorkbook = new ExcelJS.Workbook();
workbook.xlsx.readFile('/Users/shubham/Downloads/Music_player/borehole/borehole.xlsx').then(() => {
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow((row, rowNumber) => {
        if (coveredBoreHoles.includes(row.getCell(2).value)) {
            const obj={
                boreHoleNumber:row.getCell(2).value,
                Northing_LC:row.getCell(3).value,
                Easting_LC:row.getCell(4).value,
                Easting_UC:row.getCell(5).value,
                Northing_UC:row.getCell(6).value,
                RL:row.getCell(7).value,
                depth:row.getCell(8).value,
            }
            set.add(obj);
        }
    });
}).then(() => {
    // console.log(set);
    const newWorksheet = newWorkbook.addWorksheet('Collar File');
    newWorksheet.columns = [
        { header: 'Borehole Number', key: 'boreHoleNumber', width: 20 },
        
        { header: 'Northing_LC', key: 'Northing_LC', width: 20 },
        { header: 'Easting_LC', key: 'Easting_LC', width: 20 },
        { header: 'Easting_UC', key: 'Easting_UC', width: 20 },
        { header: 'Northing_UC', key: 'Northing_UC', width: 20 },
        { header: 'RL', key: 'RL', width: 20 },
        { header: 'Depth', key: 'depth', width: 20 }
    ];
    set.forEach((value) => {
        newWorksheet.addRow(value);
    });
    return newWorkbook.xlsx.writeFile('/Users/shubham/Downloads/Music_player/borehole/coveredBoreholes.xlsx');

}).catch(err => {
    console.error('Error:', err);
});
const set2=new Set();
const set3=new Set();
const map=new Map();
workbook.xlsx.readFile('/Users/shubham/Downloads/Music_player/borehole/borehole.xlsx').then(() => {
    const worksheet = workbook.getWorksheet(2);
    worksheet.eachRow((row, rowNumber) => {
        if (coveredBoreHoles.includes(row.getCell(1).value)) {
            const obj={
                boreHoleNumber:row.getCell(1).value,
                From:row.getCell(2).value.result,
                To:row.getCell(3).value,
                Thickness:row.getCell(4).value.result,
                LithologyDescription:row.getCell(5).value,
                SeamName:row.getCell(6).value,
                SeamID:row.getCell(7).value,
                
            }
            
            if(obj.From===undefined){
                obj.From=0;
            }
           
            set2.add(obj);
            if(obj.SeamName==='SEAM-VIII'){
                map.set(obj.boreHoleNumber,true);
            }
            else{
                if(map.get(obj.boreHoleNumber)===undefined){
                    map.set(obj.boreHoleNumber,false);
                }
            }
        }
    });
}
).then(() => {
    // console.log(set2);
    const newWorksheet = newWorkbook.addWorksheet('Lithology File');
    newWorksheet.columns = [
        { header: 'Borehole Number', key: 'boreHoleNumber', width: 20 },
        { header: 'From', key: 'From', width: 20 },
        { header: 'To', key: 'To', width: 20 },
        { header: 'Thickness', key: 'Thickness', width: 20 },
        { header: 'Lithology Description', key: 'LithologyDescription', width: 20 },
        { header: 'Seam Name', key: 'SeamName', width: 20 },
        { header: 'Seam ID', key: 'SeamID', width: 20 }
    ];
   
    coveredBoreHoles.forEach((value) => {
        for (const value2 of set2) {
            if (value === value2.boreHoleNumber) {
                // console.log(value2)
                // newWorksheet.addRow(value);
               
                if(map.get(value2.boreHoleNumber)===true){
                    // console.log(value2.boreHoleNumber)
                    set3.add(value2);
                }
                
                if (value2.SeamName === 'SEAM-VIII') {
                    break;
                }
            }
        }
    });
    console.log(map)
    // console.log(set3)
    set3.forEach((value) => {
        // console.log(value)
        newWorksheet.addRow(value);
    });
    return newWorkbook.xlsx.writeFile('/Users/shubham/Downloads/Music_player/borehole/coveredBoreholes.xlsx');
    
}).catch(err => {
    console.error('Error:', err);
});



