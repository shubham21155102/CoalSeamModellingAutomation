# CoalSeamModellingAutomation

### Note: Due to restrictions I can not compromise with data used

# main.py

The script provided reads geological data from an Excel file and generates a DXF file visualizing lithological columns.

### 1. Loading the Excel Workbook

The script begins by loading an Excel workbook using the `openpyxl` library:

```python
from openpyxl import load_workbook
file_path = '/Users/shubham/Downloads/Music_player/pyautocad/geological_log.xlsx'
workbook = load_workbook(filename=file_path, data_only=True)
print("Sheet names:", workbook.sheetnames)
sheets = workbook.sheetnames
```

### 2. Extracting Collar Data

The first sheet contains collar data (borehole identification and coordinates):

```python
collar_data = []
sheet = workbook[sheets[0]]
for row in sheet.iter_rows(values_only=True):
    BHID = row[1]
    XCOLLAR = row[4]
    YCOLLAR = row[5]
    ZCOLLAR = row[6]
    DEPTH = row[7]
    collar_row = (BHID, XCOLLAR, YCOLLAR, ZCOLLAR, DEPTH)
    collar_data.append(collar_row)
```

### 3. Extracting Lithology Data

The second sheet contains lithological data:

```python
litho_data = []
sheet = workbook[sheets[1]]
temp=""
temp_arr=[]
for row in sheet.iter_rows(values_only=True):
    BHID = row[0]
    if BHID!=temp:
        temp=BHID
        litho_data.append(temp_arr)
        temp_arr=[]
    FROM = row[1]
    TO = row[2]
    DEPTH = row[3]
    DESC = row[4]
    LITHOLOGY = row[5]
    LITHOID = row[6]
    litho_row = (BHID, FROM, TO, DEPTH, DESC, LITHOLOGY, LITHOID)
    temp_arr.append(litho_row)
```

### 4. Initializing DXF File Creation

The script initializes the creation of a DXF file using the `ezdxf` library:

```python
import ezdxf
from ezdxf.enums import TextEntityAlignment

doc = ezdxf.new(dxfversion='R2010')
msp = doc.modelspace()
x_start = 0
y_start = 0
width = 25
```

### 5. Drawing Lithological Columns

For each borehole in the lithology data, it creates a corresponding lithological column in the DXF file. Specific colors are assigned based on lithology:

```python
lithology_colors = {
    'TS': 4,            # Cyan
    'OB': 2,            # Yellow
    'WM': 3,            # Green
    'SEAM-IX': 1,       # Red
    'SEAM-XI': 1,       # Red
    'SEAM-X': 1,        # Red
    'SEAM-VIII': 1,     # Red
    'IB': 5,            # Magenta
    'SST': 6,           # Cyan
}
spacing=0
for drawing in drawings:
    text_position = (10+spacing,10)
    text = msp.add_text(drawing[0][0], dxfattribs={'height': 1.5, 'layer': drawing[0][0], 'color': 0})
    text.set_placement(text_position, align=TextEntityAlignment.MIDDLE_CENTER)
    for x in drawing:
        start_depth = float(x[1])
        end_depth = float(x[2])
        thickness = float(x[3])
        lithology = x[5]
        layer=x[6]
        color = lithology_colors.get(lithology, 1)
        p1=(x_start+spacing, y_start - start_depth)
        p2=(x_start+spacing + width, y_start - start_depth)
        p3=(x_start+spacing + width, y_start - start_depth - thickness)
        p4=(x_start+spacing, y_start - start_depth - thickness)
        try:
            msp.add_lwpolyline([p1, p2, p3, p4, p1], dxfattribs={'layer': 'Layer1', 'color': color})
            hatch = msp.add_hatch(color=color)
            hatch.paths.add_polyline_path([p1, p2, p3, p4, p1])
            text_position = (x_start +spacing+ width + 2, y_start - (start_depth + thickness / 2))
            text = msp.add_text(lithology, dxfattribs={'height': 1.5, 'layer': layer, 'color': 0})
            text.set_placement(text_position, align=TextEntityAlignment.LEFT)
        except:
            print("Error")
    spacing+=100
```

### 6. Saving the DXF File

Finally, the DXF file is saved:

```python
doc.saveas('geological_log.dxf')
print("DXF file 'geological_log.dxf' created successfully.")
```

This script reads borehole data from an Excel file and generates a visual representation of lithological columns in a DXF file, with each lithology type colored differently for easy identification.

# borehole/index.js


The provided script is designed to read data from an Excel file using the `exceljs` library, filter specific borehole data, and create a new Excel file containing filtered information. Here’s a detailed explanation of the script:

### 1. Importing the Required Libraries

The script starts by importing the `exceljs` library:

```javascript
import ExcelJS from 'exceljs';
```

### 2. Defining the Covered Boreholes

A list of borehole IDs (`coveredBoreHoles`) is defined, which will be used to filter the data:

```javascript
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
```

### 3. Reading the Excel File and Filtering Collar Data

The script initializes sets to store unique borehole data and reads the first sheet of the Excel file to extract collar data:

```javascript
const set = new Set();
const workbook = new ExcelJS.Workbook();
const newWorkbook = new ExcelJS.Workbook();

workbook.xlsx.readFile('/Users/shubham/Downloads/Music_player/borehole/borehole.xlsx').then(() => {
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow((row, rowNumber) => {
        if (coveredBoreHoles.includes(row.getCell(2).value)) {
            const obj = {
                boreHoleNumber: row.getCell(2).value,
                Northing_LC: row.getCell(3).value,
                Easting_LC: row.getCell(4).value,
                Easting_UC: row.getCell(5).value,
                Northing_UC: row.getCell(6).value,
                RL: row.getCell(7).value,
                depth: row.getCell(8).value,
            }
            set.add(obj);
        }
    });
}).then(() => {
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
```

### 4. Reading the Excel File and Filtering Lithology Data

The script initializes more sets and a map to handle the lithology data from the second sheet of the Excel file:

```javascript
const set2 = new Set();
const set3 = new Set();
const map = new Map();

workbook.xlsx.readFile('/Users/shubham/Downloads/Music_player/borehole/borehole.xlsx').then(() => {
    const worksheet = workbook.getWorksheet(2);
    worksheet.eachRow((row, rowNumber) => {
        if (coveredBoreHoles.includes(row.getCell(1).value)) {
            const obj = {
                boreHoleNumber: row.getCell(1).value,
                From: row.getCell(2).value.result,
                To: row.getCell(3).value,
                Thickness: row.getCell(4).value.result,
                LithologyDescription: row.getCell(5).value,
                SeamName: row.getCell(6).value,
                SeamID: row.getCell(7).value,
            }

            if (obj.From === undefined) {
                obj.From = 0;
            }

            set2.add(obj);
            if (obj.SeamName === 'SEAM-VIII') {
                map.set(obj.boreHoleNumber, true);
            } else {
                if (map.get(obj.boreHoleNumber) === undefined) {
                    map.set(obj.boreHoleNumber, false);
                }
            }
        }
    });
}).then(() => {
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
                if (map.get(value2.boreHoleNumber) === true) {
                    set3.add(value2);
                }

                if (value2.SeamName === 'SEAM-VIII') {
                    break;
                }
            }
        }
    });

    set3.forEach((value) => {
        newWorksheet.addRow(value);
    });

    return newWorkbook.xlsx.writeFile('/Users/shubham/Downloads/Music_player/borehole/coveredBoreholes.xlsx');
}).catch(err => {
    console.error('Error:', err);
});
```

### Explanation of Key Operations

1. **Reading the Excel file**:
   The script reads the Excel file located at `/Users/shubham/Downloads/Music_player/borehole/borehole.xlsx`.
2. **Filtering Collar Data**:
   It reads the first sheet of the Excel file and filters rows based on the `coveredBoreHoles` array. The filtered data is stored in a set and later written to a new Excel file.
3. **Filtering Lithology Data**:
   It reads the second sheet of the Excel file, filters rows, and checks for the presence of specific seam names (e.g., 'SEAM-VIII'). The filtered data is processed and written to the new Excel file.
4. **Writing to a New Excel File**:
   The filtered collar and lithology data are written to new worksheets ('Collar File' and 'Lithology File') in a new Excel file located at `/Users/shubham/Downloads/Music_player/borehole/coveredBoreholes.xlsx`.

### Error Handling

The script includes error handling to catch and log any errors that occur during the file operations:

```javascript
}).catch(err => {
    console.error('Error:', err);
});
```

This ensures that any issues encountered during the execution are reported, making it easier to debug and fix potential problems.
