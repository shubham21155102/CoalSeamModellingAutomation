# MineGeologyAutomationOnAutoCad
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
