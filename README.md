# Excel Builder

The `excel_builder.js` file is responsible for generating the Excel file.

It contains a function named `main`, which accepts a JSON object as input.

### JSON Requirements

The JSON data **must include the full path (location) of each image**, not just the image file name.

**Example:**

```json
[
  {
    "Beam width": 200,
    "Beam Depth": 500,
    "Beam Length": 4980,
    "Structural Member name": "OUTER RING STIRRUPS-TRIE",
    "Spacing": "",
    "Dia (T) (mm)": 8,
    "Shape": "images/shape_1",
    "L1 (m)": 0.065,
    "L2 (m)": 0.14,
    "L3 (m)": 0.44,
    "L4 (m)": 0.14,
    "L5 (m)": 0.44,
    "L6 (m)": 0.065,
    "Lap (m)": 0.328,
    "No. Of Bentup": "",
    "Add bentup length": "",
    "No of Bend": 5,
    "Bend deduction (m)": 0.096,
    "Cutting length (m)": 1.194,
    "No. of member (nos.)": 1,
    "No. of bars in each (nos.)": 25,
    "Total nos. (nos.)": 25,
    "Total Length (m)": 29.85,
    "8 mm": 9.875,
    "10 mm": "-",
    "12 mm": "-",
    "16 mm": "-",
    "20 mm": "-",
    "25 mm": "-",
    "32 mm": "-"
  }
]
```

> **Note:** Supplying only the image file name (for example, `product1.jpg`) is not sufficient. The `image` field must contain the complete file path so that `createExcelWithImages` can locate and embed the image into the generated Excel file.
