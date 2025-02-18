# xlWriter
simple python code to write xlsx files without pandas or openpyxl
just in case you can't install pandas or openpyxl wherever you need to write to excel workbook, you can use this.

```
SAMPLE_DATA = {
    'sales':{
        'table': [
            ["Product", "Units Sold", "Revenue ($)"],
            ["Laptop", 120, 96000],
            ["Smartphone", 250, 125000],
            ["Tablet", 180, 54000],
            ["Headphones", 300, 45000]
        ],
        'header': True,
        'columnWidths': [20, 10, 10]
    },
    'plants': {
        'table': [
            ["Plant Name", "Type", "Water Needs", "Sunlight"],
            ["Rose", "Flower", "Medium", "Full Sun"],
            ["Cactus", "Succulent", "Low", "Full Sun"],
            ["Fern", "Foliage", "High", "Partial Shade"],
            ["Aloe Vera", "Succulent", "Low", "Bright Indirect"],
            ["Tulip", "Flower", "Medium", "Full Sun"]
        ]
    }
}

from xlWriter import createWorkBook
createWorkBook(SAMPLE_DATA, 'TEST.xlsx')
```
