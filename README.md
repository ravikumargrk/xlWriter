# xlWriter
simple python code to write xlsx files without pandas or openpyxl

just in case you can't install pandas or openpyxl wherever you need to write to excel workbook, you can use this.

```python
from xlWriter import createWorkBook, SAMPLE_DATA
createWorkBook(SAMPLE_DATA, 'TEST.xlsx')
```

Sample input: 
```python
# Print sample data to see how the input should be.
>>> import pprint
>>> pp = pprint.PrettyPrinter(indent=4)
>>> from xlWriter import SAMPLE_DATA
>>> pp.pprint(SAMPLE_DATA)
{  'sales': {     'columnWidths': [20, 10, 10],
                  'header': True,
                  'table': [  ['Product'   , 'Units Sold', 'Revenue ($)'],     
                              ['Laptop'    , 120         , 96000        ],
                              ['Smartphone', 250         , 125000       ],
                              ['Tablet'    , 180         , 54000        ],
                              ['Headphones', 300         , 45000        ]]}
   'plants': {   'table': [   ['Plant Name', 'Type'     , 'Water Needs', 'Sunlight'       ],
                              ['Rose'      , 'Flower'   , 'Medium'     , 'Full Sun'       ],    
                              ['Cactus'    , 'Succulent', 'Low'        , 'Full Sun'       ],  
                              ['Fern'      , 'Foliage'  , 'High'       , 'Partial Shade'  ],
                              ['Aloe Vera' , 'Succulent', 'Low'        , 'Bright Indirect'],
                              ['Tulip'     , 'Flower'   , 'Medium'     , 'Full Sun'       ]]}, 
}
```
![image](https://github.com/user-attachments/assets/1bd26afe-0c55-4490-9ef8-83cf68f541fc)
