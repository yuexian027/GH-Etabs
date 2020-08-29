# Grasshopper and Etabs
## Background
   How to make Grasshopper talk to etabs?

## Result
![Result](Image/Result.png)

## Workflow
![Workflow](Image/Workflow.png)
1. Save Grid Info of the building in an Excel File ([Grid Info](Grasshopper))
2. Draw a Rhino by using Grasshopper ([Sample Model](Grasshopper))
3. Save model information in Excel file, which is generated by the grasshopper file([aa](Grasshopper))
4. Use this component in Grasshopper to build a model in Etabs which is based on Etabs API


## Prerequisites
* Grasshopper
   * Install Python package (xlrd) to read Excel file ([xlrd.PyPI](https://pypi.org/project/xlrd/#files))
   ~~~~
   pip install xlrd-1.2.0-py2.py3-none-any.whl
   ~~~~
   * Install Python package (xlsxwriter) to write to Excel file ([xlsxwriter.PyPI](https://pypi.org/project/XlsxWriter/))
   ~~~~
   pip install XlsxWriter-1.3.2-py2.py3-none-any.whl
   ~~~~ 
   * Install Python package to Grasshopper
      1. In Rhino type in the command "EditPythonScript".
      2. In this Python editor go Tools -> Options -> Files. Then you will see an overview of current paths used
      3. Move the folders xlrd and xlsxwriter to one of the directories.



   
### Visaul Studio ()
