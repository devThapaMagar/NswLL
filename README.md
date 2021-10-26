New South Wales liquor Logistics Data Manipulation from Unstructured PDF

Author

Dev Thapa Magar

Introduction

This is an Application which takes different type of unstructured data from the PDF and converts the data into sets of CSV. Since, the data were unstructured; the data are retrieved via help of the co-ordinate of the data in the document. PymuPDF was used to get the value of the coordinates which helped in retrieving the necessary data for the company.
For the current project, 4 companies invoice are uploaded into the application where the application uses ABN to separate the files from one another. 

Requirement

Python & PIP

Toolkit:

PyMuPDF: Python bindings for the PDF toolkit and renderer MuPDF
Pandas: Powerful data structures for data analysis, time series, and statistics
Pdfplumber: Plumb a PDF for detailed information about each char, rectangle, and line.

Ways to install the toolkits:

pip install pandas

pip install pdfplumber

pip install PyMuPDF==1.18.19


