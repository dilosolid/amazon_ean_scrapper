# amazon_ean_scrapper


Before running, run theses commands in the terminal

pip install python-amazon-product-api

pip install pandas

pip install xlrd

pip install openpyxl

pip install lxml


this script will read an excel file called data datasmall.xls by default (this can be change in the setting.py) and will read the columns in this format (EAN Description Volume Brand) and the script will use the amazon api to search the EAN in the excel file in the folowing amazon sites amazon.com amazon.ca amazon.com.mx amazon.co.uk amazon.de amazon.fr amazon.it amazon.es and will extract the ASIN number and the diferent titles in all locales and will save the data in another excel file.

to run type

python main.py 
