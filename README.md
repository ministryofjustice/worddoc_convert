worddoc_convert
===============

##This script does 2 things:
 - downloads .docs and .rtf files from a CSV list
 - converts them to PDF and HTML files


##How to use it:

- Make sure you have Libra Office or Open Office installed
- Make sure you have Python installed
- If you are using OSX and have Libra office installed in the standard place, run:
         
    python worddoc_convert.py -i [path to csv file] -o [path of output directory]

- Otherwise run:
    
    python worddoc_convert.py -i [path to csv file] -o [path of output directory] -c [location of soffice executable]
