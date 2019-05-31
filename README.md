Installation


How to use

0) Put all the marksheets in the current directory

1) install python 3.6, numpy, pandas statsmodel,seaborn and matplotlib. This could be done by running 
  pip3 install -r requirements.txt
or
  conda install numpy pandas statsmodel matplotlib seaborn

2) create a "metadata" excel file with the first 5 columns titled "Filename", "Start row", "Number of students", "Marks column" and "Short name". We will assume that this file is called "metadata.xlsx"

3) In each row, put the filename of the marksheet in column 1, the first row in which students appear in column 2, and the total number of students on the front sheet of the marksheet in column 3. Column 4 must be the column where the total mark is displayed (e.g., "AM") _AS A 0-22 CGS MARK_, and the 5th column should be the short name for the course, e.g., "CS1024"

4) run the following command "python marks.py metadata.xlsx"
