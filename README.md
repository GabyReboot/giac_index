# giac_index
A script for formatting large .tsv GIAC indexes to easier to read, space efficient docx files. 

Here's what it looks like!


<img src="outputEx.png" width="500">

### Requirements

[Python-docx](https://pypi.org/project/python-docx/)

### Expected Format

**Rows** of your index should be in alphabetical order (*the script will not do this for you!*).

**Columns** of your index should be in the following order: Term, Page Number, Description

### Prettifying

Name your .tsv file 'index1.tsv' and place it into the same folder as this script.

Then naviagte to the folder, run `python build.py`, and a new file named 'index_finalized.docx' will appear like a wild pokemon! 


Good luck on your exam!! 

