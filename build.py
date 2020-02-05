import csv
from docx import Document


document = Document()


with open('index1.tsv','r') as original_index:
    index_read = csv.reader(original_index, delimiter='\t')

    for line in index_read:
        para = document.add_paragraph('')
        para.add_run(line[0], style="Heading 2 Char")
        para.add_run(" [")
        para.add_run(line[1]).italic = True
        para.add_run("] ")
        para.add_run("\n")
        para.add_run(line[2])


document.save('buildTo2.docx')