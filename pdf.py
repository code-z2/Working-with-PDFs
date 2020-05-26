import PyPDF2
import sys
import os
import comtypes.client

with open('dummy.pdf', 'rb') as file:
    reader = PyPDF2.PdfFileReader(file)
    print(reader.numPages)
    page = reader.getPage(0)
    page.rotateClockwise(90)
    writer = PyPDF2.PdfFileWriter()
    writer.addPage(page)
    with open('rotated.pdf', 'wb') as rote:
        writer.write(rote)

# pdf merger
inputs = sys.argv[1:]


def merger(pdf_list):
    meges = PyPDF2.PdfFileMerger()
    for pdf in pdf_list:
        meges.append(pdf)
    meges.write('meged.pdf')


# merger(inputs)


# word to pdf converter

wdFormatPDF = 17
filename = input('enter name of word document')+'.docx'
in_file = os.path.abspath(filename)
clean_name = os.path.splitext(filename)[0]
out_file = os.path.abspath(f'{clean_name}.pdf')

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

# water mark all page in a PDF

supa = PyPDF2.PdfFileReader(open('super.pdf', 'rb'))
wtr = PyPDF2.PdfFileReader(open('wtr.pdf', 'rb'))
output = PyPDF2.PdfFileWriter()
count = supa.getNumPages()
mark = 0
while mark < count:
    page = supa.getPage(mark)
    page.mergePage(wtr.getPage(0))
    output.addPage(page)

    with open('watermarked.pdf', 'wb') as file:
        output.write(file)
    mark += 1
