from docx import Document

document = open('Poyasnitelnaya_zapiska_Pleschev_Danil_021-1.docx', 'rb')
document = Document(document)


for i in range(len(document.paragraphs)):
	# print(i)
	
	if len(document.paragraphs[i].runs) > 0:
		
		for j in range(len(document.paragraphs[i].runs)):
			
			# print(i.runs[j].text == 'Таблица 1.1 – Функции меломана ')
			
			if document.paragraphs[i].runs[j].text == 'Таблица 1.1 – Функции меломана ':
				print(document.paragraphs[i].runs[j].text)
				print(document.tables[0].rows[0].cells[0].text)

	# print('\n\n')

