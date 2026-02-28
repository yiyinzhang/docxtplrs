from docxtplrs import DocxTemplate

doc = DocxTemplate('template.docx')

context = {
    "company_name": "Example Corp",
    "user_name": "John Doe",
}

doc.render(context)
doc.save('output.docx')
