from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

invoice_list = [[]]


doc.render({"invoicelist": invoice_list})
doc.save("new_invoice.docx")