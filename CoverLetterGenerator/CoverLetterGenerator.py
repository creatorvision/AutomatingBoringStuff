import datetime
from docx import Document


def docx_replace_regex(doc_obj, regex1, regex2, replace1, replace2):
    for p in doc_obj.paragraphs:
        words = p.text.split()
        if regex1 in words:
            indexes = words.index(regex1)
            words[indexes] = replace1
            p.text = ""
            for w in words:
                p.text = p.text + w + " "

        if regex2 in words:
            indexes = words.index(regex2)
            words[indexes] = replace2
            p.text = ""
            for w in words:
                p.text = p.text + w + " "


Company = ["Adobe", "ARM", "Google", "Facebook"]
# Just add companies for which you want the cover letter for, in the above declared List.

for company in Company:
    regex1 = "Company_name"
    regex2 = "_Date_"
    replace1 = company
    date = datetime.datetime.now()
    replace2 = str(date.day) + ' /' + str(date.month) + ' /' + str(date.year)
    filename = "template.docx"
    doc = Document(filename)
    docx_replace_regex(doc, regex1, regex2, replace1, replace2)
    doc.save('CoverLetter_' + company + '.docx')
