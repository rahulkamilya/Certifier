import csv
from docxtpl import DocxTemplate

# Configuration
template = DocxTemplate('certificate-template.docx')
filename = 'data.csv'
output = '/certifications'

# get all rows
getList = []

# Open and reading csv
with open(filename, 'r') as data:
    for line in csv.reader(data, delimiter=','):
        getList.append(line)


# Function to create files .docx
def create_certification():
    for names in getList[2:]:
        # Column's Name
        date = names[0]
        name = names[1]
        head_event = names[2]
        mentor = names[3]
        context = {
            'date': date,
            'name': name,
            'head_event': head_event,
            'mentor': mentor,
        }
        template.render(context)
        template.save(f".{output}/{name}.docx")


# Run function
create_certification()