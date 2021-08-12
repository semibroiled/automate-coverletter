
#Import relevant packages
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import datetime

#import docx template

template = DocxTemplate ('template-auto.docx')

#declare template variables in a dict
#filled with placeholders
context = {
    'day': datetime.datetime.now().strftime('%d'),
    'month': datetime.datetime.now().strftime('%B'),
    'year': datetime.datetime.now().strftime('%Y'),

    'name': 'Max Mustermann',
    'Company_umbrella': 'Frauenhofer',
    'Company': 'IKS',
    'Location': 'Darmstadt',
    'job': 'Werkstudentenstell'
    #'type1':None,
    #'type2':None
}

#render report
template.render(context)
company_umbrella = context['Company_umbrella']
name= context['name']
savefile = f'{ company_umbrella }_{ name }'
print(savefile)
template.save(savefile+'.docx')
