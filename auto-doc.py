
#Import relevant packages
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import datetime

#import pandas to read csv
import pandas as pd

#import docx template
    
template = DocxTemplate ('template-auto.docx')

#read csv

data = pd.read_csv('template-data.csv')

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
    'job': 'Werkstudentenstell',
    'type': 'Werkstudentstelle'
    #'type1':None,
    #'type2':None
}
for i in range(len(data)):

    #update context with info rows from data
    context['name'] = data.loc[i,'name']
    context['Company_umbrella'] = data.loc[i,'Company_umbrella']
    context['Company'] = data.loc[i,'Company']
    context['Location'] = data.loc[i,'Location']
    context['job'] = data.loc[i,'job']
    context['type'] = data.loc[i,'type']
    #render report
    
    #set template
    template = DocxTemplate ('template-auto.docx')
    
    #render doc
    
    template.render(context, autoescape=True)
    
    #save doc

    company_umbrella = context['Company_umbrella']
    name= context['name']
    savefile = f'ACM_Anschreiben_{name}'
    
    template.save(savefile+'.docx')
