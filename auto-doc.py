
#Import relevant packages
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import datetime
import os
#import pandas to read csv
import pandas as pd
#import convert to pdf
#from docx2pdf import convert

#import docx template
    
#template = DocxTemplate ('lorem-template.docx')

#current directory
cwd = os.getcwd()


#read csv

data = pd.read_csv('template-cl.csv')

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
    'type': 'Werkstudentstelle',
    'subject': 'Bewerbung',
    'start':'sofort'
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
    context['subject'] = data.loc[i,'subject']
    context['start'] = data.loc[i,'start']
    #render report
    
    #set template
    template = DocxTemplate ('template-cl.docx')
    
    #render doc
    
    template.render(context, autoescape=True)
    
    #save doc

    company_umbrella = context['Company_umbrella']
    name= context['name']
    savefile = f'Example_Generated_{name}_{i}'
    
    template.save(cwd+'/generated/'+savefile+'.docx')

#convert rendered docx to pdf

#convert(cwd+'/generated/', cwd+'/generted2pdf')