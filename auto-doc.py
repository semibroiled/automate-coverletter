
#Import Relevant Packages

from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import datetime
import os

#Import pandas to read CSV file
import pandas as pd

#import convert to pdf
#from docx2pdf import convert

#import docx template
#template = DocxTemplate ('lorem-template.docx')

#Current Directory
cwd = os.getcwd()

print(f'Reading list...')
#Read CSV
path = cwd+'/'
filename = 'template-cl.csv'
data = pd.read_csv(path+filename)

#Declare template variables in a dict
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

print(f'Running loop over {filename}...')
#Run loop over List from CSV
for i in range(len(data)):
    
    print(f'Generating context...')
    #Update context with information from datafile
    context['name'] = data.loc[i,'name']
    context['Company_umbrella'] = data.loc[i,'Company_umbrella']
    context['Company'] = data.loc[i,'Company']
    context['Location'] = data.loc[i,'Location']
    context['job'] = data.loc[i,'job']
    context['type'] = data.loc[i,'type']
    context['subject'] = data.loc[i,'subject']
    context['start'] = data.loc[i,'start']
    
    print('Importing template...')
    #set template
    template = DocxTemplate ('template.docx')
    
    print(f'Populating document...')
    #Render document
    template.render(context, autoescape=True)
    
    print(f'Saving file...')
    #save doc
    company_umbrella = context['Company_umbrella']
    name= context['name']
    savefile = f'{i}_ACM_Anschreiben_{name}'
    
    template.save(cwd+'/generated/'+savefile+'.docx')
    print(f'{savefile} rendered!')

#convert rendered docx to pdf

#convert(cwd+'/generated/', cwd+'/generted2pdf')