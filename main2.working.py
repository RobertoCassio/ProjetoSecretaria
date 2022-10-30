import csv
from docxtpl import DocxTemplate

#termom = 'Prof.'
#termof= 'Profa.'
numprof = 0
with open(r'C:\Users\Usuário\Desktop\Teste Phyton\declaracoes.csv', 'r', encoding='latin1') as csv_file:
    csv_reader = csv.DictReader(csv_file, delimiter=';')
    for line in csv_reader:
        print(line)
    for numprof in range (1,6):
        print (numprof)
        # Sexoprof = line [f'Prof {numprof}']
        # Sexoprof.split (None,1) [0]
        # if line['Sexo A'] == 'M' and Sexoprof == 'Dr.':
        doc = DocxTemplate(r"C:\Users\Usuário\Desktop\Teste Phyton\Template - DecPart.docx") #aqui é o nome do arquivo template
        context = {'aluno': line['Nome'], 'Tese' : line ['Nome da Tese'], 'Orientador' : line ['Prof 5'], 'Prof1' : line ['Prof 1'], 'Uni1' : line['Uni 1'],
                   'Prof2' : line ['Prof 2'], 'Uni2' : line ['Uni 2'], 'Prof3' : line ['Prof 3'], 'Uni3' : line['Uni 3'], 'Prof4' : line['Prof 4'], 'Uni4' : line['Uni 4'],
                   'data': line['Dia'], 'hora' : line['Horario'], 'prof' : line[f'Prof {numprof}']}  # o que colocar aqui é o que vai substituir no texto
        doc.render(context)
        doc.save(fr"C:\Users\Usuário\Desktop\Teste Phyton\Declaracao de Participação {line [f'Prof {numprof}']}.docx") #filename
