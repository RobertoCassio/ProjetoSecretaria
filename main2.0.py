import csv
from docxtpl import DocxTemplate
from datetime import date


numprof = 0
profsup = 0
RangeCoord = 6
RangeSup = 0
DataHoje = date.today()
hoje = DataHoje.strftime("%d/%m/%Y")
print (hoje)


with open(r'declaracoes.csv', 'r', encoding='latin1') as csv_file:
    csv_reader = csv.DictReader(csv_file, delimiter=',')


    for line in csv_reader:

        if line['Nivel'] == 'M':
            RangeCoord = 4
        elif line ['Nivel'] == 'D':
            RangeCoord = 6

        if line['Prof 6'] != '':
            RangeCoord = RangeCoord+1
        else:  # Importante fazer voltar para 6 caso o próximo n tenha subcoordenador
            RangeCoord = RangeCoord  # Caso Coordenador esteja preenchido, o programa vai rodar mais uma vez pra fazer ele\Preciso editar os docs
        print(line)

        if line['SProf 3'] != '':
            RangeSup = 4
        else:
            RangeSup = 3
        for numprof in range(1, RangeCoord):
            print(numprof)
            if line['Nivel'] == 'M' and numprof == 3:
                numprof = numprof + 2
            def SexoCoorientador():
                global StrSexoCoo
                StrSexoCoo = ''
                SexoCoorient = line['Prof 6']
                if SexoCoorient == '':
                    print ('Texto Vazio')
                else:
                    SexoCoorient = SexoCoorient.split(None,1)[0]
                    print (SexoCoorient)
                    if SexoCoorient == 'Prof.':
                        StrSexoCoo = ' (Coorientador), '
                    elif SexoCoorient == 'Profa.':
                        StrSexoCoo = ' (Coorientadora), '


            def SexoOrientador():
                global StrSexoO
                SexoOrient = line['Prof 5']
                print(SexoOrient)
                SexoOrient = SexoOrient.split(None, 1)[0]
                print(SexoOrient)
                if SexoOrient == 'Prof.':
                    StrSexoO = 'Orientador'
                elif SexoOrient == 'Profa.':
                    StrSexoO = 'Orientadora'


            SexoOrientador()  # Para definir o sexo da palavra orientador
            Sexoprof = line[f'Prof {numprof}']
            Sexoprof = Sexoprof.split(None, 1)[0]  # .split para separar a primeira palavra None (parâmetro de separação pois usarei oe spaço mesmo, e 1 pois é pra separar 1ª palavra


            #  print (Sexoprof)
            Coordenacao = 'Maria Cristina Rosa'
            Cargo = 'Coordenadora'
            if line ['Prof 1'] == 'Profa. Dra. Maria Cristina Rosa' or ['Prof 2'] == 'Profa. Dra. Maria Cristina Rosa' or line['Prof 3'] == 'Profa. Dra. Maria Cristina Rosa' or line['Prof 4'] == 'Profa. Dra. Maria Cristina Rosa' or line['Prof 5'] == 'Profa. Dra. Maria Cristina Rosa' or line['Prof 6'] == 'Profa. Dra. Maria Cristina Rosa' or line['SProf 1'] == 'Profa. Dra. Maria Cristina Rosa' or line['SProf 2'] == 'Profa. Dra. Maria Cristina Rosa' or line['SProf 3'] == 'Profa. Dra. Maria Cristina Rosa':
                Coordenacao = 'Profa. Dra. Ana Paula Santos Guimarães'
                Cargo = 'Subcoordenadora'

            def TemplateUse():  #Defino sexo do aluno e template a ser usado
                global doc
                global nivel
                if line['Sexo A'] == 'M' and Sexoprof == 'Prof.':
                    doc = DocxTemplate(r"Template - AMOM.docx")  # aqui é o nome do arquivo template
                    if line['Nivel'] == 'D':
                        nivel = 'tese do doutorando'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação do mestrando' #Define sexo d oa
                elif line['Sexo A'] == 'M' and Sexoprof == 'Profa.':
                    doc = DocxTemplate(r"Template - AMOF.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese do doutorando'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação do mestrando'
                elif line['Sexo A'] == 'F' and Sexoprof == 'Prof.':
                    doc = DocxTemplate(r"Template - AFOM.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese da doutoranda'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação da mestranda'
                elif line['Sexo A'] == 'F' and Sexoprof == 'Profa.':
                    doc = DocxTemplate(r"Template - AFOF.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese da doutoranda'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação da mestranda'



            TemplateUse()
            SexoCoorientador()
            if numprof == 5:
                presidente = ' e presidente'
            else:
                presidente = ''
            def NomesUnis():
                global Uni3,Prof3,OUni3,OUni4
                if line['SUni 3'] != '':
                    Uni3 = line['SUni 3']
                    Uni3 = f'({Uni3})'
                else:
                    Uni3 = ''
                if line['SProf 3'] != '':
                    Prof3 = line['SProf 3']
                    Prof3 = f', {Prof3}'
                else:
                    Prof3 =''
                if line ['Uni 3'] != '':
                    OUni3 = line['Uni 3']
                    OUni3 = f' ({OUni3}), '
                else:
                    OUni3 = ''
                if line ['Uni 4'] != '':
                    OUni4 = line['Uni 4']
                    OUni4 = f'({OUni4})'
                else:
                    OUni4 = ''

            NomesUnis()
            def Ata():
                if numprof == 1:
                    doc = DocxTemplate('Template - Ata')
                    if line['Nivel'] == 'M':
                        onivel = 'Mestrem(a)'
                        nivel = 'MESTRADO'
                    elif line ['Nivel'] == 'D':
                        onivel = 'Doutor(a)'
                        nivel = 'DOUTORADO'
                        context = {'aluno': line['Nome'], 'Tese': line['Nome da Tese'], 'Orientador': line['Prof 5'],
                                   'Prof1': line['Prof 1'], 'Uni1': line['Uni 1'],
                                   'Prof2': line['Prof 2'], 'Uni2': f"({line['Uni 2']}), ", 'Prof3': line['Prof 3'],
                                   'Uni3': OUni3,
                                   'Prof4': line['Prof 4'], 'Uni4': OUni4,
                                   'data': line['Dia'], 'hora': line['Horario'],
                                   'SexoCoor': StrSexoCoo, 'Coorientador': line['Prof 6'],
                                   'nivel': nivel, 'hoje': hoje,
                                   'ONivel': onivel}  # 2 o que colocar aqui é o que vai substituir no texto
                        doc.render(context)
                        doc.save(fr"Divulgação\{line['Nome']}.docx")  # filename #


            def Divulgação():
                if numprof == 1:
                    doc = DocxTemplate(r"Template - Divulgação.docx")
                    if line['Nivel'] == 'M':
                        onivel = 'Mestrando(a)'
                        nivel = 'MESTRADO'
                    elif line ['Nivel'] == 'D':
                        onivel = 'Doutorando(a)'
                        nivel = 'DOUTORADO'


                    context = {'aluno': line['Nome'], 'Tese': line['Nome da Tese'], 'Orientador': line['Prof 5'],
                               'Prof1': line['Prof 1'], 'Uni1': line['Uni 1'],
                               'Prof2': line['Prof 2'], 'Uni2': f"({line['Uni 2']}), ", 'Prof3': line['Prof 3'],
                               'Uni3': OUni3,
                               'Prof4': line['Prof 4'], 'Uni4': OUni4,
                               'data': line['Dia'], 'hora': line['Horario'],
                               'Sup1': line['SProf 1'], 'USup1': line['SUni 1'],
                               'Sup2': line['SProf 2'], 'USup2': line['SUni 2'],
                               'SexoOr': StrSexoO, 'Sup3': Prof3, 'USup3': Uni3,
                               'SexoCoor': StrSexoCoo, 'Coorientador': line['Prof 6'],
                               'nivel': nivel, 'hoje': hoje, 'ONivel': onivel}  # 2 o que colocar aqui é o que vai substituir no texto
                    doc.render(context)
                    doc.save(fr"Divulgação\{line['Nome']}.docx")  # filename #
            Divulgação()


            context = {'aluno': line['Nome'], 'Tese': line['Nome da Tese'], 'Orientador': line['Prof 5'],
                       'Prof1': line['Prof 1'], 'Uni1': line['Uni 1'],
                       'Prof2': line['Prof 2'], 'Uni2': f"({line['Uni 2']}), ", 'Prof3': line['Prof 3'], 'Uni3': OUni3,
                       'Prof4': line['Prof 4'], 'Uni4': OUni4,
                       'data': line['Dia'], 'hora': line['Horario'], 'prof': line[f'Prof {numprof}'],
                       'Sup1': line['SProf 1'], 'USup1': line['SUni 1'],
                       'Sup2': line['SProf 2'], 'Usup2': line['SUni 2'],
                       'SexoOr': StrSexoO, 'Sup3': Prof3, 'USup3': Uni3,
                       'SexoCoor' : StrSexoCoo,'Coorientador' : line['Prof 6'], 'presidente': presidente, 'nivel':nivel,
                       'Coordenacao' : Coordenacao, 'Cargo' : Cargo, 'hoje' : hoje}  # 2 o que colocar aqui é o que vai substituir no texto
            doc.render(context)
            doc.save(fr"Declarações\{line ['Nome']} {line [f'Prof {numprof}']}.docx") #filename #
        for profsup in range(1, RangeSup):
            SexoOrientador()
            Sexosprof = line[f'SProf {profsup}']
            Sexosprof = Sexosprof.split(None, 1)[0]


            def TemplateUseSup():
                global doc
                global nivel
                if line['Sexo A'] == 'M' and Sexosprof == 'Prof.':
                    doc = DocxTemplate(r"Template - AMOM - Sup.docx")  # aqui é o nome do arquivo template
                    if line['Nivel'] == 'D':
                        nivel = 'tese do doutorando'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação do mestrando'
                elif line['Sexo A'] == 'M' and Sexosprof == 'Profa.':
                    doc = DocxTemplate(r"Template - AMOF - Sup.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese do doutorando'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação do mestrando'
                elif line['Sexo A'] == 'F' and Sexosprof == 'Prof.':
                    doc = DocxTemplate(r"Template - AFOM - Sup.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese da doutoranda'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação da mestranda'
                elif line['Sexo A'] == 'F' and Sexosprof == 'Profa.':
                    doc = DocxTemplate(r"Template - AFOF - Sup.docx")
                    if line['Nivel'] == 'D':
                        nivel = 'tese da doutoranda'
                    elif line['Nivel'] == 'M':
                        nivel = 'dissertação da mestranda'

            NomesUnis()
            TemplateUseSup()
            context = {'aluno': line['Nome'], 'Tese': line['Nome da Tese'], 'Orientador': line['Prof 5'],
                       'Prof1': line['Prof 1'], 'Uni1': line['Uni 1'],
                       'Prof2': line['Prof 2'], 'Uni2': f"({line['Uni 2']}), ", 'Prof3': line['Prof 3'], 'Uni3': OUni3,
                       'Prof4': line['Prof 4'], 'Uni4': OUni4,
                       'data': line['Dia'], 'hora': line['Horario'], 'prof': line[f'SProf {profsup}'],
                       'Sup1': line['SProf 1'], 'USup1': line['SUni 1'],
                       'Sup2': line['SProf 2'], 'Usup2': line['SUni 2'], 'SexoOr': StrSexoO, 'Sup3': Prof3, 'USup3': Uni3,
                       'SexoCoor' : StrSexoCoo,'Coorientador' : line['Prof 6'], 'nivel':nivel, 'Coordenacao' : Coordenacao, 'Cargo' : Cargo, 'hoje' : hoje}
            doc.render(context)
            doc.save(fr"Declarações\{line['Nome']} {line[f'SProf {profsup}']}.docx")
