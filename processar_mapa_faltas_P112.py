# -*- coding: utf-8 -*-
"""
Created on Wed Nov  8 12:22:17 2023

@author: rui.fartaria
"""
import sys
import xlrd
import zipfile
import xlsxwriter  

def colindex(colA):
    colA = colA.upper()
    if len(colA) == 1:
        return ord(colA[0])-ord('A')
    else:
        leftDigit = colA[0]
        rightDigit = colA[1]
        return (ord(leftDigit)-ord('A'))**2+26+(ord(rightDigit)-ord('A'))

def process_zip(fpath):
    disciplinas = []
    with zipfile.ZipFile(fpath, mode="r") as archive:
        #archive.printdir()
        for filename in archive.namelist():
            with archive.open(filename, 'r') as xls:
                xls_bytes = xls.read()
                disciplinas.append(process_xls(xls_bytes))
    return disciplinas

def process_xls(xls_bytes):
    iF = colx=colindex('F')
    iB = colx=colindex('B')
    iC = colx=colindex('C')
    
    book = xlrd.open_workbook(file_contents=xls_bytes)
    sh = book.sheet_by_index(0)
    
    disciplina = sh.cell_value(rowx=13-1, colx=iB).replace('Disciplina:','').strip()  # B13

    # MODULOS
    modulos = []
    
    # identificação dos módulos
    modulos_col = []
    count_empty = 0
    for i in range(200):
        mstr = sh.cell_value(rowx=16-1, colx=iF + i)
        if not mstr:
            count_empty += 1
            if count_empty > 10:
                break
            continue
        count_empty = 0
        if 'Mod' in mstr:
            modulos_col.append(iF + i)
        else:
            raise 'Erro na indentificação dos módulos.'
    
    for i in modulos_col:
        mstr = sh.cell_value(rowx=16-1, colx=i)
        if not mstr:
            raise 'Erro na indexação de colunas dos módulos.'
        mstr = mstr.split(' ')
        modulos.append({
            'num':int(mstr[0].replace('Mod.','').strip()), 
            'tempos':int(mstr[1].replace('(','').replace(')','').replace('T',''))
            })
    
    #print(modulos)
    
    # identificação das colunas FI, FJ
    cols_FI_FJ = []
    row_FI_FJ = 19-1
    for i in modulos_col:
        FI_FJ = []
        for j in range(4):
            fstr = sh.cell_value(rowx=row_FI_FJ, colx=i+j)
            #print(fstr)
            if fstr.strip() == 'FI':
                FI_FJ.append(i+j)
            if fstr.strip() == 'FJ':
                FI_FJ.append(i+j)
            if len(FI_FJ) == 2:
                break
        if len(FI_FJ) != 2:
            #print(FI_FJ)
            raise 'Erro na identificação de colunas FI FJ'
        cols_FI_FJ.append(FI_FJ)

    #print()
    #print(cols_FI_FJ)

    # ALUNOS
    l0 = 22-1
    alunos = {}
    for i in range(50):
        l = l0 + i
        numero = sh.cell_value(rowx=l, colx=iB)
        
        if not numero:
            continue
        
        if type(numero) is float:
            numero = int(numero)
            nome = sh.cell_value(rowx=l, colx=iB+1)
            faltas = dict()
            for i, mc in enumerate(cols_FI_FJ):
                #print(mc)
                _FI = sh.cell_value(rowx=l, colx=mc[0])
                _FJ = sh.cell_value(rowx=l, colx=mc[1])
                _FI = int(_FI) if type(_FI) is float else 0
                _FJ = int(_FJ) if type(_FJ) is float else 0
                faltas[modulos[i]['num']] = {'FI':_FI, 'FJ':_FJ}
            alunos[numero] = {
                'num':numero,
                'nome':nome,
                'faltas': faltas
                }
        elif 'Legenda' in numero:
            break
    
    return {
        'disciplina': disciplina,
        'modulos': modulos,
        'alunos': alunos
        }    

def calcular_pap_papr(disciplinas):
    
    # calcular limites
    for d in disciplinas:
        for m in d['modulos']:
            limite_f = m['tempos'] * 0.1
            limite_i = int(limite_f)
            m['limite'] = limite_i+1 if limite_f > limite_i else limite_i
        #print(d['disciplina'])
        #print(d['modulos'])
    
    for d in disciplinas:
        #print(d['disciplina'])
        PAR = dict([(k['num'],[]) for k in d['modulos']])
        PAPr = dict([(k['num'],[]) for k in d['modulos']])
        for ak,a in d['alunos'].items():
            for mk,f in a['faltas'].items(): # mk é o número do módulo
                limite = [m['limite'] for m in d['modulos'] if m['num']==mk][0]
                if f['FI'] > limite or (f['FI']+f['FJ']>limite and f['FI']>f['FJ']):  # I>L or (I+J>L and I>j)
                    PAR[mk].append(a['num'])
                if f['FI']+f['FJ']>limite and f['FI']<f['FJ']:   # I+J>L and J>I
                    PAPr[mk].append(a['num'])
        d['PAR'] = PAR
        d['PAPr'] = PAPr


def escrever_ficheiro_XLSX(disciplinas, fpath):
    with xlsxwriter.Workbook(fpath) as workbook:
        sh = workbook.add_worksheet()
        
        # formats
        cf_1 = workbook.add_format({'bg_color': '#FFFF99'})
        cf_2 = workbook.add_format({'bg_color': '#CCECFF'})
        
        cf_1d = workbook.add_format({'bold':True, 'bg_color': '#FFFF99'})
        cf_2d = workbook.add_format({'bold':True, 'bg_color': '#CCECFF'})
        
        # TABELA DE FALTAS AS DISCIPLINAS
        col_discip_base = 3
        row_discip_base = 2
        
        col_mod_base = col_discip_base
        row_mod_base = row_discip_base + 1
        
        col_lf_base = col_discip_base
        row_lf_base = row_discip_base + 2
        
        col_IJ_base = col_discip_base
        row_IJ_base = row_discip_base + 3
        
        col_numAluno_base = col_discip_base - 2
        row_numAluno_base = row_discip_base + 4
        
        # escrever cabeçalho: diciplinas
        k = col_discip_base
        cf = cf_1d
        for d in disciplinas:
            #sh.write(row_discip_base, k, d['disciplina'])
            mergeLen = len(d['modulos']) * 2
            cf = cf_1d if cf==cf_2d else cf_2d
            sh.merge_range(row_discip_base, k, row_discip_base, k+mergeLen-1, d['disciplina'], cf)
            k += mergeLen
        
        # escrever cabeçalho: modulos
        sh.write(row_mod_base, col_mod_base-1, 'Módulos')
        k = col_mod_base
        cf = cf_1
        for d in disciplinas:
            cf = cf_1 if cf==cf_2 else cf_2
            for m in d['modulos']:
                #sh.write(row_mod_base, k, m['num'])
                sh.merge_range(row_mod_base, k, row_mod_base, k+2-1, m['num'], cf)
                k += 2
        
        # escrever cabeçalho: limite de faltas
        sh.write(row_lf_base, col_lf_base-1, 'Limite de Faltas')
        k = col_lf_base
        cf = cf_1
        for d in disciplinas:
            cf = cf_1 if cf==cf_2 else cf_2
            for m in d['modulos']:
                #sh.write(row_lf_base, k, m['limite'])
                sh.merge_range(row_lf_base, k, row_lf_base, k+2-1, m['limite'], cf)
                k += 2
        
        # escrever cabeçalho: I J
        k = col_IJ_base
        cf = cf_1
        for d in disciplinas:
            cf = cf_1 if cf==cf_2 else cf_2
            for m in d['modulos']:
                sh.write(row_IJ_base, k, 'I', cf)
                k += 1
                sh.write(row_IJ_base, k, 'J', cf)
                k += 1
        
        # escrever nomes dos alunos (ordenados por numero)
        i = row_numAluno_base
        for ak in sorted(disciplinas[0]['alunos'].keys()):
            sh.write(i+ak-1, col_numAluno_base, ak)
            sh.write(i+ak-1, col_numAluno_base+1, disciplinas[0]['alunos'][ak]['nome'])

        # escrever faltas dos alunos
        k = col_discip_base
        cf = cf_1
        for d in disciplinas:
            cf = cf_1 if cf==cf_2 else cf_2
            for ak in sorted(d['alunos'].keys()):
                aluno = d['alunos'][ak]
                for im,m in enumerate(d['modulos']):
                    sh.write(row_numAluno_base+ak-1, k + im*2, "" if aluno['faltas'][m['num']]['FI']==0 else aluno['faltas'][m['num']]['FI'], cf)
                    sh.write(row_numAluno_base+ak-1, k + im*2 + 1, "" if aluno['faltas'][m['num']]['FJ']==0 else aluno['faltas'][m['num']]['FJ'], cf)
            k += len(d['modulos']) * 2
        
        # TABELA DE PAP PAPr
        col_PAPr_base = col_discip_base + sum([len(d['modulos']) for d in disciplinas])*2 + 1
        row_PAPr_base = row_numAluno_base - 1
        
        # escrever cabeçalho: PAPr, PAR
        cf_PAPr = workbook.add_format({'bold':True,'bg_color': '#C6E0B4'})    
        cf_PAR = workbook.add_format({'bold':True,'bg_color': '#FFCCCC'})    
        
        sh.write(row_PAPr_base, col_PAPr_base, 'PAPr', cf_PAPr)
        sh.write(row_PAPr_base, col_PAPr_base+1, 'PAR', cf_PAR)
        for ak in sorted(disciplinas[0]['alunos'].keys()):
            d_PAPr = [d['disciplina']+' (M'+str(m['num'])+'); ' for d in disciplinas for m in d['modulos'] if ak in d['PAPr'][m['num']]]
            d_PAR = [d['disciplina']+' (M'+str(m['num'])+'); ' for d in disciplinas for m in d['modulos'] if ak in d['PAR'][m['num']]]
            sh.write(row_PAPr_base+1+(ak-1), col_PAPr_base, ''.join(d_PAPr))
            sh.write(row_PAPr_base+1+(ak-1), col_PAPr_base+1, ''.join(d_PAR))

        sh.autofit()

if __name__ == "__main__":
    #zip_path = "Faltas_Disciplinas_Mod_P112.zip"
    
    if len(sys.argv) != 2 or not sys.argv[1].endswith('.zip'):
        print('Needs a zip file')
        sys.exit(1)
    zip_path = sys.argv[1]
    
    disciplinas = process_zip(zip_path)
    calcular_pap_papr(disciplinas)
    escrever_ficheiro_XLSX(disciplinas, zip_path+'.xlsx')
    
    