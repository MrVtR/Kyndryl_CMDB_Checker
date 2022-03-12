import subprocess
import sys

# Instalação de libs caso não estejam na máquina do desenvolvedor
try:
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pandas'])
finally:
    import pandas as pd

try:
    from pyfiglet import Figlet
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", 'pyfiglet'])
finally:
    from pyfiglet import Figlet

f = Figlet(width=100)
print(f.renderText("Script made by\nVitor R. G. Gomes - Kyndryl\n2022\n"))

# Import de arquivos necessários para rodar o script
AllJobsPath = 'ALLJOBS.txt'
CMDBJobsPath = 'CMDB.txt'
hfdCallListPath = 'hfd.xlsx'
hliCallListPath = 'hli.xlsx'
devCallListPath = 'dev.xlsx'
bldCallListPath = 'bld.xlsx'


print('Abrindo arquivos...')
with open(AllJobsPath, 'r') as f1:
    allCallLists = f1.read().split('\n')

with open(CMDBJobsPath, 'r') as f1:
    cmdb = f1.read().split('\n')

print("Verificando quais arquivos já estão no CMDB...")
AllCallListsMinusCMDB = [line for line in allCallLists if line not in cmdb]
print("Terceiro arquivo:", len(AllCallListsMinusCMDB))
f = open("AllCallListsMinusCMDB.txt", "a")
f.truncate(0)
f.write("Jobs que estão na Call List, mas não estão no CMDB ainda:\n")
for jobName in AllCallListsMinusCMDB:
    f.write(jobName+"\n")
f.close()

print("Abrindo Call Lists...")
hfdCallListDf = pd.read_excel(hfdCallListPath)
hliCallListDf = pd.read_excel(hliCallListPath)
devCallListDf = pd.read_excel(devCallListPath)
bldCallListDf = pd.read_excel(bldCallListPath)
print("hfd:", len(hfdCallListPath))
print("hli:", len(hliCallListPath))
print("dev:", len(devCallListPath))
print("bld:", len(bldCallListPath))
dfFinalResult = pd.DataFrame(columns=['Job Name', 'Instructions'])

print("Verificando quais Jobs estão nas Call Lists e pegando suas instruções...")
for idxJob, job in enumerate(AllCallListsMinusCMDB):
    print("Cont:", idxJob)
    achou = False
    detail = ''
    if job != '':
        if not achou:
            for i, j in enumerate(hfdCallListDf['Job']):
                if(j == job):
                    jobName = job
                    achou = True
                    for idx, val in enumerate(hfdCallListDf.iloc[i, :]):
                        if(idx > 0 and not pd.isna(val)):
                            detail += str(idx)+"- "+str(val)+'\n'
                    if not detail:
                        detail = 'Instructions not found'
                    break
        if not achou:
            for i, j in enumerate(hliCallListDf['Job']):
                if(j == job):
                    jobName = job
                    achou = True
                    for idx, val in enumerate(hliCallListDf.iloc[i, :]):
                        if(idx > 0 and not pd.isna(val)):
                            detail += str(idx)+"- "+str(val)+'\n'
                    if not detail:
                        detail = 'Instructions not found'
                    break
        if not achou:
            for i, j in enumerate(devCallListDf['Job']):
                if(j == job):
                    jobName = job
                    achou = True
                    for idx, val in enumerate(devCallListDf.iloc[i, :]):
                        if(idx > 0 and not pd.isna(val)):
                            detail += str(idx)+"- "+str(val)+'\n'
                    if not detail:
                        detail = 'Instructions not found'
                    break
        if not achou:
            for i, j in enumerate(bldCallListDf['Job']):
                if(j == job):
                    jobName = job
                    achou = True
                    for idx, val in enumerate(bldCallListDf.iloc[i, :]):
                        if(idx > 0 and not pd.isna(val)):
                            detail += str(idx)+"- "+str(val)+'\n'
                    if not detail:
                        detail = 'Instructions not found'
                    break
        if not achou:
            jobName = job
            detail = "Job not found in any call list"

        dfFinalResult.loc[dfFinalResult.shape[0]] = [jobName, detail]

dfFinalResult.to_excel("output.xlsx",
                       sheet_name='Output')

dfFinalResult.replace(r'\n', ' ', regex=True, inplace=True)

print("Escrevendo arquivo final...")
with open('RelacaoFinalCMDBCallList.txt', 'a', encoding="utf-8") as f:
    f.truncate(0)
    dfAsString = dfFinalResult.to_string(header=True, index=True)
    f.write(dfAsString)
print("Tudo pronto, finalizando programa...\nSee ya :D")
