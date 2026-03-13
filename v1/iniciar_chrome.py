import subprocess
import os

diretorio = os.path.dirname(os.path.abspath(__file__))

rodar_arquivo = os.path.join(diretorio, 'iniciar_robo.bat')

subprocess.run(rodar_arquivo, shell=True)