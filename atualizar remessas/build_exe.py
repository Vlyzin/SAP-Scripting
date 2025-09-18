import subprocess
import sys
import os

# Caminho absoluto do script que você quer transformar em exe
script_path = os.path.join(os.getcwd(), "atualizar remessas", "atualiza_remessas.py")

# Verifica se o script existe
if not os.path.exists(script_path):
    print(f"ERRO: Script não encontrado em {script_path}")
    sys.exit(1)

# Comando do PyInstaller
cmd = [
    sys.executable, 
    "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--distpath", os.path.join(os.getcwd(), "dist"),
    f'"{script_path}"' 
]

print("Rodando PyInstaller...")
subprocess.run(" ".join(cmd), shell=True, check=True)

print("Build finalizado! .exe está na pasta 'dist'")
