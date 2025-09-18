import time
import pyautogui
import openpyxl
import os
import sys
import keyboard
 
def escolher_arquivo():
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename
    Tk().withdraw()
    filename = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return filename
 
def ler_planilha(filename):
    wb = openpyxl.load_workbook(filename, data_only=True)
    ws = wb['Planilha2']
    dados = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1] and row[2]:
            dados.append({
                'material': str(row[0]),
                'lote': str(row[1]),
                'peneira': str(row[2])
            })
    return dados
 
def log_erro(material, lote):
    with open("log_erro.txt", "a") as f:
        f.write(f"Imagem não encontrada para Material: {material} - Lote: {lote}\n")
 
def localizar_imagem(caminho, conf=0.9):
    try:
        return pyautogui.locateOnScreen(caminho, confidence=conf)
    except pyautogui.ImageNotFoundException:
        return None
 
def localizar_imagem_centro(caminho, conf=0.9):
    try:
        return pyautogui.locateCenterOnScreen(caminho, confidence=conf)
    except pyautogui.ImageNotFoundException:
        return None
 
def reiniciar_execucao():
    print("Reiniciando execução do script...")
    python = sys.executable
    os.execl(python, python, *sys.argv)
 
def verificar_escape():
    """Encerra o script se a tecla ESC for pressionada"""
    if keyboard.is_pressed("esc"):
        print("\nExecução interrompida pelo usuário (ESC pressionado).")
        sys.exit()
 
def main():
    arquivo = escolher_arquivo()
    if not arquivo:
        print("Nenhum arquivo selecionado. Saindo.")
        return
 
    dados = ler_planilha(arquivo)
    print(f"Total de registros: {len(dados)}")
    print("Coloque a tela do SAP na frente. Começando em 5 segundos...")
    time.sleep(5)
 
    caminho_apareceu = os.path.join(img_dir, "apareceu.png")
    caminho_classe = os.path.join(img_dir, "classe.png")
    caminho_sim = os.path.join(img_dir, "sim.png")
    caminho_peneira = os.path.join(img_dir, "peneira.png")
    caminho_nlote = os.path.join(img_dir, "nlote.png")
 
    for item in dados:
        verificar_escape()  # <-- Checa ESC no início do loop
 
        material = item['material']
        lote = item['lote']

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace') 
        time.sleep(0.3)
        pyautogui.write(material, interval=0.05)
        pyautogui.press('tab')
 
        verificar_escape()

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace') 
        time.sleep(0.3) 
        pyautogui.write(lote, interval=0.05)
        pyautogui.press('enter')
 
        verificar_escape()
 
        apareceu = None
        timeout = 3
        start = time.time()
        while time.time() - start < timeout:
            verificar_escape()
            apareceu = localizar_imagem(caminho_apareceu, conf=0.9)
            if apareceu is not None:
                break
            time.sleep(0.5)
 
        if apareceu is None:
            print(f"Imagem 'apareceu.png' não encontrada para Material {material} e Lote {lote}. Pulando...")
            log_erro(material, lote)
            continue
 
        pos_classe = localizar_imagem_centro(caminho_classe, conf=0.9)
        if pos_classe:
            pyautogui.click(pos_classe)
            print(f"Clicado em classe.png para Material {material} e Lote {lote}.")
            time.sleep(0.5)
 
            for _ in range(7):
                verificar_escape()
                pyautogui.press('down')
                time.sleep(0.6)
 
            for _ in range(3):
                verificar_escape()
                pyautogui.keyDown('shift')
                pyautogui.press('down')
                pyautogui.keyUp('shift')
                time.sleep(0.9)
 
            for _ in range(5):
                verificar_escape()
                pyautogui.press('down')
                time.sleep(0.6)
 
            pos_peneira = localizar_imagem_centro(caminho_peneira, conf=0.9)
            if not pos_peneira:
                print(f"Imagem 'peneira.png' não encontrada para Material {material} e Lote {lote}. Reiniciando execução...")
                log_erro(material, lote)
                pyautogui.press('f3')
                time.sleep(1)
                reiniciar_execucao()
 
            peneira = item['peneira']
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(peneira, interval=0.05)
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.6)
 
            for _ in range(27):
                verificar_escape()
                pyautogui.keyDown('shift')
                pyautogui.press('down')
                pyautogui.keyUp('shift')
                time.sleep(0.9)
 
            pos_nlote = localizar_imagem_centro(caminho_nlote, conf=0.9)
            if not pos_nlote:
                print(f"Imagem 'nlote.png' não encontrada para Material {material} e Lote {lote}. Reiniciando execução...")
                pyautogui.press('f3')
                time.sleep(1)
                reiniciar_execucao()
 
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(lote, interval=0.05)
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.6)
 
            pyautogui.press('down')
            time.sleep(0.6)
 
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(lote, interval=0.05)
            time.sleep(0.6)
 
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)
 
            pos_sim = localizar_imagem_centro(caminho_sim, conf=0.9)
            if pos_sim:
                pyautogui.click(pos_sim)
                print(f"Clicado em sim.png para Material {material} e Lote {lote}.")
            else:
                print(f"Imagem 'sim.png' não encontrada para Material {material} e Lote {lote}.")
 
        else:
            print(f"Imagem 'classe.png' não encontrada na tela para Material {material} e Lote {lote}.")
 
        time.sleep(1)
 
if __name__ == "__main__":
    main()
    