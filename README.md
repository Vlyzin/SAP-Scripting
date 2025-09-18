Instale os pacotes com:

```bash
pip install -r requirements.txt

---

## Como Rodar

### Atualiza Remessas
- **Python:** `python atualiza_remessas.py`  
- **EXE:** Clique em `atualiza_remessas.exe` na pasta `dist/`  

> Precisa SAP GUI instalado e scripting habilitado.

### Outros Scripts
- **Python:** `python script_automacao1.py`  
- **Python:** `python script_automacao2.py`  

> Eles vão abrir janela para selecionar a planilha Excel.

---

## Logs

- `log_remessas.txt` → status de cada remessa processada  
- `log_erro.txt` → imagens não encontradas (somente para o script PyAutoGUI)

---

## Dicas Rápidas

- Não minimize o SAP ou a janela que será controlada pelo PyAutoGUI.  
- Se o script não achar alguma imagem, ele pode pular ou reiniciar a execução.  
- Mantendo a pasta `img/` junto do `.exe` tudo funciona sem precisar de Python.  
- Para encerrar rapidamente, pressione `ESC` (nos scripts que suportam).

