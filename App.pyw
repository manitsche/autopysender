import requests
from io import BytesIO
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
import tkinter as tk
from tkinter import messagebox
import threading
import os
from datetime import datetime

# ==============================
# CONFIG
# ==============================

FILE_ID = "1FwhozdQrgbIpfeUN9oMJjX1H1DD07SQj"
ABA = "Página1"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

status_label = None
pausado = False
linha_atual = 0
total_linhas = 0

# ==============================
# UTIL
# ==============================

def caminho_arquivo(nome):
    return os.path.join(BASE_DIR, nome)

# ==============================
# LOG
# ==============================

def log_erro(nome, tipo_erro, mensagem):
    try:
        with open(caminho_arquivo("Erro.log"), "a", encoding="utf-8") as f:
            f.write(
                f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] "
                f"| Linha: {linha_atual}/{total_linhas} "
                f"| Nome: {nome} | Tipo: {tipo_erro} | Erro: {mensagem}\n"
            )
    except:
        pass

def atualizar_status(msg, erro=False):
    try:
        if status_label:
            status_label.config(
                text=f"{msg}",
                fg="red" if erro else "blue"
            )
    except:
        pass

# ==============================
# PAUSA
# ==============================

def aguardar_se_pausado():
    global pausado
    while pausado:
        time.sleep(1)

# ==============================
# TELEFONE
# ==============================

def tratar_telefone(valor):
    try:
        telefone = str(valor).strip()
        if telefone.endswith(".0"):
            telefone = telefone[:-2]
        telefone = ''.join(filter(str.isdigit, telefone))
        if not telefone.startswith("55"):
            telefone = "55" + telefone
        return telefone
    except:
        return None

# ==============================
# PLANILHA
# ==============================

def carregar_planilha():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
        r = requests.get(url)
        r.raise_for_status()
        return openpyxl.load_workbook(BytesIO(r.content))
    except Exception as e:
        log_erro("GERAL", "PLANILHA", str(e))
        atualizar_status("Erro ao carregar planilha", True)
        return None

# ==============================
# MENSAGEM
# ==============================

def carregar_mensagem():
    try:
        with open(caminho_arquivo("mensagem.txt"), "r", encoding="utf-8") as f:
            return f.read()
    except:
        return "Olá {nome}, mensagem automática."

def salvar_mensagem(texto):
    with open(caminho_arquivo("mensagem.txt"), "w", encoding="utf-8") as f:
        f.write(texto)

def gerar_mensagem(nome):
    return carregar_mensagem().replace("{nome}", nome)

# ==============================
# SELENIUM
# ==============================

def iniciar_driver():
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument(r"--user-data-dir=C:\autopysender\selenium\perfil")

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

        driver.get("https://web.whatsapp.com/")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "pane-side"))
        )

        return driver

    except Exception as e:
        log_erro("GERAL", "DRIVER", str(e))
        atualizar_status("Erro ao iniciar navegador", True)
        return None

# ==============================
# ENVIO
# ==============================

def enviar_mensagem(driver, nome, telefone):
    aguardar_se_pausado()

    try:
        mensagem = gerar_mensagem(nome)

        driver.get(f"https://web.whatsapp.com/send?phone={telefone}")

        caixa = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//footer//div[@contenteditable="true"]'))
        )

        aguardar_se_pausado()

        caixa.click()
        caixa.send_keys(mensagem)
        caixa.send_keys(Keys.ENTER)

        atualizar_status(f"Enviado: {nome}")

    except Exception as e:
        log_erro(nome, "ENVIO", str(e))
        atualizar_status(f"Erro ao enviar para {nome}", True)

# ==============================
# BOT
# ==============================

def executar_bot():
    global linha_atual, total_linhas

    driver = iniciar_driver()
    if not driver:
        return

    workbook = carregar_planilha()
    if not workbook:
        return

    aba = workbook[ABA]

    linhas = list(aba.iter_rows(min_row=2))
    total_linhas = len(linhas)
    linha_atual = 0

    for i, linha in enumerate(linhas, start=1):

        aguardar_se_pausado()
        linha_atual = i

        try:
            nome = linha[0].value
            telefone = tratar_telefone(linha[1].value)

            if nome and telefone:
                atualizar_status(f"Enviando para {nome}")
                enviar_mensagem(driver, nome, telefone)
                time.sleep(random.uniform(5, 10))
            else:
                raise Exception("Dados inválidos")

        except Exception as e:
            log_erro(nome if nome else "SEM NOME", "LINHA", str(e))
            atualizar_status(f"Erro na linha {i}", True)

    driver.quit()
    atualizar_status("Finalizado")

# ==============================
# LOOP
# ==============================

def loop_automatico():
    while True:
        executar_bot()

        for _ in range(300):
            aguardar_se_pausado()
            time.sleep(1)

# ==============================
# INTERFACE
# ==============================

class Interface:

    def __init__(self, root):
        global status_label

        self.root = root

        root.title("AutoPySender")
        root.geometry("500x420")

        tk.Label(root, text="Mensagem:").pack()

        self.texto = tk.Text(root, height=6)
        self.texto.pack()
        self.texto.insert("1.0", carregar_mensagem())

        tk.Button(root, text="Salvar Mensagem", bg="#007BFF", fg="white",
                  command=self.salvar).pack(pady=10)

        status_label = tk.Label(root, text="Rodando automático...", fg="blue")
        status_label.pack()

        self.label_contador = tk.Label(root, text="Linhas: 0/0")
        self.label_contador.pack()

        root.bind("<Map>", self.on_restore)
        root.bind("<Unmap>", self.on_minimize)

        root.after(500, root.iconify)

        self.atualizar_contador()

        tk.Label(root, text="Desenvolvido por: Marco Antonio Nitsche", font=("Arial", 8)).pack(pady=(10, 0))
        tk.Label(root, text="@manitsche", font=("Arial", 8, "italic")).pack()

    def atualizar_contador(self):
        self.label_contador.config(text=f"Linhas: {linha_atual}/{total_linhas}")
        self.root.after(1000, self.atualizar_contador)

    def on_restore(self, event):
        global pausado
        if self.root.state() == "normal":
            pausado = True
            atualizar_status("Sistema pausado")

    def on_minimize(self, event):
        global pausado
        if self.root.state() == "iconic":
            pausado = False
            atualizar_status("Rodando automático")

    def salvar(self):
        salvar_mensagem(self.texto.get("1.0", tk.END))
        messagebox.showinfo("OK", "Mensagem salva")

# ==============================
# MAIN
# ==============================

if __name__ == "__main__":
    root = tk.Tk()
    Interface(root)

    threading.Thread(target=loop_automatico, daemon=True).start()

    root.mainloop()