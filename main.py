"""
PROJETO: MediaConverter Pro
AUTOR: Alexandre G Paz
VERSÃO: 1.6
DATA: 11/07/2026
"""

import os
import sys
import shutil
import requests
import winshell
from win32com.client import Dispatch
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
import pillow_heif
import pillow_avif

pillow_heif.register_heif_opener()

# --- CONFIGURAÇÕES DE ATUALIZAÇÃO ---
# Você deve colocar o link de um arquivo .json ou .txt no seu GitHub/Drive aqui
UPDATE_URL = "https://raw.githubusercontent.com/seu-usuario/seu-repo/main/version.json"

class MediaConverterPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.version = "1.6"
        self.author = "Alexandre G Paz"
        self.install_path = os.path.join(os.environ["LOCALAPPDATA"], "MediaConverterPro")
        
        self.title(f"MediaConverter Pro - {self.version}")
        self.geometry("1100x850")

        # Verificar se o app já está instalado ou se é a primeira vez
        if not self.is_installed() and not "--portable" in sys.argv:
            self.show_welcome_wizard()
        else:
            self.check_for_updates()
            self.setup_main_ui()

    def is_installed(self):
        """Verifica se o executável está rodando da pasta de instalação"""
        return sys.executable.startswith(self.install_path)

    def show_welcome_wizard(self):
        """Tela inicial de escolha: Instalar ou Portátil"""
        self.wizard = ctk.CTkToplevel(self)
        self.wizard.title("Assistente de Configuração")
        self.wizard.geometry("500x400")
        self.wizard.grab_set() # Foca apenas nesta janela

        ctk.CTkLabel(self.wizard, text="Bem-vindo ao MediaConverter Pro", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        ctk.CTkLabel(self.wizard, text="Como deseja utilizar o programa?", font=ctk.CTkFont(size=14)).pack(pady=10)

        ctk.CTkButton(self.wizard, text="💻 INSTALAR NO COMPUTADOR\n(Cria atalhos e pasta fixa)", 
                      height=60, command=self.install_app).pack(pady=10, padx=50, fill="x")
        
        ctk.CTkButton(self.wizard, text="🚀 EXECUTAR MODO PORTÁTIL\n(Não instala nada)", 
                      fg_color="gray", height=60, command=self.run_portable).pack(pady=10, padx=50, fill="x")

    def install_app(self):
        """Copia o executável para o AppData e cria atalho na Área de Trabalho"""
        try:
            if not os.path.exists(self.install_path):
                os.makedirs(self.install_path)

            dest_exe = os.path.join(self.install_path, "MediaConverterPro.exe")
            shutil.copy2(sys.executable, dest_exe)

            # Criar Atalho
            desktop = winshell.desktop()
            path = os.path.join(desktop, "MediaConverter Pro.lnk")
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = dest_exe
            shortcut.WorkingDirectory = self.install_path
            shortcut.IconLocation = dest_exe
            shortcut.save()

            messagebox.showinfo("Sucesso", "Instalação concluída! Use o atalho na sua Área de Trabalho.")
            self.wizard.destroy()
            self.setup_main_ui()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na instalação: {e}")

    def run_portable(self):
        self.wizard.destroy()
        self.setup_main_ui()

    def check_for_updates(self):
        """Verifica se existe uma versão mais nova online"""
        try:
            # Simulando uma requisição (em produção, use requests.get(UPDATE_URL))
            # response = requests.get(UPDATE_URL, timeout=5)
            # online_version = response.json()['version']
            online_version = "1.6" # Exemplo

            if online_version > self.version:
                if messagebox.askyesno("Atualização Disponível", f"Uma nova versão ({online_version}) foi encontrada! Deseja baixar?"):
                    import webbrowser
                    webbrowser.open("https://seusite.com/download")
        except:
            pass # Silencioso se estiver sem internet

    def setup_main_ui(self):
        # [O RESTANTE DO SEU CÓDIGO DE INTERFACE v1.5 AQUI...]
        # (Para economizar espaço, mantenha a lógica de abas e conversão que já fizemos)
        self.render_interface()

    def render_interface(self):
        # Aqui entra todo o código da UI que já estruturamos (Sidebar, Main Frame, etc)
        # Lembre-se de manter o rodapé com seu nome Alexandre G Paz e a data 11/07/2026
        pass

if __name__ == "__main__":
    app = MediaConverterPro()
    app.mainloop()