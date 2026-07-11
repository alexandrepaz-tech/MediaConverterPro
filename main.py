"""
PROJETO: MediaConverter Pro
AUTOR: Alexandre G Paz
VERSÃO: 1.5
DATA DE LANÇAMENTO: 11/07/2026
DIREITOS AUTORAIS © 2026 - Todos os direitos reservados.
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

# --- CONFIGURAÇÕES ---
GITHUB_USER = "seu-usuario"
REPO_NAME = "MediaConverterPro"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/main/version.json"
INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "MediaConverterPro")

pillow_heif.register_heif_opener()

# --- FUNÇÕES DE SISTEMA (FORA DA CLASSE PARA ESTABILIDADE) ---

def install_app():
    """Realiza a instalação do executável no sistema"""
    try:
        current_exe = os.path.abspath(sys.executable)
        if not os.path.exists(INSTALL_DIR):
            os.makedirs(INSTALL_DIR)
        
        dest_exe = os.path.join(INSTALL_DIR, "MediaConverterPro.exe")
        shutil.copy2(current_exe, dest_exe)

        # Criar Atalho na Área de Trabalho
        desktop = winshell.desktop()
        path = os.path.join(desktop, "MediaConverter Pro.lnk")
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = dest_exe
        shortcut.WorkingDirectory = INSTALL_DIR
        shortcut.IconLocation = dest_exe
        shortcut.save()

        messagebox.showinfo("Sucesso", "Instalação concluída!\nUse o atalho na sua Área de Trabalho.")
        sys.exit()
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na instalação: {e}")
        sys.exit()

def check_startup_logic():
    """Decide se o app deve instalar, rodar portátil ou fechar"""
    if not getattr(sys, 'frozen', False):
        return # Se for código fonte, apenas segue

    current_dir = os.path.dirname(os.path.abspath(sys.executable)).lower()
    
    # Se não estiver na pasta de instalação, pergunta o que fazer
    if current_dir != INSTALL_DIR.lower():
        msg = "Deseja INSTALAR o MediaConverter Pro no computador?\n\nSim: Instalar e criar atalho.\nNão: Rodar como Portátil.\nCancelar: Sair."
        choice = messagebox.askyesnocancel("MediaConverter Pro - Início", msg)
        
        if choice is True:
            install_app()
        elif choice is None:
            sys.exit()
        # Se for False (Não), o código continua para abrir o app portátil

# --- CLASSE PRINCIPAL DA INTERFACE ---

class MediaConverterPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.version = "1.5"
        self.author = "Alexandre G Paz"
        self.release_date = "11/07/2026"

        self.title(f"MediaConverter Pro - {self.version}")
        self.geometry("1100x850")

        # Verificar atualizações em segundo plano
        self.check_for_updates()

        # Dicionário de Formatos
        self.format_library = {
            "PNG": "💎 [Lossless] Alta qualidade com transparência. Ideal para logos e web.",
            "JPEG": "📸 [Padrão] O mais comum para fotos. Ótima compressão.",
            "WEBP": "🌐 [Google] Formato moderno para internet. Muito leve.",
            "AVIF": "🚀 [Próxima Geração] Melhor compressão do mundo atual.",
            "HEIC": "🍏 [Apple] Padrão do iPhone. Alta qualidade, pouco espaço.",
            "ICO": "🖱️ [Ícone] Usado para ícones de programas e sites.",
            "BMP": "🖼️ [Bitmap] Formato bruto do Windows. Sem compressão.",
            "TIFF": "🖨️ [Gráfica] Alta fidelidade para impressão profissional.",
            "GIF": "🎞️ [Animável] Limitado a 256 cores. Ícones e memes.",
            "TGA": "🎮 [Games] Comum em texturas de jogos e 3D.",
            "DDS": "🕹️ [DirectDraw] Formato técnico para placas de vídeo.",
            "PCX": "💾 [Legacy] Formato clássico do Paintbrush antigo.",
            "SVG": "📐 [Vetor] Formato baseado em código (Raster-Wrap).",
            "PSD": "🎨 [Photoshop] Formato Adobe (Imagem achatada).",
            "HDR": "☀️ [High Dynamic Range] Iluminação 3D e fotos de alta luz.",
            "EXR": "🎬 [Cinema] Profissional para efeitos visuais.",
            "PPM/PGM": "🔬 [Científico] Processamento acadêmico de imagens.",
            "XBM/XWD": "🖥️ [Unix] Formatos antigos de sistemas Linux."
        }

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.logo = ctk.CTkLabel(self.sidebar, text="🚀 MEDIA PRO", font=ctk.CTkFont(size=24, weight="bold"))
        self.logo.pack(pady=30)

        self.btn_img = ctk.CTkButton(self.sidebar, text="🖼️ Imagens", command=self.setup_image_ui, anchor="w")
        self.btn_img.pack(pady=10, padx=20, fill="x")
        
        self.btn_vid = ctk.CTkButton(self.sidebar, text="🎥 Vídeos (Em Breve)", state="disabled", fg_color="gray25", anchor="w")
        self.btn_vid.pack(pady=10, padx=20, fill="x")

        self.copyright_label = ctk.CTkLabel(
            self.sidebar, 
            text=f"© 2026 {self.author}\nLançamento: {self.release_date}\nVersão {self.version}",
            font=ctk.CTkFont(size=11), text_color="gray"
        )
        self.copyright_label.pack(side="bottom", pady=20)

        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.setup_image_ui()

    def check_for_updates(self):
        try:
            response = requests.get(VERSION_URL, timeout=3)
            if response.status_code == 200:
                data = response.json()
                if data.get("version") > self.version:
                    if messagebox.askyesno("Atualização", "Nova versão disponível! Baixar?"):
                        import webbrowser
                        webbrowser.open(data.get("download_url"))
        except: pass

    def setup_image_ui(self):
        for widget in self.main_frame.winfo_children(): widget.destroy()
        ctk.CTkLabel(self.main_frame, text="Conversor Universal de Imagens", font=ctk.CTkFont(size=26, weight="bold")).pack(pady=20)
        self.mode_var = tk.IntVar(value=1)
        self.frame_mode = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.frame_mode.pack(pady=10)
        ctk.CTkRadioButton(self.frame_mode, text="📄 Arquivo Único", variable=self.mode_var, value=0).grid(row=0, column=0, padx=20)
        ctk.CTkRadioButton(self.frame_mode, text="📁 Pasta Inteira", variable=self.mode_var, value=1).grid(row=0, column=1, padx=20)
        self.btn_select = ctk.CTkButton(self.main_frame, text="🔍 SELECIONAR ORIGEM", command=self.select_source, height=45, font=ctk.CTkFont(weight="bold"))
        self.btn_select.pack(pady=15)
        self.lbl_path = ctk.CTkLabel(self.main_frame, text="Nenhum item selecionado", text_color="gray")
        self.lbl_path.pack()
        self.format_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
        self.format_frame.pack(pady=20, padx=40, fill="x")
        ctk.CTkLabel(self.format_frame, text="Escolha o Formato de Saída:", font=ctk.CTkFont(size=14)).pack(pady=(10,5))
        self.format_menu = ctk.CTkOptionMenu(self.format_frame, values=list(self.format_library.keys()), command=self.update_info_text)
        self.format_menu.pack(pady=10)
        self.info_label = ctk.CTkLabel(self.format_frame, text=self.format_library["PNG"], wraplength=500, text_color="#3498db", font=ctk.CTkFont(size=13, slant="italic"))
        self.info_label.pack(pady=15, padx=20)
        self.btn_run = ctk.CTkButton(self.main_frame, text="⚡ INICIAR CONVERSÃO E SUBSTITUIR", fg_color="#c0392b", hover_color="#a93226", font=ctk.CTkFont(size=18, weight="bold"), height=60, command=self.run_conversion)
        self.btn_run.pack(pady=20, padx=50, fill="x")
        self.progress = ctk.CTkProgressBar(self.main_frame)
        self.progress.set(0)
        self.progress.pack(pady=10, padx=50, fill="x")

    def update_info_text(self, choice): self.info_label.configure(text=self.format_library[choice])
    def select_source(self):
        path = filedialog.askopenfilename() if self.mode_var.get() == 0 else filedialog.askdirectory()
        if path:
            self.source_path = path
            self.lbl_path.configure(text=f"📍 {path}", text_color="white")

    def run_conversion(self):
        if not hasattr(self, 'source_path') or not self.source_path:
            messagebox.showwarning("Erro", "Selecione a origem primeiro!")
            return
        target_fmt = self.format_menu.get()
        files = [self.source_path] if os.path.isfile(self.source_path) else [os.path.join(self.source_path, f) for f in os.listdir(self.source_path) if os.path.isfile(os.path.join(self.source_path, f))]
        base_dir = os.path.dirname(self.source_path) if os.path.isfile(self.source_path) else self.source_path
        if not files: return
        backup_dir = os.path.join(base_dir, f"BACKUP_ORIGINAIS_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(backup_dir, exist_ok=True)
        for f in files: shutil.copy2(f, backup_dir)
        total = len(files)
        sucesso = 0
        for i, file_path in enumerate(files):
            try:
                img = Image.open(file_path)
                icc = img.info.get("icc_profile")
                final_path = os.path.join(os.path.dirname(file_path), f"{os.path.splitext(os.path.basename(file_path))[0]}.{target_fmt.lower()}")
                if target_fmt in ["JPEG", "BMP"] and img.mode in ("RGBA", "P"): img = img.convert("RGB")
                if target_fmt == "ICO": img.save(final_path, format="ICO", sizes=[(256, 256)])
                else: img.save(final_path, format=target_fmt, icc_profile=icc, quality=95)
                if os.path.exists(final_path) and os.path.abspath(file_path).lower() != os.path.abspath(final_path).lower(): os.remove(file_path)
                sucesso += 1
            except: pass
            self.progress.set((i + 1) / total)
            self.update_idletasks()
        messagebox.showinfo("Sucesso", f"✅ {sucesso} arquivos processados por {self.author}!\n\nBackup: {os.path.basename(backup_dir)}")
        self.progress.set(0)

# --- EXECUÇÃO ---

if __name__ == "__main__":
    # 1. Decide se instala ou roda portátil ANTES de iniciar a interface
    check_startup_logic()
    
    # 2. Se chegou aqui, inicia o app normalmente
    app = MediaConverterPro()
    app.mainloop()