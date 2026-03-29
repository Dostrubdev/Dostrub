import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import keyboard
import threading
import time
import win32api
import win32con
import win32gui
import sys
import os
import webbrowser
import logging
import subprocess
from PIL import Image, ImageTk
from constants import AZERTY_TO_SCAN, SCAN_TO_AZERTY

logger = logging.getLogger("Dostrub.GUI")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

GOLD = "#c8922a"
GOLD_LIGHT = "#e8b84b"
BG_VOID = "#0a0a0a"
BG_CARD = "#121212"
BG_BTN_DARK = "#1a1a1a"


ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("green") # Neutral base theme, but we override most colors


class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent_gui):
        super().__init__(parent_gui.root)
        self.parent = parent_gui
        self.app = parent_gui.app
        self.title("⚙️ Paramètres Dostrub")
        
        self.geometry("520x700")
        self.minsize(350, 300)
        self.attributes("-topmost", True)
        self.resizable(True, True)


        self.scroll_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_container.pack(fill="both", expand=True, padx=5, pady=5)

        title_font = ctk.CTkFont(size=16, weight="bold")
        
        ctk.CTkLabel(self.scroll_container, text="Raccourcis Jeu & Macros", font=title_font, text_color=GOLD).pack(pady=(15, 5))
        frame_game = ctk.CTkFrame(self.scroll_container, fg_color=BG_CARD)
        frame_game.pack(fill="x", padx=10, pady=5)
        
      
        self.parent.create_hotkey_row(frame_game, "Inventaire", "game_inv_key", 0, 0, "Touche jeu pour ouvrir l'inventaire")
        self.parent.create_hotkey_row(frame_game, "Carac", "game_char_key", 0, 3, "Touche jeu pour ouvrir les caractéristiques")
        
        
        self.parent.create_hotkey_row(frame_game, "Sorts", "game_spell_key", 1, 0, "Touche jeu pour ouvrir les sorts")
        self.parent.create_hotkey_row(frame_game, "Havre-Sac", "game_haven_key", 1, 3, "Touche jeu pour ouvrir le Havre-Sac")

    
        separateur = ctk.CTkFrame(frame_game, height=2, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD)
        separateur.grid(row=2, column=0, columnspan=6, pady=15, sticky="ew", padx=10)

   
        self.parent.create_hotkey_row(frame_game, "Inviter Groupe", "invite_group_key", 3, 0, "Inviter automatiquement le groupe")
        self.parent.create_hotkey_row(frame_game, "Auto-Zaap", "auto_zaap_key", 3, 3, "Lancer la macro Auto-Zaap")

   
        self.parent.create_hotkey_row(frame_game, "Ctrl+V/Entrée", "paste_enter_key", 4, 0, "Coller et Entrée sur tout le monde (Ctrl+V)")
        self.parent.create_hotkey_row(frame_game, "Actualiser", "refresh_key", 4, 3, "Actualisation des pages")

  
        self.parent.create_hotkey_row(frame_game, "Trier Barre", "sort_taskbar_key", 5, 0, "Trie la barre des tâches Windows")
        self.parent.create_hotkey_row(frame_game, "Calibrage (Bind)", "calib_key", 5, 3, "Touche de bind calibrage")

        ctk.CTkLabel(self.scroll_container, text="Roue de Focus (Radiale)", font=title_font, text_color=GOLD).pack(pady=(20, 5))
        frame_radial = ctk.CTkFrame(self.scroll_container, fg_color=BG_CARD)
        frame_radial.pack(fill="x", padx=10, pady=5)
        
        self.var_radial = ctk.BooleanVar(value=self.app.config.data.get("radial_menu_active", True))
        sw_radial = ctk.CTkSwitch(frame_radial, text="Activer la roue", variable=self.var_radial, command=self.save_settings, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_radial.pack(pady=10)
        self.parent.bind_tooltip(sw_radial, "Activer ou désactiver complètement la roue de sélection")

        frame_radial_hk = ctk.CTkFrame(frame_radial, fg_color="transparent")
        frame_radial_hk.pack(pady=(0, 10))

        lbl_hk = ctk.CTkLabel(frame_radial_hk, text="Raccourci :", text_color=GOLD)
        lbl_hk.pack(side="left", padx=(0, 5))
        
        current_val = self.app.config.data.get("radial_menu_hotkey", "alt+left_click")
        btn_hk = ctk.CTkButton(frame_radial_hk, text=current_val if current_val else "Aucun", width=120, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.parent.catch_key("radial_menu_hotkey", btn_hk, allow_mouse=True))
        btn_hk.pack(side="left", padx=5)
        
        self.parent.hotkey_btns["radial_menu_hotkey"] = btn_hk

        btn_x = ctk.CTkButton(frame_radial_hk, text="✖", width=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.parent.clear_key("radial_menu_hotkey", btn_hk))
        btn_x.pack(side="left", padx=5)
        
       
        ctk.CTkLabel(self.scroll_container, text="Vitesse du Clic Multi", font=title_font, text_color=GOLD).pack(pady=(15, 5))
        frame_speed = ctk.CTkFrame(self.scroll_container, fg_color=BG_CARD)
        frame_speed.pack(fill="x", padx=10, pady=5)
        
        lbl_speed = ctk.CTkLabel(frame_speed, text="Vitesse :", text_color=GOLD)
        lbl_speed.pack(side="left", padx=10, pady=10)
        
        self.combo_speed = ctk.CTkOptionMenu(frame_speed, values=["Rapide", "Moyen", "Lent"], command=self.save_speed, fg_color=BG_BTN_DARK, text_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT)
        self.combo_speed.set(self.app.config.data.get("click_speed", "Lent"))
        self.combo_speed.pack(side="left", padx=10, pady=10)
        self.parent.bind_tooltip(self.combo_speed, "Ajuste la vitesse d'exécution des clics synchronisés (Gauche et Droit)")
        
        ctk.CTkLabel(self.scroll_container, text="Macro Auto-Zaap", font=title_font, text_color=GOLD).pack(pady=(15, 5))
        frame_zaap = ctk.CTkFrame(self.scroll_container, fg_color=BG_CARD)
        frame_zaap.pack(fill="x", padx=10, pady=5)

        lbl_zaap = ctk.CTkLabel(frame_zaap, text="Délai avant clic (sec) :", text_color=GOLD)
        lbl_zaap.pack(side="left", padx=10, pady=10)

        self.entry_zaap_delay = ctk.CTkEntry(frame_zaap, width=60, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD)
        self.entry_zaap_delay.insert(0, str(self.app.config.data.get("zaap_delay", "1.8")))
        self.entry_zaap_delay.pack(side="left", padx=10, pady=10)
        self.parent.bind_tooltip(self.entry_zaap_delay, "Temps d'attente pour l'ouverture du Havre-Sac (ex: 1.0 ou 1.8)")

        btn_reset = ctk.CTkButton(self.scroll_container, text="Reset Settings", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.parent.reset_all)
        btn_reset.pack(pady=(10, 0))
        self.parent.bind_tooltip(btn_reset, "Remise à zéro de tous les paramètres et raccourcis")

        btn_close = ctk.CTkButton(self.scroll_container, text="Fermer & Sauvegarder", fg_color=GOLD, hover_color=GOLD_LIGHT, text_color=BG_VOID, command=self.close_and_save)
        btn_close.pack(pady=(20, 10))

    def save_settings(self):
        self.app.config.data["radial_menu_active"] = self.var_radial.get()
        self.app.config.save()
        
    def close_and_save(self):
        self.app.config.data["zaap_delay"] = self.entry_zaap_delay.get()
        self.app.config.save()
        self.destroy()

    def save_speed(self, choice):
        self.app.config.data["click_speed"] = choice
        self.app.config.save()

class FloatingToolbar:
    def __init__(self, app_controller, parent_gui):
        self.app = app_controller
        self.parent_gui = parent_gui
        self.window = tk.Toplevel(parent_gui.root)
        self.window.overrideredirect(True) 
        self.window.attributes("-topmost", True)
        self.window.attributes("-alpha", 0.90) 
        self.window.configure(bg=BG_CARD)
        
        self.x = 0; self.y = 0
        saved_x = self.app.config.data.get("toolbar_x", 100)
        saved_y = self.app.config.data.get("toolbar_y", 100)
        self.window.geometry(f"+{saved_x}+{saved_y}")
        
        title_frame = tk.Frame(self.window, bg=BG_VOID, cursor="fleur")
        title_frame.pack(fill="x")
        
        path_overlay = resource_path("skin/overlay_logo.png")
        if os.path.exists(path_overlay):
            img_logo = Image.open(path_overlay)
            w, h = img_logo.size
            new_h = 26
            new_w = int(w * (new_h / h))
            self.img_logo_tk = ctk.CTkImage(img_logo, size=(new_w, new_h))
            tk_lbl = ctk.CTkLabel(title_frame, image=self.img_logo_tk, text="", cursor="fleur", fg_color=BG_VOID)
        else:
            tk_lbl = tk.Label(title_frame, text="≡ Dostrub ≡", bg=BG_VOID, fg=GOLD, font=("Arial", 8, "bold"))

        tk_lbl.pack(pady=2)
        tk_lbl.bind("<Button-1>", self.start_move)
        tk_lbl.bind("<B1-Motion>", self.do_move)
        tk_lbl.bind("<ButtonRelease-1>", self.stop_move) 

        content = tk.Frame(self.window, bg=BG_CARD)
        content.pack(padx=5, pady=5)
        
        top_row = tk.Frame(content, bg=BG_CARD)
        top_row.pack(fill="x", pady=(0, 5))
        
        self.btn_show_gui = ctk.CTkButton(top_row, text="⬅", width=25, height=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.parent_gui.show_gui)
        self.btn_show_gui.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.btn_show_gui, "Ouvrir l'interface principale")
        
        self.combo_mode = ctk.CTkOptionMenu(top_row, values=["ALL", "Team 1", "Team 2"], width=80, height=25, command=self.on_mode_change, fg_color=BG_BTN_DARK, text_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT)
        self.combo_mode.set(self.app.config.data.get("current_mode", "ALL"))
        self.combo_mode.pack(side="left", padx=2)
        


        img_grp = self.load_icon("groupe.png")
        self.btn_auto_group = ctk.CTkButton(top_row, text="" if img_grp else "G", image=img_grp, width=25, height=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: threading.Thread(target=self.app.logic.execute_group_invite, daemon=True).start())
        self.parent_gui.bind_tooltip(self.btn_auto_group, "Inviter le groupe")
        
        self.update_overlay_icons()
        
        img_dup = self.load_icon("dupliquer.png")
        self.btn_paste = ctk.CTkButton(top_row, text="📋" if not img_dup else "", image=img_dup, width=25, height=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: threading.Thread(target=self.app.logic.execute_paste_enter, daemon=True).start())
        self.btn_paste.pack(side="right", padx=2)
        self.parent_gui.bind_tooltip(self.btn_paste, "Coller + Entrée sur toutes les pages (Ctrl+V)")

        self.btn_refresh_overlay = ctk.CTkButton(top_row, text="F5", font=ctk.CTkFont(size=11, weight="bold"), width=25, height=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.parent_gui.original_app_refresh)
        self.btn_refresh_overlay.pack(side="right", padx=2)
        self.parent_gui.bind_tooltip(self.btn_refresh_overlay, "Rafraîchir les pages Dofus")

        self.bot_row = tk.Frame(content, bg=BG_CARD)
        self.bot_row.pack(fill="x")

        img_inv = self.load_icon("inventaire.png")  
        img_carac = self.load_icon("carac.png")    
        img_sort = self.load_icon("sort.png")      
        img_havre = self.load_icon("havresac.png")
        img_zaap = self.load_icon("zaap.png")

        self.b1 = ctk.CTkButton(self.bot_row, text="I" if not img_inv else "", image=img_inv, width=28, height=28, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.bcast("game_inv_key", "i"))
        self.b1.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.b1, "Ouvrir Inventaire")
        
        self.b2 = ctk.CTkButton(self.bot_row, text="C" if not img_carac else "", image=img_carac, width=28, height=28, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.bcast("game_char_key", "c"))
        self.b2.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.b2, "Ouvrir Caractéristiques")
        
        self.b3 = ctk.CTkButton(self.bot_row, text="S" if not img_sort else "", image=img_sort, width=28, height=28, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.bcast("game_spell_key", "s"))
        self.b3.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.b3, "Ouvrir Sorts")
        
        self.b4 = ctk.CTkButton(self.bot_row, text="H" if not img_havre else "", image=img_havre, width=28, height=28, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.bcast("game_haven_key", "h"))
        self.b4.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.b4, "Ouvrir Havre-Sac")
        
        self.b5 = ctk.CTkButton(self.bot_row, text="Z" if not img_zaap else "", image=img_zaap, width=28, height=28, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: threading.Thread(target=self.app.logic.execute_auto_zaap, daemon=True).start())
        self.b5.pack(side="left", padx=2)
        self.parent_gui.bind_tooltip(self.b5, "Auto-Zaap (Ouvre havre-sac et clique le Zaap)")

        self.window.withdraw()

    def bcast(self, config_key, default):
        k = self.app.config.data.get(config_key, default)
        threading.Thread(target=self.app.logic.broadcast_key, args=(k,), daemon=True).start()

    def load_icon(self, filename):
        path = resource_path(f"skin/{filename}")
        if os.path.exists(path):
            return ctk.CTkImage(light_image=Image.open(path).resize((20, 20), Image.Resampling.LANCZOS), dark_image=Image.open(path).resize((20, 20), Image.Resampling.LANCZOS), size=(20, 20))
        return None

    def start_move(self, event):
        self.x = event.x
        self.y = event.y
    def do_move(self, event):
        self.window.geometry(f"+{self.window.winfo_x() + (event.x - self.x)}+{self.window.winfo_y() + (event.y - self.y)}")
    def stop_move(self, event):
        self.app.config.data["toolbar_x"], self.app.config.data["toolbar_y"] = self.window.winfo_x(), self.window.winfo_y()
        self.app.config.save()
    def on_mode_change(self, choice):
        self.app.logic.set_mode(choice)
        self.app.current_idx = 0
        self.parent_gui.combo_mode.set(choice)
    def show(self): self.window.deiconify()
    def hide(self): self.window.withdraw()

    def update_overlay_icons(self):
        cfg = self.app.config.data
        if cfg.get("auto_group_enabled", False):
            self.btn_auto_group.pack(side="right", padx=2)
        else:
            self.btn_auto_group.pack_forget()

class OrganizerGUI:
    def __init__(self, app_controller):
        self.app = app_controller
        self.root = ctk.CTk()
        self.root.title(f"Dostrub v{self.app.version}")
        self.root.configure(fg_color=BG_VOID)
        
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w = 620
        win_h = min(screen_h - 80, 800)
        x_pos = int((screen_w / 2) - (win_w / 2))
        y_pos = int((screen_h / 2) - (win_h / 2))
        self.root.geometry(f"{win_w}x{win_h}+{x_pos}+{y_pos}") 
        
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.original_app_refresh = self.app.refresh
        
        if os.path.exists(resource_path("Dostrub.ico")):
            try:
                self.root.iconbitmap(resource_path("Dostrub.ico"))
            except Exception:
                pass
                
        self.is_listening = False
        self.is_visible = True 
        cfg = self.app.config.data 
        
        self.toolbar = FloatingToolbar(self.app, self)
        
        self.var_tooltips = ctk.BooleanVar(value=cfg.get("show_tooltips", True))
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.overrideredirect(True)
        self.tooltip.attributes("-topmost", True)
        self.tooltip_lbl = tk.Label(self.tooltip, text="", fg="#ecf0f1", bg=BG_CARD, font=("Segoe UI", 9), padx=6, pady=3, relief="solid", borderwidth=1)
        self.tooltip_lbl.pack()
        self.tooltip.withdraw() 
        self.hotkey_btns = {} 
        self.checkbox_vars = {} 

        self.header_f = ctk.CTkFrame(self.root, fg_color="transparent")
        self.header_f.pack(fill="x", padx=15, pady=(15, 5))
        
        path_head = resource_path("dostrubhead.png")
        if os.path.exists(path_head):
            img = Image.open(path_head)
            w, h = img.size
            new_w = int(w * (45.0/h))
            head_img = ctk.CTkImage(img, size=(new_w, 45))
            ctk.CTkLabel(self.header_f, image=head_img, text="").pack(side="left")
        else:
            ctk.CTkLabel(self.header_f, text=f"DOSTRUB v{self.app.version}", font=ctk.CTkFont(size=20, weight="bold"), text_color=GOLD).pack(side="left")
        
        self.btn_kill = ctk.CTkButton(self.header_f, text="⚪ OFF", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, width=60, command=self.hard_kill_app)
        self.btn_kill.pack(side="right", padx=(0, 10))
        self.bind_tooltip(self.btn_kill, "Shutdown complet : Force la fermeture immédiate de Dostrub")

        sep = ctk.CTkFrame(self.root, height=2, fg_color=GOLD)
        sep.pack(fill="x", padx=15, pady=(0, 5))

        # ------------------- TABVIEW CONTAINER -------------------
        self.tabview = ctk.CTkTabview(self.root, corner_radius=10, 
                                      fg_color=BG_CARD, 
                                      segmented_button_fg_color=BG_VOID, 
                                      segmented_button_selected_color=BG_CARD, # Use card color for selected tab to keep Gold text visible
                                      segmented_button_selected_hover_color=GOLD_LIGHT, 
                                      segmented_button_unselected_color=BG_VOID, 
                                      segmented_button_unselected_hover_color=BG_BTN_DARK)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=5)
        self.tab_comptes = self.tabview.add("👥 Comptes")
        self.tabview.tab("👥 Comptes").configure(fg_color=BG_VOID)
        self.tab_macros = self.tabview.add("⚡ Binds & Options")
        self.tabview.tab("⚡ Binds & Options").configure(fg_color=BG_VOID)
        self.tab_calib = self.tabview.add("⚙️ Calibrages")
        self.tabview.tab("⚙️ Calibrages").configure(fg_color=BG_VOID)
        self.tab_settings = self.tabview.add("🛠️ Paramètres")
        self.tabview.tab("🛠️ Paramètres").configure(fg_color=BG_VOID)
        
        # Coloration des boutons de segmentation (onglets)
        self.tabview._segmented_button.configure(text_color=GOLD)

        # ------------------- TAB: COMPTES -------------------
        self.frame_mode = ctk.CTkFrame(self.tab_comptes, fg_color="transparent")
        self.frame_mode.pack(fill="x", padx=10, pady=(10, 5))
        lbl_ctrl = ctk.CTkLabel(self.frame_mode, text="Contrôler :", font=ctk.CTkFont(weight="bold"), text_color=GOLD)
        lbl_ctrl.pack(side="left", padx=5)
        self.combo_mode = ctk.CTkOptionMenu(self.frame_mode, values=["ALL", "Team 1", "Team 2"], command=self.on_mode_change, fg_color=BG_BTN_DARK, text_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT)
        self.combo_mode.set(cfg.get("current_mode", "ALL"))
        self.combo_mode.pack(side="left", padx=5)
        self.bind_tooltip(self.combo_mode, "Cible du focus (Tout le monde ou équipe spécifique)")

        pill_frame = ctk.CTkFrame(self.tab_comptes, fg_color=BG_CARD, corner_radius=8)
        pill_frame.pack(fill="x", padx=10, pady=(5, 5), ipadx=5, ipady=2)
        lbl_accounts = ctk.CTkLabel(pill_frame, text="Comptes actifs", font=ctk.CTkFont(size=13, weight="bold"), text_color=GOLD)
        lbl_accounts.pack(side="left", padx=10, pady=4)

        self.scroll_frame = ctk.CTkScrollableFrame(self.tab_comptes, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.action_frame = ctk.CTkFrame(self.tab_comptes, fg_color="transparent")
        self.action_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        row_act = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        row_act.pack(fill="x", pady=5)
        self.btn_invite = ctk.CTkButton(row_act, text="Inviter Groupe", fg_color=GOLD, hover_color=GOLD_LIGHT, text_color=BG_VOID, command=self.app.logic.execute_group_invite)
        self.btn_invite.pack(side="left", padx=15, expand=True)
        self.btn_close_all = ctk.CTkButton(row_act, text="Fermer Team", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.close_all_and_refresh)
        self.btn_close_all.pack(side="right", padx=15, expand=True)

        frame_audio = ctk.CTkFrame(self.tab_comptes, fg_color="transparent")
        frame_audio.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(frame_audio, text="🔊 Volume Roulette :", font=ctk.CTkFont(weight="bold"), text_color=GOLD).pack(side="left", padx=15, pady=10)
        self.slider_volume = ctk.CTkSlider(frame_audio, from_=0, to=100, command=self.on_volume_change, width=200, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT)
        self.slider_volume.set(cfg.get("volume_level", 50))
        self.slider_volume.pack(side="right", padx=15, pady=10)

        # ------------------- TAB: MACROS & OPTIONS -------------------
        top_container = ctk.CTkFrame(self.tab_macros, fg_color="transparent")
        top_container.pack(fill="both", expand=True, padx=5, pady=5)

        self.frame_keys = ctk.CTkFrame(top_container, fg_color=BG_CARD, corner_radius=8)
        self.frame_keys.pack(side="left", fill="both", expand=True, padx=(0, 5))
        ctk.CTkLabel(self.frame_keys, text="Raccourcis Clavier", font=ctk.CTkFont(weight="bold"), text_color=GOLD).grid(row=0, column=0, columnspan=6, pady=(10, 5))

        self.create_hotkey_row(self.frame_keys, "Précédent", "prev_key", 1, 0, "Focus perso précédent")
        self.create_hotkey_row(self.frame_keys, "Suivant", "next_key", 1, 3, "Focus perso suivant")
        self.create_hotkey_row(self.frame_keys, "Chef", "leader_key", 2, 0, "Reprendre focus sur le Chef")
        self.create_hotkey_row(self.frame_keys, "Clic Multi G.", "sync_key", 2, 3, "Clic gauche synchronisé")
        self.create_hotkey_row(self.frame_keys, "Chasse", "treasure_key", 3, 0, "Copier/Coller auto")
        self.create_hotkey_row(self.frame_keys, "Swap XP", "swap_xp_drop_key", 4, 0, "Clic synchro (XP/Drop)")
        self.create_hotkey_row(self.frame_keys, "Clic Multi D.", "sync_right_key", 3, 3, "Clic droit synchronisé")
        self.create_hotkey_row(self.frame_keys, "Afficher UI", "toggle_app_key", 4, 3, "Masquer/Afficher l'app")

        self.frame_switches = ctk.CTkFrame(top_container, fg_color=BG_CARD, corner_radius=8, width=180)
        self.frame_switches.pack(side="right", fill="y", expand=False, padx=(5, 0))
        self.frame_switches.pack_propagate(False)
        ctk.CTkLabel(self.frame_switches, text="Options Globales", font=ctk.CTkFont(weight="bold"), text_color=GOLD).pack(pady=(12, 8))

        self.var_toolbar = ctk.BooleanVar(value=cfg.get("toolbar_active", False))
        self.sw_overlay = ctk.CTkSwitch(self.frame_switches, text="Overlay", variable=self.var_toolbar, command=self.toggle_toolbar, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        self.sw_overlay.pack(anchor="w", padx=15, pady=6)
        self.bind_tooltip(self.sw_overlay, "Afficher la barre de macros flottante en jeu")

        self.var_return = ctk.BooleanVar(value=cfg.get("return_to_leader", True))
        sw_return = ctk.CTkSwitch(self.frame_switches, text="Focus Chef", variable=self.var_return, command=self.toggle_return, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_return.pack(anchor="w", padx=15, pady=6)
        self.bind_tooltip(sw_return, "Retour auto au chef apres action")

        self.var_spam = ctk.BooleanVar(value=cfg.get("spam_click_active", False))
        sw_spam = ctk.CTkSwitch(self.frame_switches, text="Spam Clic", variable=self.var_spam, command=self.toggle_macros, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_spam.pack(anchor="w", padx=15, pady=6)
        self.bind_tooltip(sw_spam, "Maintenez clic Molette pour spam clic")

        self.var_auto_group = ctk.BooleanVar(value=cfg.get("auto_group_enabled", False))
        sw_agrp = ctk.CTkSwitch(self.frame_switches, text="Auto Groupe", variable=self.var_auto_group, command=self.toggle_auto_group_opt, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_agrp.pack(anchor="w", padx=15, pady=6)
        self.bind_tooltip(sw_agrp, "Afficher le bouton Auto Groupe sur l'overlay")

        self.var_auto_trade = ctk.BooleanVar(value=cfg.get("auto_trade_enabled", False))
        sw_atrd = ctk.CTkSwitch(self.frame_switches, text="Auto Échange (Scanner)", variable=self.var_auto_trade, command=self.toggle_auto_trade_opt, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_atrd.pack(anchor="w", padx=15, pady=6)
        self.bind_tooltip(
            sw_atrd,
            "Vert → même enchaînement que ton raccourci Suivant (next_char), cherche le Valider sur les autres persos, "
            "clic puis Chef. Recalibre VALIDER (popup ouverte) pour la couleur. Si raté : settings.json "
            'auto_trade_step_delay (ex. 0.35) ou auto_trade_color_tolerance (ex. 55).',
        )

        def is_calibrated(key):
            val = self.app.config.data.get("macro_positions", {}).get(key)
            if not val:
                return False
            if key == "game_zone" and val == [0.0, 0.0, 0.0, 0.0]:
                return False
            if isinstance(val, dict) and not val:
                return False
            if key == "trade_validate_click_pos" and len(val) < 2:
                return False
            return True

        def reset_calib(key):
            self.app.config.data.setdefault("macro_positions", {}).pop(key, None)
            self.app.config.save()
            self.show_temporary_message(f"Calibrage '{key}' reinitialise.", GOLD)
            self.root.after(0, self.populate_calibrations)

        def calib_btn(parent, key, label, cmd):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill="x", padx=8, pady=3)
            row.columnconfigure(0, weight=1)
            btn_main = ctk.CTkButton(row, text=label, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD,
                                     command=cmd, anchor="w")
            btn_main.pack(side="left", fill="x", expand=True)
            if is_calibrated(key):
                btn_reset = ctk.CTkButton(row, text="Reset", width=52, height=28,
                                          fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, hover_color=BG_CARD,
                                          text_color=GOLD, font=ctk.CTkFont(size=11),
                                          command=lambda k=key: reset_calib(k))
                btn_reset.pack(side="right", padx=(6, 0))

        self.calib_container = ctk.CTkScrollableFrame(self.tab_calib, fg_color="transparent")
        self.calib_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        self._is_calibrated = is_calibrated
        self._calib_btn = calib_btn
        self.skin_cache = {}
        self._build_settings_tab_and_footer()
        self.populate_calibrations()

        # ------------------- TAB: METIER -------------------

        if self.var_toolbar.get():
            self.toolbar.show()

    def _build_settings_tab_and_footer(self):
        cfg = self.app.config.data
        stt = ctk.CTkScrollableFrame(self.tab_settings, fg_color="transparent")
        stt.pack(fill="both", expand=True, padx=5, pady=5)
        bold14 = ctk.CTkFont(size=14, weight="bold")

        def section_label(parent, text):
            ctk.CTkLabel(parent, text=text, font=bold14, text_color=GOLD, anchor="w").pack(fill="x", padx=10, pady=(10, 3))

        def section_card(parent):
            f = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=8)
            f.pack(fill="x", padx=8, pady=(0, 8))
            return f

        section_label(stt, "🎮 Raccourcis Jeu")
        card_game = section_card(stt)
        self.create_hotkey_row(card_game, "Inventaire", "game_inv_key", 0, 0, "Touche jeu pour ouvrir l'inventaire")
        self.create_hotkey_row(card_game, "Carac", "game_char_key", 0, 3, "Touche jeu pour ouvrir les caractéristiques")
        self.create_hotkey_row(card_game, "Sorts", "game_spell_key", 1, 0, "Touche jeu pour ouvrir les sorts")
        self.create_hotkey_row(card_game, "Havre-Sac", "game_haven_key", 1, 3, "Touche jeu pour ouvrir le Havre-Sac")
        self.create_hotkey_row(card_game, "Inviter Groupe", "invite_group_key", 3, 0, "Inviter automatiquement le groupe")
        self.create_hotkey_row(card_game, "Auto-Zaap", "auto_zaap_key", 3, 3, "Lancer la macro Auto-Zaap")
        self.create_hotkey_row(card_game, "Ctrl+V/Entrée", "paste_enter_key", 4, 0, "Coller et Entrée sur tout le monde")
        self.create_hotkey_row(card_game, "Actualiser", "refresh_key", 4, 3, "Actualisation des pages")
        self.create_hotkey_row(card_game, "Trier Barre", "sort_taskbar_key", 5, 0, "Trie la barre des tâches Windows")
        self.create_hotkey_row(card_game, "Calibrage (Bind)", "calib_key", 5, 3, "Touche de bind calibrage")

        section_label(stt, "🎡 Roue de Focus (Radiale)")
        card_radial = section_card(stt)
        self.var_radial = ctk.BooleanVar(value=cfg.get("radial_menu_active", True))

        def _save_radial():
            self.app.config.data["radial_menu_active"] = self.var_radial.get()
            self.app.config.save()

        _rv = cfg.get("radial_menu_hotkey", "alt+left_click")
        radial_row = ctk.CTkFrame(card_radial, fg_color="transparent")
        radial_row.pack(fill="x", padx=10, pady=10)
        sw_radial = ctk.CTkSwitch(radial_row, text="Activer", variable=self.var_radial, command=_save_radial, progress_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, text_color=GOLD)
        sw_radial.pack(side="left", padx=(5, 0))
        self.bind_tooltip(sw_radial, "Activer ou desactiver la roue de selection")
        right_hk = ctk.CTkFrame(radial_row, fg_color="transparent")
        right_hk.pack(side="right")
        ctk.CTkLabel(right_hk, text="Raccourci :", text_color=GOLD).pack(side="left", padx=(0, 5))
        btn_radial_hk = ctk.CTkButton(right_hk, text=_rv or "Aucun", width=120, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.catch_key("radial_menu_hotkey", btn_radial_hk, allow_mouse=True))
        btn_radial_hk.pack(side="left", padx=5)
        self.hotkey_btns["radial_menu_hotkey"] = btn_radial_hk
        ctk.CTkButton(right_hk, text="X", width=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.clear_key("radial_menu_hotkey", btn_radial_hk)).pack(side="left", padx=3)

        section_label(stt, "🖱️ Vitesse du Clic Multi")
        card_speed = section_card(stt)
        speed_row = ctk.CTkFrame(card_speed, fg_color="transparent")
        speed_row.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(speed_row, text="Vitesse :", text_color=GOLD).pack(side="left", padx=5)

        def _save_speed(c):
            self.app.config.data["click_speed"] = c
            self.app.config.save()

        self.combo_speed_settings = ctk.CTkOptionMenu(speed_row, values=["Rapide", "Moyen", "Lent"], command=_save_speed, fg_color=BG_BTN_DARK, text_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT)
        self.combo_speed_settings.set(cfg.get("click_speed", "Lent"))
        self.combo_speed_settings.pack(side="left", padx=10)
        self.bind_tooltip(self.combo_speed_settings, "Ajuste la vitesse des clics synchronisés")

        section_label(stt, "⚡ Macro Auto-Zaap")
        card_zaap = section_card(stt)
        zaap_row = ctk.CTkFrame(card_zaap, fg_color="transparent")
        zaap_row.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(zaap_row, text="Délai avant clic (sec) :", text_color=GOLD).pack(side="left", padx=5)
        self.entry_zaap_delay = ctk.CTkEntry(zaap_row, width=70, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD)
        self.entry_zaap_delay.insert(0, str(cfg.get("zaap_delay", "1.8")))
        self.entry_zaap_delay.pack(side="left", padx=10)
        self.bind_tooltip(self.entry_zaap_delay, "Temps d'attente pour l'ouverture du Havre-Sac")

        def _save_zaap_delay():
            self.app.config.data["zaap_delay"] = self.entry_zaap_delay.get()
            self.app.config.save()
            self.show_temporary_message("✅ Délai Zaap sauvegardé !", "#4CAF50")

        ctk.CTkButton(zaap_row, text="Sauver", width=70, fg_color=GOLD, hover_color=GOLD_LIGHT, text_color=BG_VOID, command=_save_zaap_delay).pack(side="left", padx=5)

        ctk.CTkButton(stt, text="🔄 Reset Tous les Paramètres", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.reset_all).pack(pady=(8, 4), padx=10, fill="x")
        ctk.CTkButton(stt, text="💾 Sauvegarder", fg_color=GOLD, hover_color=GOLD_LIGHT, text_color=BG_VOID, command=_save_zaap_delay).pack(pady=(0, 12), padx=10, fill="x")

        self.frame_footer = ctk.CTkFrame(self.root, fg_color="transparent")
        self.frame_footer.pack(side="bottom", fill="x", padx=15, pady=(5, 15))

        self.btn_refresh = ctk.CTkButton(self.frame_footer, text="Rafraîchir", fg_color=GOLD, hover_color=GOLD_LIGHT, text_color=BG_VOID, command=self.original_app_refresh, width=80)
        self.btn_refresh.pack(side="left", padx=(0, 5))
        self.btn_sort_win = ctk.CTkButton(self.frame_footer, text="Trier Barre Win", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=self.trigger_sort_taskbar, width=110)
        self.btn_sort_win.pack(side="left", padx=5)

        self.chk_tooltips = ctk.CTkCheckBox(self.frame_footer, text="Bulles", variable=self.var_tooltips, command=self.toggle_tooltips_setting, width=60, checkmark_color="#FFFFFF", fg_color=GOLD, text_color=GOLD)
        self.chk_tooltips.pack(side="left", padx=10)

        self.btn_hide = ctk.CTkButton(self.frame_footer, text="Cacher l'UI", command=self.toggle_visibility, fg_color="transparent", text_color=GOLD, border_width=1, border_color=GOLD, hover_color="#3B3B4F", width=70)
        self.btn_hide.pack(side="right")

        frame_msg = ctk.CTkFrame(self.root, fg_color="transparent", height=20)
        frame_msg.pack(side="bottom", fill="x", padx=15, pady=0)
        self.lbl_feedback = ctk.CTkLabel(frame_msg, text="", font=ctk.CTkFont(size=13, weight="bold"), text_color=GOLD)
        self.lbl_feedback.pack(expand=True)

    def populate_calibrations(self):
        for widget in list(self.calib_container.winfo_children()):
            widget.destroy()
        self.calib_container.update_idletasks()

        # --- Section: Interface Jeu ---
        ctk.CTkLabel(self.calib_container, text="Interface Jeu", font=ctk.CTkFont(weight="bold"), text_color=GOLD, anchor="w").pack(fill="x", padx=10, pady=(8, 2))
        card1 = ctk.CTkFrame(self.calib_container, fg_color=BG_CARD, corner_radius=8)
        card1.pack(fill="x", padx=8, pady=(0, 10))
        self._calib_btn(card1, "chat_position", "Calibrer Chat", self.start_calib_chat)
        self._calib_btn(card1, "xp_drop_button", "Calibrer XP / Drop", self.start_calib_xp_drop)

        # --- Section: Navigation ---
        ctk.CTkLabel(self.calib_container, text="Navigation", font=ctk.CTkFont(weight="bold"), text_color=GOLD, anchor="w").pack(fill="x", padx=10, pady=(4, 2))
        card2 = ctk.CTkFrame(self.calib_container, fg_color=BG_CARD, corner_radius=8)
        card2.pack(fill="x", padx=8, pady=(0, 10))
        self._calib_btn(card2, "game_zone", "Calibrer Zone de Jeu", self.start_calib_zone_jeu)
        self._calib_btn(card2, "map_borders", "Calibrer Bords de Map", self.start_calib_map_borders)
        self._calib_btn(card2, "map_coords", "Calibrer Coordonnees", self.start_calib_coord)

        # --- Section: Automatisation ---
        ctk.CTkLabel(self.calib_container, text="Automatisation (Farm)", font=ctk.CTkFont(weight="bold"), text_color=GOLD, anchor="w").pack(fill="x", padx=10, pady=(4, 2))
        card3 = ctk.CTkFrame(self.calib_container, fg_color=BG_CARD, corner_radius=8)
        card3.pack(fill="x", padx=8, pady=(0, 10))
        self._calib_btn(card3, "zaaps", "Calibrer Havre-Sac (Zaap)", self.start_calib_zaap)

        # --- Section: Calibrages Spéciaux ---
        ctk.CTkLabel(self.calib_container, text="Calibrages Spéciaux", font=ctk.CTkFont(weight="bold"), text_color=GOLD, anchor="w").pack(fill="x", padx=10, pady=(4, 2))
        card4 = ctk.CTkFrame(self.calib_container, fg_color=BG_CARD, corner_radius=8)
        card4.pack(fill="x", padx=8, pady=(0, 10))
        self._calib_btn(card4, "group_accept_pos", "Calibrer 'Accepter Groupe'", self.start_calib_group_accept)
        self._calib_btn(card4, "trade_notif_pos", "Calibrer 'Notif Échange'", self.start_calib_trade_notif)
        self._calib_btn(card4, "trade_validate_pos", "Calibrer pixel VERT (zone validée)", self.start_calib_trade_validate)
        self._calib_btn(card4, "trade_validate_click_pos", "Calibrer bouton VALIDER (relatif)", self.start_calib_trade_validate_click)

    def hard_kill_app(self):
        """ Fermeture propre de l'application """
        logger.info("Fermeture forcée via l'interface...")
        if hasattr(self.app, 'tray_icon'):
            self.app.tray_icon.stop()
        sys.exit(0)
        
    def open_settings(self):
        if not hasattr(self, 'settings_window') or not self.settings_window.winfo_exists():
            self.settings_window = SettingsWindow(self)
        else:
            self.settings_window.deiconify()
            self.settings_window.lift()
            self.settings_window.focus_force()

    def show_temporary_message(self, text, color="#2ecc71"):
        self.lbl_feedback.configure(text=text, text_color=color)
        if hasattr(self, "feedback_timer"):
            self.root.after_cancel(self.feedback_timer)
        self.feedback_timer = self.root.after(2500, lambda: self.lbl_feedback.configure(text=""))

    def change_position(self, name, new_val_str):
        new_index = int(new_val_str) - 1
        self.app.logic.set_account_position(name, new_index)
        self.original_app_refresh()

    def move_row(self, name, direction):
        self.app.logic.move_account(name, direction)
        self.original_app_refresh()

    def trigger_sort_taskbar(self):
        self.app.logic.sort_taskbar()
        self.show_temporary_message("🚀 Les pages ont été rangées avec succès !", GOLD)

    def close_and_refresh(self, name):
        self.app.logic.close_account_window(name)
        time.sleep(0.5) 
        self.original_app_refresh()
        
    def close_all_and_refresh(self):
        self.app.logic.close_all_active_accounts()
        time.sleep(0.5)
        self.original_app_refresh()
        self.show_temporary_message("💥 La team a été fermée !", GOLD)

    def toggle_tooltips_setting(self):
        self.app.config.data["show_tooltips"] = self.var_tooltips.get()
        self.app.config.save()
        if not self.var_tooltips.get():
            self.tooltip.withdraw()

    def show_gui(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.is_visible = True

    def hide_to_tray(self):
        self.root.withdraw()
        self.is_visible = False

    def toggle_visibility(self):
        if self.is_visible:
            self.hide_to_tray()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force() 
            self.is_visible = True
            
    def toggle_macros(self):
        self.app.config.data["spam_click_active"] = self.var_spam.get() 
        self.app.config.save()

    def toggle_toolbar(self):
        state = self.var_toolbar.get()
        self.app.config.data["toolbar_active"] = state
        self.app.config.save()
        if state: self.toolbar.show()
        else: self.toolbar.hide()

    def toggle_return(self):
        self.app.config.data["return_to_leader"] = self.var_return.get()
        self.app.config.save()

    def toggle_auto_group_opt(self):
        self.app.config.data["auto_group_enabled"] = self.var_auto_group.get()
        self.app.config.save()
        self.toolbar.update_overlay_icons()

    def toggle_auto_trade_opt(self):
        self.app.config.data["auto_trade_enabled"] = bool(self.var_auto_trade.get())
        self.app.config.save()
        self.toolbar.update_overlay_icons()

    def on_mode_change(self, choice):
        self.app.logic.set_mode(choice)
        self.app.current_idx = 0
        self.toolbar.combo_mode.set(choice)
        self.app.setup_hotkeys()

    def get_class_image(self, class_name):
        if class_name in self.skin_cache: return self.skin_cache[class_name]
        path = resource_path(f"skin/{class_name}.png")
        if not os.path.exists(path): return None
        img = ctk.CTkImage(light_image=Image.open(path), dark_image=Image.open(path), size=(24, 24))
        self.skin_cache[class_name] = img
        return img

    def set_leader(self, name):
        self.app.logic.set_leader(name)
        self.original_app_refresh()

    def reset_all(self):
        if messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir tout réinitialiser ?\n\nToutes vos touches seront perdues."):
            self.app.config.reset_settings()
            self.original_app_refresh()
            self.root.after(0, self.populate_calibrations)

    def show_tooltip(self, text):
        self.tooltip_lbl.config(text=text)
        self.tooltip.deiconify()
        self.tooltip.attributes("-topmost", True)
        self.tooltip.lift()
        self.update_tooltip_pos()

    def update_tooltip_pos(self):
        if self.is_listening:
            x, y = win32api.GetCursorPos()
            self.tooltip.geometry(f"+{x + 20}+{y + 20}")
            self.tooltip.lift()
            self.root.after(10, self.update_tooltip_pos)
        else:
            self.tooltip.withdraw()

    def bind_tooltip(self, widget, text):
        def on_enter(event):
            if self.is_listening or not self.app.config.data.get("show_tooltips", True): return 
            self.tooltip_lbl.config(text=text)
            self.tooltip.deiconify()
            self.tooltip.attributes("-topmost", True)
            self.tooltip.lift()
            x, y = win32api.GetCursorPos()
            self.tooltip.geometry(f"+{x + 15}+{y + 15}")
        def on_leave(event):
            if not self.is_listening: self.tooltip.withdraw()
        def on_motion(event):
            if self.is_listening or not self.app.config.data.get("show_tooltips", True): return
            x, y = win32api.GetCursorPos()
            self.tooltip.geometry(f"+{x + 15}+{y + 15}")

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
        widget.bind("<Motion>", on_motion)

    def toggle_team_ui(self, name, btn):
        current_team = self.app.config.data.get("accounts_team", {}).get(name, "Team 1")
        new_team = "Team 2" if current_team == "Team 1" else "Team 1"
        self.app.logic.change_team(name, new_team)
        team_color = GOLD if new_team == "Team 1" else "#4d4d4d" # Neutral dark gray for Team 2 instead of purple
        team_hover = GOLD_LIGHT if new_team == "Team 1" else "#5c5c5c"
        if btn:
            if new_team == "Team 2":
                 btn.configure(text="T2", fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD)
            else:
                 btn.configure(text="T1", fg_color=GOLD, text_color=BG_VOID, border_width=0, hover_color=GOLD_LIGHT)

    def catch_account_focus_key(self, pseudo, btn):
        if self.is_listening:
            return
        self.is_listening = True
        btn.configure(text="...", fg_color=GOLD)
        threading.Thread(target=self._listen_account_focus_key_thread, args=(pseudo, btn), daemon=True).start()

    def _listen_account_focus_key_thread(self, pseudo, btn):
        captured_key = None
        captured_mods = []

        while win32api.GetAsyncKeyState(win32con.VK_LBUTTON) < 0 or win32api.GetAsyncKeyState(win32con.VK_RBUTTON) < 0 or win32api.GetAsyncKeyState(win32con.VK_MBUTTON) < 0:
            time.sleep(0.01)
        time.sleep(0.1)

        def get_current_mods():
            mods = []
            if win32api.GetAsyncKeyState(win32con.VK_CONTROL) < 0:
                mods.append("ctrl")
            if win32api.GetAsyncKeyState(win32con.VK_MENU) < 0:
                mods.append("alt")
            if win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0:
                mods.append("shift")
            return mods

        while True:
            event = keyboard.read_event(suppress=True)
            if event.event_type == keyboard.KEY_DOWN:
                if event.name not in [
                    "alt", "ctrl", "shift", "maj", "right alt", "right ctrl",
                    "left alt", "left ctrl", "menu", "windows", "cmd",
                ]:
                    captured_mods = get_current_mods()
                    if event.scan_code in SCAN_TO_AZERTY:
                        captured_key = SCAN_TO_AZERTY[event.scan_code]
                    else:
                        captured_key = event.name
                    break

        if captured_key == "esc":
            self.root.after(0, self.apply_account_focus_hotkey, pseudo, "", btn)
            return

        final_key = "+".join(captured_mods) + "+" + captured_key if captured_mods else captured_key
        time.sleep(0.35)
        self.root.after(0, self.apply_account_focus_hotkey, pseudo, final_key, btn)

    def apply_account_focus_hotkey(self, pseudo, final_key, btn):
        self.release_modifiers()
        hk = self.app.config.data.setdefault("account_focus_hotkeys", {})
        if final_key:
            for other_p, v in list(hk.items()):
                if v == final_key and other_p != pseudo:
                    hk[other_p] = ""
            for k in list(self.app.config.data.keys()):
                if (k.endswith("_key") or k.endswith("_hotkey")) and self.app.config.data.get(k) == final_key:
                    self.app.config.data[k] = ""
                    if k in self.hotkey_btns:
                        self.hotkey_btns[k].configure(text="Aucun", fg_color=BG_BTN_DARK, text_color=GOLD, border_width=1, border_color=GOLD)
        hk[pseudo] = final_key or ""
        self.app.config.save()
        btn.configure(
            text=final_key.upper() if final_key else "···",
            fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD,
            hover_color=BG_CARD,
        )
        self.app.setup_hotkeys()
        self.is_listening = False
        self.show_temporary_message("✅ Touche perso enregistrée", "#2ecc71")
        self.root.after(0, self.populate_calibrations)
        self.root.after(0, self.original_app_refresh)

    def clear_account_focus_key(self, pseudo, btn):
        if self.is_listening:
            return
        self.apply_account_focus_hotkey(pseudo, "", btn)

    def refresh_list(self, accounts):
        for widget in self.scroll_frame.winfo_children(): widget.destroy()
        leader_name = self.app.config.data.get("leader_name", "")
        
        def make_toggle_cmd(n, v_obj):
            return lambda: self.app.logic.toggle_account(n, v_obj.get())
        
        for idx, acc in enumerate(accounts):
            row_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)
            
            img = self.get_class_image(acc.get('classe', 'Inconnu'))
            if img: ctk.CTkLabel(row_frame, image=img, text="").pack(side="left", padx=5)
            else: ctk.CTkLabel(row_frame, text="👤", text_color=GOLD).pack(side="left", padx=5) 
            
            var = tk.BooleanVar(value=acc['active'])
            self.checkbox_vars[acc['name']] = var 
            
            chk = ctk.CTkCheckBox(row_frame, text=acc['name'][:20], width=160, fg_color=GOLD, hover_color=GOLD_LIGHT)
            if acc['active']:
                chk.select()
            else:
                chk.deselect()
                
            chk.configure(command=lambda n=acc['name'], c=chk: self.app.logic.toggle_account(n, bool(c.get())))
            chk.pack(side="left", padx=(5, 0))

            btn_close = ctk.CTkButton(row_frame, text="✖", width=25, fg_color=BG_BTN_DARK, text_color=GOLD, border_width=1, border_color=GOLD, hover_color=BG_CARD, command=lambda n=acc['name']: self.close_and_refresh(n))
            btn_close.pack(side="right", padx=(2, 5))
            self.bind_tooltip(btn_close, "Fermer la fenêtre")
            
            is_leader = (acc['name'] == leader_name)
            leader_txt = "★" if is_leader else "☆"
            btn_lead = ctk.CTkButton(row_frame, text=leader_txt, text_color=GOLD if is_leader else "#b4b4b4", width=35, fg_color="transparent", border_width=1, border_color=GOLD if is_leader else BG_BTN_DARK, hover_color=BG_CARD, command=lambda n=acc['name']: self.set_leader(n))
            btn_lead.pack(side="right", padx=2)
            self.bind_tooltip(btn_lead, "Définir comme Chef")

            team_val = acc.get('team', 'Team 1')
            if team_val == "Team 1":
                btn_team = ctk.CTkButton(row_frame, text="T1", width=35, fg_color=GOLD, text_color=BG_VOID, hover_color=GOLD_LIGHT)
            else:
                btn_team = ctk.CTkButton(row_frame, text="T2", width=35, fg_color=BG_BTN_DARK, text_color=GOLD, border_width=1, border_color=GOLD, hover_color=BG_CARD)
            btn_team.configure(command=lambda n=acc['name'], b=btn_team: self.toggle_team_ui(n, b))
            btn_team.pack(side="right", padx=5)
            self.bind_tooltip(btn_team, "Changer l'équipe")

            hk_val = self.app.config.data.get("account_focus_hotkeys", {}).get(acc["name"], "") or ""
            btn_afk = ctk.CTkButton(
                row_frame,
                text=hk_val.upper() if hk_val else "···",
                width=40,
                height=26,
                fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD,
                hover_color=BG_CARD,
                font=ctk.CTkFont(size=11, weight="bold"),
            )
            btn_afk.configure(command=lambda n=acc["name"], b=btn_afk: self.catch_account_focus_key(n, b))
            btn_afk.pack(side="right", padx=(0, 2))
            self.bind_tooltip(btn_afk, "Touche pour focus ce perso (auto-échange / switch)")

            btn_afk_clr = ctk.CTkButton(
                row_frame, text="✖", width=22, height=26,
                fg_color=BG_CARD, hover_color=BG_BTN_DARK, text_color=GOLD,
                command=lambda n=acc["name"], b=btn_afk: self.clear_account_focus_key(n, b),
            )
            btn_afk_clr.pack(side="right", padx=(0, 4))
            self.bind_tooltip(btn_afk_clr, "Effacer la touche perso")

            btn_down = ctk.CTkButton(row_frame, text="▼", width=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, command=lambda n=acc['name']: self.move_row(n, 1))
            btn_down.pack(side="right", padx=(2, 10))
            self.bind_tooltip(btn_down, "Descendre")
            
            btn_up = ctk.CTkButton(row_frame, text="▲", width=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, command=lambda n=acc['name']: self.move_row(n, -1))
            btn_up.pack(side="right", padx=2)
            self.bind_tooltip(btn_up, "Monter")

            pos_values = [str(i+1) for i in range(len(accounts))]
            current_pos = str(idx + 1)
            combo_pos = ctk.CTkOptionMenu(row_frame, values=pos_values, width=50, height=24, fg_color=BG_BTN_DARK, text_color=GOLD, button_color=GOLD, button_hover_color=GOLD_LIGHT, command=lambda val, n=acc['name']: self.change_position(n, val))
            combo_pos.set(current_pos)
            combo_pos.pack(side="right", padx=(2, 5))
            self.bind_tooltip(combo_pos, "Position exacte")

    def on_volume_change(self, value): 
        self.app.update_volume(int(value))
    
    def create_hotkey_row(self, parent, label_text, config_key, row, col_offset, tooltip_txt=""):
        lbl = ctk.CTkLabel(parent, text=f"{label_text}:", text_color=GOLD)
        lbl.grid(row=row, column=col_offset, padx=5, sticky="w")
        if tooltip_txt: self.bind_tooltip(lbl, tooltip_txt)
        
        current_val = self.app.config.data.get(config_key, "")
        btn = ctk.CTkButton(parent, text=current_val if current_val else "Aucun", width=80, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD,
                            command=lambda: self.catch_key(config_key, btn, allow_mouse=True))
        btn.grid(row=row, column=col_offset+1, padx=2, pady=2)
        
        self.hotkey_btns[config_key] = btn 
        
        btn_x = ctk.CTkButton(parent, text="✖", width=25, fg_color=BG_BTN_DARK, border_width=1, border_color=GOLD, text_color=GOLD, hover_color=BG_CARD, command=lambda: self.clear_key(config_key, btn))
        btn_x.grid(row=row, column=col_offset+2, padx=(0, 10))
        self.bind_tooltip(btn_x, "Effacer le raccourci")

    def catch_key(self, config_key, btn, allow_mouse=False):
        if self.is_listening: return 
        self.is_listening = True
        btn.configure(text="...", fg_color=GOLD)
        threading.Thread(target=self._listen_hotkey_thread, args=(config_key, btn, allow_mouse), daemon=True).start()

    def _listen_hotkey_thread(self, config_key, btn, allow_mouse):
        captured_key = None
        captured_mods = []
        
        while win32api.GetAsyncKeyState(win32con.VK_LBUTTON) < 0 or win32api.GetAsyncKeyState(win32con.VK_RBUTTON) < 0 or win32api.GetAsyncKeyState(win32con.VK_MBUTTON) < 0:
            time.sleep(0.01)
        time.sleep(0.1) 
        
        def get_current_mods():
            mods = []
            if win32api.GetAsyncKeyState(win32con.VK_CONTROL) < 0: mods.append("ctrl")
            if win32api.GetAsyncKeyState(win32con.VK_MENU) < 0: mods.append("alt")
            if win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0: mods.append("shift")
            return mods

        if not allow_mouse:
            while True:
                event = keyboard.read_event(suppress=True)
                if event.event_type == keyboard.KEY_DOWN:
                    if event.name not in ['alt', 'ctrl', 'shift', 'maj', 'right alt', 'right ctrl', 'left alt', 'left ctrl', 'menu', 'windows', 'cmd']:
                        captured_mods = get_current_mods()
                        
                        if event.scan_code in SCAN_TO_AZERTY:
                            captured_key = SCAN_TO_AZERTY[event.scan_code]
                        else:
                            captured_key = event.name
                        break
        else:
            def on_key(e):
                nonlocal captured_key, captured_mods
                if e.event_type == keyboard.KEY_DOWN:
                    if e.name not in ['alt', 'ctrl', 'shift', 'maj', 'right alt', 'right ctrl', 'left alt', 'left ctrl', 'menu']:
                        captured_mods = get_current_mods()
                        if e.scan_code in SCAN_TO_AZERTY:
                            captured_key = SCAN_TO_AZERTY[e.scan_code]
                        else:
                            captured_key = e.name
            hook = keyboard.hook(on_key, suppress=True)
            
            while not captured_key:
                if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) < 0: captured_key = "left_click"
                elif win32api.GetAsyncKeyState(win32con.VK_RBUTTON) < 0: captured_key = "right_click"
                elif win32api.GetAsyncKeyState(win32con.VK_MBUTTON) < 0: captured_key = "middle_click"
                elif win32api.GetAsyncKeyState(0x05) < 0: captured_key = "mouse4" 
                elif win32api.GetAsyncKeyState(0x06) < 0: captured_key = "mouse5"
                
                if captured_key:
                    captured_mods = get_current_mods()
                    break
                time.sleep(0.01)
                
            keyboard.unhook(hook)

        if captured_key == "esc":
            final_key = self.app.config.data.get(config_key, "")
        else:
            final_key = "+".join(captured_mods) + "+" + captured_key if captured_mods else captured_key

        time.sleep(0.5)


        self.root.after(0, self.apply_single_hotkey, config_key, final_key, btn)

    def release_modifiers(self):
       
        try:
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
        except: pass

    def apply_single_hotkey(self, config_key, new_value, btn):
        self.release_modifiers()
        
        if new_value: 
            for k in list(self.app.config.data.keys()):
                if (k.endswith("_key") or k.endswith("_hotkey")) and k != config_key:
                    if self.app.config.data[k] == new_value:
                        self.app.config.data[k] = ""
                        if k in self.hotkey_btns:
                            self.hotkey_btns[k].configure(text="Aucun", fg_color=BG_BTN_DARK, text_color=GOLD, border_width=1, border_color=GOLD)
            afh = self.app.config.data.setdefault("account_focus_hotkeys", {})
            for p, v in list(afh.items()):
                if v == new_value:
                    afh[p] = ""

        self.app.config.data[config_key] = new_value
        self.app.config.save()
        
        btn.configure(text=new_value if new_value else "Aucun", fg_color=BG_BTN_DARK, text_color=GOLD, border_width=1, border_color=GOLD)
        self.app.setup_hotkeys()
        self.is_listening = False
        self.root.after(0, self.populate_calibrations)
        self.root.after(0, self.original_app_refresh)

    def clear_key(self, config_key, btn):
        if self.is_listening: return
        self.apply_single_hotkey(config_key, "", btn)

    def wait_for_calib_or_esc(self):
       
        calib_key = self.app.config.data.get("calib_key", "f4").lower()
        if not calib_key: calib_key = "f4"
        
        while True:
            event = keyboard.read_event(suppress=True)
            if event.event_type == keyboard.KEY_DOWN:
                if event.name == calib_key: return True
                elif event.name == 'esc': return False

    def start_calib_chat(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw() 
        if self.app.logic.leader_hwnd: 
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2) 
        threading.Thread(target=self.calibration_chat_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Cliquez dans votre chat dofus pour envoyer un message, puis appuyez sur {k} sur la position.\n(Echap pour annuler)")

    def calibration_chat_sequence(self):
        if self.wait_for_calib_or_esc():
            rx, ry = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
            self.app.config.data["macro_positions"]["chat_position"] = [rx, ry]
            self.app.config.save()
            self.root.after(0, self.populate_calibrations)
            self.root.after(0, lambda: self.show_temporary_message("✅ Chat calibré !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui) 

    def start_calib_xp_drop(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw() 
        if self.app.logic.leader_hwnd: 
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_xp_drop_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Lancez un combat, placez votre souris sur le bouton XP/Drop de fin de combat, puis appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_xp_drop_sequence(self):
        if self.wait_for_calib_or_esc():
            rx, ry = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
            self.app.config.data["macro_positions"]["xp_drop_button"] = [rx, ry]
            self.app.config.save()
            self.root.after(0, self.populate_calibrations)
            self.root.after(0, lambda: self.show_temporary_message("✅ XP/Drop calibré !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui) 

    def start_calib_zaap(self):
        if self.is_listening: return
        active_accounts = self.app.logic.get_cycle_list()
        if not active_accounts: return
        self.is_listening = True
        self.root.withdraw() 
        threading.Thread(target=self.calibration_zaap_sequence, args=(active_accounts,), daemon=True).start()

    def calibration_zaap_sequence(self, active_accounts):
        self.app.config.data["macro_positions"]["zaaps"] = {}
        success = True
        k = self.app.config.data.get("calib_key", "F4").upper()
        
        for acc in active_accounts:
            self.app.logic.focus_window(acc['hwnd'])
            time.sleep(0.2) 
            self.root.after(0, lambda a=acc: self.show_tooltip(f"Allez dans le havre-sac de {a['name']}, placez votre souris sur le haut du Zaap, puis appuyez sur {k}.\n(Echap pour annuler)"))
            
            if not self.wait_for_calib_or_esc():
                self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Zaap annulée.", GOLD))
                success = False
                break
                
            rx, ry = self.app.logic.get_relative_ratio_pos(acc['hwnd'])
            self.app.config.data["macro_positions"]["zaaps"][acc['name']] = [rx, ry]
            self.root.after(0, lambda a=acc: self.show_temporary_message(f"✅ Zaap de {a['name']} calibré !", "#2ecc71"))
        
        if success:
            self.app.config.save()
            self.root.after(0, self.populate_calibrations)
            self.root.after(0, lambda: self.show_temporary_message("✅ Calibration Zaap totale terminée !", "#2ecc71"))
            
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_zone_jeu(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_zone_jeu_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Placez votre souris EN HAUT À GAUCHE de la zone de jeu (exclut interfaces), puis appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_zone_jeu_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        # Point 1 (Top-Left)
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Zone annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return
            
        top_left = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        self.root.after(0, lambda: self.show_tooltip(f"Parfait ! Maintenant, placez la souris EN BAS À DROITE de la zone de jeu et appuyez sur {k}.\n(Echap pour annuler)"))
        time.sleep(0.5) # small debounce
        
        # Point 2 (Bottom-Right)
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Zone annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        bot_right = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        
        # Save as [x1, y1, x2, y2]
        self.app.config.data.setdefault("macro_positions", {})["game_zone"] = [
            min(top_left[0], bot_right[0]), min(top_left[1], bot_right[1]),
            max(top_left[0], bot_right[0]), max(top_left[1], bot_right[1])
        ]
        self.app.config.save()
        self.root.after(0, self.populate_calibrations)
        self.root.after(0, lambda: self.show_temporary_message("✅ Zone Jeu calibrée !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_coord(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_coord_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Placez votre souris EN HAUT À GAUCHE de la boîte des Coordonnées, puis appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_coord_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Coordonnées annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return
            
        top_left = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        self.root.after(0, lambda: self.show_tooltip(f"Parfait ! Maintenant, placez la souris EN BAS À DROITE de la boîte des Coordonnées et appuyez sur {k}.\n(Echap pour annuler)"))
        time.sleep(0.5)
        
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Coordonnées annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        bot_right = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        
        self.app.config.data.setdefault("macro_positions", {})["map_coords"] = [
            min(top_left[0], bot_right[0]), min(top_left[1], bot_right[1]),
            max(top_left[0], bot_right[0]), max(top_left[1], bot_right[1])
        ]
        self.app.config.save()
        self.root.after(0, self.populate_calibrations)
        self.root.after(0, lambda: self.show_temporary_message("✅ Coordonnées calibrées !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_map_borders(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_map_borders_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Placez la souris sur le passage pour changer de map HAUT puis appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_map_borders_sequence(self):
        borders = {"top": "HAUT", "bottom": "BAS", "left": "GAUCHE", "right": "DROITE"}
        results = {}
        k = self.app.config.data.get("calib_key", "F4").upper()
        
        for key, name in borders.items():
            self.root.after(0, lambda n=name: self.show_tooltip(f"Placez la souris sur le passage pour changer de map {n} puis appuyez sur {k}.\n(Echap pour annuler)"))
            time.sleep(0.5) # small debounce
            if not self.wait_for_calib_or_esc():
                self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Map annulée.", GOLD))
                self.is_listening = False
                self.root.after(0, self.tooltip.withdraw)
                self.root.after(0, self.show_gui)
                return
            results[key] = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
            
        self.app.config.data.setdefault("macro_positions", {})["map_borders"] = results
        self.app.config.save()
        self.root.after(0, self.populate_calibrations)
        self.root.after(0, lambda: self.show_temporary_message("✅ Changements de Map calibrés !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)



    def start_calib_group_accept(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_group_accept_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Ouvrez une invitation de groupe. Placez la souris sur le bouton OUI (ou Accepter) et appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_group_accept_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Groupe annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        pos = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        if pos:
            self.app.config.data.setdefault("macro_positions", {})["group_accept_pos"] = list(pos)
            self.app.config.save()
            self.root.after(0, self.populate_calibrations)
        
        self.root.after(0, lambda: self.show_temporary_message("✅ Accepter Groupe calibré !", "#2ecc71"))
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_trade_notif(self):
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_trade_notif_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"Faites apparaître une notification d'échange Dofus.\nPlacez la souris sur l'ICÔNE DOFUS (le logo orange) dans la notif puis appuyez sur {k}.\n(Echap pour annuler)")

    def calibration_trade_notif_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Scanner annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        pos = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)
        
        from PIL import ImageGrab
        coords = self.app.logic.get_screen_coords_from_saved(self.app.logic.leader_hwnd, pos)
        if coords:
            x, y = coords
            try:
                img = ImageGrab.grab(bbox=(x, y, x+1, y+1))
                r, g, b = img.getpixel((0, 0))[:3]
                self.app.config.data.setdefault("macro_positions", {})["trade_notif_pos"] = [pos[0], pos[1], r, g, b]
                self.app.config.save()
                self.root.after(0, self.populate_calibrations)
                self.root.after(0, lambda: self.show_temporary_message("✅ Pixel 1 Échange calibré !", "#2ecc71"))
            except Exception as e:
                self.root.after(0, lambda: self.show_temporary_message("❌ Erreur de capture Scanner.", GOLD))
        
        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_trade_notif2(self):
        """Point 2 : calibre un 2ème pixel de contrôle pour éviter les faux positifs Windows."""
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        if self.app.logic.leader_hwnd:
            self.app.logic.focus_window(self.app.logic.leader_hwnd)
            time.sleep(0.2)
        threading.Thread(target=self.calibration_trade_notif2_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(f"[OPTIONNEL] Faites apparaître une notification d'échange Dofus.\nPlacez la souris sur le FOND SOMBRE de la notification (pas l'icône) puis appuyez sur {k}.\nCe 2ème pixel est optionnel - il sert uniquement si les pop-ups Windows déclenchent le bot.\n(Echap pour annuler)")

    def calibration_trade_notif2_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Pixel 2 annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        pos = self.app.logic.get_relative_ratio_pos(self.app.logic.leader_hwnd)

        from PIL import ImageGrab
        coords = self.app.logic.get_screen_coords_from_saved(self.app.logic.leader_hwnd, pos)
        if coords:
            x, y = coords
            try:
                img = ImageGrab.grab(bbox=(x, y, x+1, y+1))
                r, g, b = img.getpixel((0, 0))[:3]
                self.app.config.data.setdefault("macro_positions", {})["trade_notif_pos2"] = [pos[0], pos[1], r, g, b]
                self.app.config.save()
                self.root.after(0, self.populate_calibrations)
                self.root.after(0, lambda: self.show_temporary_message("✅ Pixel 2 Anti-Faux Positif calibré !", "#2ecc71"))
            except Exception as e:
                self.root.after(0, lambda: self.show_temporary_message("❌ Erreur de capture Pixel 2.", GOLD))

        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_trade_validate(self):
        """Pixel vert : même logique sur chaque client (fenêtre active au moment du F4)."""
        if self.is_listening: return
        self.is_listening = True
        self.root.withdraw()
        threading.Thread(target=self.calibration_trade_validate_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(
            f"Sur une fenêtre d'échange : quand la case adverse devient VERTE, place la souris sur ce vert puis {k}.\n"
            f"(Même réglage pour les deux persos — le script détecte le vert sur la fenêtre active.)\n"
            f"Puis calibre le bouton VALIDER (position relative, identique sur les deux clients).\n(Echap pour annuler)"
        )

    def calibration_trade_validate_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Bouton Vert annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        fg = win32gui.GetForegroundWindow()
        known = {a["hwnd"] for a in self.app.logic.get_cycle_list()}
        if fg not in known:
            self.root.after(0, lambda: self.show_temporary_message(
                "❌ Mettez une fenêtre Dofus du cycle au premier plan.", GOLD
            ))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        pos = self.app.logic.get_relative_ratio_pos(fg)

        from PIL import ImageGrab
        coords = self.app.logic.get_screen_coords_from_saved(fg, pos)
        if coords:
            x, y = coords
            try:
                img = ImageGrab.grab(bbox=(x, y, x+1, y+1))
                r, g, b = img.getpixel((0, 0))[:3]
                self.app.config.data.setdefault("macro_positions", {})["trade_validate_pos"] = [pos[0], pos[1], r, g, b]
                self.app.config.save()
                self.root.after(0, self.populate_calibrations)
                self.root.after(0, lambda rv=(r,g,b): self.show_temporary_message(
                    f"✅ Bouton Vert calibré ! (RGB: {rv[0]},{rv[1]},{rv[2]})", "#2ecc71"
                ))
            except Exception as e:
                self.root.after(0, lambda: self.show_temporary_message("❌ Erreur de capture Bouton Vert.", GOLD))

        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)

    def start_calib_trade_validate_click(self):
        if self.is_listening:
            return
        self.is_listening = True
        self.root.withdraw()
        threading.Thread(target=self.calibration_trade_validate_click_sequence, daemon=True).start()
        k = self.app.config.data.get("calib_key", "F4").upper()
        self.show_tooltip(
            f"Ouvre l'échange (popup Valider visible), place la souris sur le centre du bouton VALIDER puis {k}.\n"
            f"La position relative et la couleur du pixel sont mémorisées pour détecter le popup sur chaque fenêtre.\n"
            f"(Echap pour annuler)"
        )

    def calibration_trade_validate_click_sequence(self):
        k = self.app.config.data.get("calib_key", "F4").upper()
        if not self.wait_for_calib_or_esc():
            self.root.after(0, lambda: self.show_temporary_message("❌ Calibration Valider annulée.", GOLD))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        fg = win32gui.GetForegroundWindow()
        known = {a["hwnd"] for a in self.app.logic.get_cycle_list()}
        if fg not in known:
            self.root.after(0, lambda: self.show_temporary_message(
                "❌ Mettez la fenêtre Dofus du bon perso au premier plan.", GOLD
            ))
            self.is_listening = False
            self.root.after(0, self.tooltip.withdraw)
            self.root.after(0, self.show_gui)
            return

        pos = self.app.logic.get_relative_ratio_pos(fg)
        if pos:
            from PIL import ImageGrab

            coords = self.app.logic.get_screen_coords_from_saved(fg, pos)
            if coords:
                x, y = coords
                try:
                    img = ImageGrab.grab(bbox=(x, y, x + 1, y + 1))
                    r, g, b = img.getpixel((0, 0))[:3]
                    self.app.config.data.setdefault("macro_positions", {})[
                        "trade_validate_click_pos"
                    ] = [pos[0], pos[1], r, g, b]
                except Exception:
                    self.app.config.data.setdefault("macro_positions", {})[
                        "trade_validate_click_pos"
                    ] = [pos[0], pos[1]]
            else:
                self.app.config.data.setdefault("macro_positions", {})[
                    "trade_validate_click_pos"
                ] = [pos[0], pos[1]]
            self.app.config.save()
            self.root.after(0, self.populate_calibrations)
            self.root.after(0, lambda: self.show_temporary_message("✅ Bouton VALIDER calibré !", "#2ecc71"))

        self.is_listening = False
        self.root.after(0, self.tooltip.withdraw)
        self.root.after(0, self.show_gui)



    def run(self): self.root.mainloop()