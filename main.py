import os
import sys
import ctypes
import threading
import keyboard
import requests
import time
import win32api
import win32con
import win32gui
import win32process
import pystray
import shutil
import subprocess
import logging
import traceback
import hashlib
from PIL import Image
from pystray import MenuItem as item
from constants import AZERTY_TO_SCAN
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox

from logging.handlers import RotatingFileHandler

# Configuration du logging
logger = logging.getLogger("Dostrub")
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Handler console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)

# Handler fichier rotatif (2 Mo, max 3 fichiers)
file_handler = RotatingFileHandler(
    "dostrub.log", maxBytes=2*1024*1024, backupCount=3, encoding="utf-8"
)
file_handler.setFormatter(formatter)

logger.addHandler(console_handler)
logger.addHandler(file_handler)

def get_current_version():
    """Charge la version depuis version.txt ou utilise la valeur par défaut."""
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        v_path = os.path.join(base_path, "version.txt")
        if os.path.exists(v_path):
            with open(v_path, "r", encoding="utf-8") as f:
                return f.read().strip()
    except Exception:
        pass
    return "1.1.0"

CURRENT_VERSION = get_current_version()
GITHUB_RAW = "https://raw.githubusercontent.com/Dostrubdev/dostrub/main"
VERSION_URL = f"{GITHUB_RAW}/version.txt"
DOWNLOAD_URL = "https://github.com/Dostrubdev/dostrub/releases/latest/download/Dostrub.exe"


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


try:
    ctypes.windll.user32.LoadKeyboardLayoutW("0000040C", 1 | 0x00000100)
except Exception:
    pass

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(application_path)

from config_manager import Config
from logic import DofusLogic
from gui import OrganizerGUI
from radial_menu import RadialMenu




class OrganizerApp:
    def __init__(self):
        self.config = Config()
        self.current_idx = 0
        self.hotkey_actions = {}
        self.logic = DofusLogic(self.config)
        self.logic._trade_focus_leader_fn = self.focus_leader
        self.logic._trade_run_valider_scan_fn = self.run_auto_trade_valider_scan
        self.version = CURRENT_VERSION
        self._action_lock = threading.Lock()
        self._last_action_time = {}
        self._switching = False
        self.gui = OrganizerGUI(self)

        self.logic.set_error_callback(self.show_calibration_error)

        self.radial_focus = RadialMenu(
            self.gui.root, self.on_radial_focus_select, accent_color="#444444"
        )
        saved_vol = self.config.data.get("volume_level", 50) / 100.0
        self.radial_focus.set_base_volume(saved_vol)

        threading.Thread(target=self.background_listener, daemon=True).start()

        self.setup_hotkeys()
        self.refresh()
        self.setup_system_tray()

        if not self.config.data.get("tutorial_done", False):
            self.config.data["tutorial_done"] = True
            self.config.save()

        self.gui.root.after(1000, self.check_conflicting_software)

    def setup_system_tray(self):

        icon_path = "Dostrub.ico"

        try:
            image = Image.open(icon_path)
        except Exception:
            logger.warning(f"Impossible de charger l'icône {icon_path}, utilisation d'une image par défaut.")
            image = Image.new("RGB", (64, 64), color=(44, 62, 80))

        menu = pystray.Menu(
            item("Afficher/Cacher", self.toggle_from_tray, default=True),
            item("Quitter", self.quit_from_tray),
        )

        self.tray_icon = pystray.Icon("dosoft_tray", image, "Dostrub", menu)
        self.tray_icon.run_detached()

    def toggle_from_tray(self, icon, item):
        def safe_toggle():
            if self.gui.root.state() == "withdrawn":
                self.gui.root.deiconify()
                self.gui.root.lift()
                self.gui.root.focus_force()
            else:
                self.gui.root.withdraw()

        self.gui.root.after(0, safe_toggle)

    def quit_from_tray(self, icon, item):

        self.tray_icon.stop()
        self.gui.root.after(0, self.gui.root.destroy)

    def check_conflicting_software(self):

        if self.config.data.get("ignore_organizer_warning", False):
            return

        try:
            output = subprocess.check_output(
                'tasklist /FI "IMAGENAME eq organizer.exe" /NH',
                shell=True,
                stderr=subprocess.STDOUT
            ).decode(errors='ignore').lower()

            if "organizer.exe" in output:
                self.show_conflict_popup()
        except Exception as e:
            logger.error(f"Erreur lors de la vérification des logiciels conflictuels : {e}")

    def show_conflict_popup(self):

        popup = ctk.CTkToplevel(self.gui.root)
        popup.title("⚠️ Conflit de logiciels détecté")
        popup.geometry("480x250")
        popup.attributes("-topmost", True)
        popup.resizable(False, False)
        popup.transient(self.gui.root)

        popup.update_idletasks()
        x = self.gui.root.winfo_x() + (self.gui.root.winfo_width() // 2) - (480 // 2)
        y = self.gui.root.winfo_y() + (self.gui.root.winfo_height() // 2) - (250 // 2)
        popup.geometry(f"+{x}+{y}")

        msg = (
            "Le logiciel 'Organizer' est actuellement ouvert.\n"
            "L'utilisation de deux gestionnaires de pages simultanément\n"
            "va créer des bugs et des conflits de focus sur Dostrub.\n\n"
            "Nous vous recommandons fortement de le fermer."
        )

        lbl = ctk.CTkLabel(popup, text=msg, justify="center", font=ctk.CTkFont(size=13))
        lbl.pack(pady=(20, 15))

        var_ignore = ctk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(
            popup, text="Ne plus m'afficher cet avertissement", variable=var_ignore
        )
        chk.pack(pady=(0, 20))

        frame_btn = ctk.CTkFrame(popup, fg_color="transparent")
        frame_btn.pack(fill="x", padx=20)

        def on_close_organizer():
            if var_ignore.get():
                self.config.data["ignore_organizer_warning"] = True
                self.config.save()

            subprocess.run(["taskkill", "/F", "/IM", "organizer.exe", "/T"], shell=True, capture_output=True)

            popup.destroy()
            self.gui.show_temporary_message(
                "✅ Organizer fermé avec succès !", "#444444"
            )

        def on_keep_organizer():
            if var_ignore.get():
                self.config.data["ignore_organizer_warning"] = True
                self.config.save()
            popup.destroy()

        btn_close = ctk.CTkButton(
            frame_btn, text="Fermer Organizer", fg_color="#444444", command=on_close_organizer
        )
        btn_close.pack(side="left", expand=True, padx=10)

        btn_keep = ctk.CTkButton(frame_btn, text="Conserver", fg_color="#444444", command=on_keep_organizer)
        btn_keep.pack(side="right", expand=True, padx=10)

        popup.grab_set()

    def show_calibration_error(self, msg):
        def trigger_error():
            if not self.gui.is_visible:
                self.gui.show_gui()
            messagebox.showwarning("Action Impossible", msg)

        self.gui.root.after(0, trigger_error)

    def update_volume(self, volume_val):
        self.config.data["volume_level"] = volume_val
        self.config.save()
        vol_float = volume_val / 100.0
        self.radial_focus.set_base_volume(vol_float)

    def get_vk(self, key_str):
        key_str = key_str.lower().strip()
        mapping = {
            "alt": win32con.VK_MENU,
            "ctrl": win32con.VK_CONTROL,
            "shift": win32con.VK_SHIFT,
            "left_click": 0x01,
            "right_click": 0x02,
            "middle_click": 0x04,
            "mouse4": 0x05,
            "mouse5": 0x06,
        }
        if key_str in mapping:
            return mapping[key_str]

        scan_code = AZERTY_TO_SCAN.get(key_str)
        if scan_code is not None:
            vk = ctypes.windll.user32.MapVirtualKeyW(scan_code, 1)
            if vk:
                return vk

        if len(key_str) == 1:
            return ord(key_str.upper())
        if key_str.startswith("f") and key_str[1:].isdigit():
            return 0x6F + int(key_str[1:])
        return None


    def background_listener(self):
        radial_was_open = False

        while True:

            if hasattr(self, "mouse_hotkeys"):
                for hk_str, func in self.mouse_hotkeys.items():
                    is_pressed = self.is_hotkey_pressed(hk_str)
                    was_pressed = self.mouse_states.get(hk_str, False)

                    if is_pressed and not was_pressed:
                        self.mouse_states[hk_str] = True

                        def safe_mouse_execute(f=func):
                            self.release_modifiers()
                            f()

                        threading.Thread(target=safe_mouse_execute, daemon=True).start()
                    elif not is_pressed and was_pressed:
                        self.mouse_states[hk_str] = False

            m_pressed = win32api.GetAsyncKeyState(win32con.VK_MBUTTON) < 0
            if m_pressed and self.config.data.get("spam_click_active", False):
                fg_hwnd = win32gui.GetForegroundWindow()
                is_dofus = any(
                    acc["hwnd"] == fg_hwnd for acc in self.logic.all_accounts
                )
                if is_dofus:
                    while win32api.GetAsyncKeyState(win32con.VK_MBUTTON) < 0:
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                        time.sleep(0.02)
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        time.sleep(0.02)
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                        time.sleep(0.02)
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        time.sleep(0.20)
                    time.sleep(0.05)

            radial_hk = self.config.data.get("radial_menu_hotkey", "")
            radial_active = self.config.data.get("radial_menu_active", True)

            if radial_active and radial_hk:
                is_pressed = self.is_hotkey_pressed(radial_hk)

                if is_pressed and not radial_was_open:
                    radial_was_open = True
                    active_accs = [
                        {
                            "name": acc["name"],
                            "classe": acc.get("classe", "Inconnu"),
                            "hwnd": acc["hwnd"],
                        }
                        for acc in self.logic.get_cycle_list()
                    ]

                    fg_hwnd = win32gui.GetForegroundWindow()
                    current_name = None
                    for acc in active_accs:
                        if acc["hwnd"] == fg_hwnd:
                            current_name = acc["name"]
                            break

                    x, y = win32api.GetCursorPos()

                    self.gui.root.after(
                        0, self.radial_focus.show, x, y, active_accs, current_name
                    )

                elif radial_was_open and not is_pressed:
                    radial_was_open = False
                    self.gui.root.after(0, self.radial_focus.hide)

            try:
                if not self._switching:
                    fg_hwnd = win32gui.GetForegroundWindow()
                    cycle_list = self.logic.get_cycle_list()
                    if cycle_list:
                        for index, acc in enumerate(cycle_list):
                            if acc["hwnd"] == fg_hwnd:
                                if self.current_idx != index:
                                    self.current_idx = index
                                break
            except Exception:
                pass

            time.sleep(0.01)

    def on_radial_focus_select(self, target_name):
        self._switching = True
        try:
            for acc in self.logic.all_accounts:
                if acc["name"] == target_name:
                    self.logic.focus_window(acc["hwnd"])
                    break

            cycle_list = self.logic.get_cycle_list()
            for index, acc in enumerate(cycle_list):
                if acc["name"] == target_name:
                    self.current_idx = index
                    break
        finally:
            self._switching = False

    def is_hotkey_pressed(self, hk_str):

        if not hk_str:
            return False
        parts = hk_str.split("+")
        for p in parts:
            vk = self.get_vk(p.strip())
            if not vk:
                return False
            if win32api.GetAsyncKeyState(vk) >= 0:
                return False
        return True

    def focus_account_by_name(self, name):
        self._switching = True
        try:
            cycle_list = self.logic.get_cycle_list()
            for index, acc in enumerate(cycle_list):
                if acc["name"] == name:
                    self.logic.focus_window(acc["hwnd"])
                    self.current_idx = index
                    return
        finally:
            self._switching = False

    def register_action(self, hk_str, func):
        if not hk_str:
            return
        parts = hk_str.lower().split("+")

        if "click" in hk_str or "mouse" in hk_str:
            self.mouse_hotkeys[hk_str] = func
            return

        mods = set()
        main_scan = None
        for p in parts:
            if p in ["ctrl", "alt", "shift"]:
                mods.add(p)
            elif p in AZERTY_TO_SCAN:
                main_scan = AZERTY_TO_SCAN[p]
            else:
                try:
                    main_scan = keyboard.key_to_scan_codes(p)[0]
                except Exception:
                    pass
        if main_scan is not None:
            self.hotkey_actions[(frozenset(mods), main_scan)] = func

    def release_modifiers(self):
        try:
            if win32api.GetAsyncKeyState(win32con.VK_MENU) < 0:
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            if win32api.GetAsyncKeyState(win32con.VK_CONTROL) < 0:
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            if win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0:
                win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            if win32api.GetAsyncKeyState(0x05) < 0:  # XButton1
                win32api.mouse_event(0x0100, 0, 0, 0x0001, 0)
            if win32api.GetAsyncKeyState(0x06) < 0:  # XButton2
                win32api.mouse_event(0x0100, 0, 0, 0x0002, 0)
        except:
            pass

    def restore_modifiers(self, mods):

        try:

            if "alt" in mods and win32api.GetAsyncKeyState(win32con.VK_MENU) < 0:
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)

            if "ctrl" in mods and win32api.GetAsyncKeyState(win32con.VK_CONTROL) < 0:
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)

            if "shift" in mods and win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0:
                win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
        except:
            pass

    def global_hook_listener(self, event):
        if event.event_type != keyboard.KEY_DOWN:
            return

        current_mods = set()
        if win32api.GetAsyncKeyState(win32con.VK_CONTROL) < 0:
            current_mods.add("ctrl")
        if win32api.GetAsyncKeyState(win32con.VK_MENU) < 0:
            current_mods.add("alt")
        if win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0:
            current_mods.add("shift")

        key = (frozenset(current_mods), event.scan_code)
        if key in self.hotkey_actions:
            now = time.time()
            if now - self._last_action_time.get(key, 0) < 0.3:  # debounce 300ms
                return
            self._last_action_time[key] = now
            if self._action_lock.locked():
                return

            def safe_execute(mods=current_mods):
                with self._action_lock:
                    self.release_modifiers()
                    self.hotkey_actions[key]()
                    time.sleep(0.05)
                    self.restore_modifiers(mods)

            threading.Thread(target=safe_execute, daemon=True).start()

    def setup_hotkeys(self):
        keyboard.unhook_all()
        self.hotkey_actions = {}
        self.mouse_hotkeys = {}
        self.mouse_states = {}

        self.register_action("f12", self.quit_app)

        cfg = self.config.data

        for pseudo, bind_str in (cfg.get("account_focus_hotkeys") or {}).items():
            if bind_str:
                self.register_action(
                    bind_str,
                    lambda ps=pseudo: self.focus_account_by_name(ps),
                )

        try:
            if cfg.get("refresh_key"):
                self.register_action(cfg["refresh_key"], self.refresh)
            if cfg.get("auto_zaap_key"):
                self.register_action(cfg["auto_zaap_key"], self.logic.execute_auto_zaap)
            if cfg.get("sort_taskbar_key"):
                self.register_action(cfg["sort_taskbar_key"], self.logic.sort_taskbar)
            if cfg.get("invite_group_key"):
                self.register_action(
                    cfg["invite_group_key"], self.logic.execute_group_invite
                )
            if cfg.get("prev_key"):
                self.register_action(cfg["prev_key"], self.prev_char)
            if cfg.get("next_key"):
                self.register_action(cfg["next_key"], self.next_char)
            if cfg.get("leader_key"):
                self.register_action(cfg["leader_key"], self.focus_leader)
            if cfg.get("sync_key"):
                self.register_action(cfg["sync_key"], self.logic.sync_click_all)
            if cfg.get("sync_right_key"):
                self.register_action(
                    cfg["sync_right_key"], self.logic.sync_right_click_all
                )
            if cfg.get("treasure_key"):
                self.register_action(
                    cfg["treasure_key"], self.logic.execute_treasure_hunt
                )
            if cfg.get("swap_xp_drop_key"):
                self.register_action(
                    cfg["swap_xp_drop_key"], self.logic.execute_swap_xp_drop
                )
            if cfg.get("toggle_app_key"):
                self.register_action(
                    cfg["toggle_app_key"],
                    lambda: self.gui.root.after(0, self.gui.toggle_visibility),
                )
            if cfg.get("paste_enter_key"):
                self.register_action(
                    cfg["paste_enter_key"], self.logic.execute_paste_enter
                )

            keyboard.hook(self.global_hook_listener)
        except Exception as e:
            logger.error(f"Erreur de raccourci : {e}", exc_info=True)

    def refresh(self):

        slots = self.logic.scan_slots()

        self.gui.root.after(0, self.gui.refresh_list, slots)

    def focus_leader(self):
        if self.logic.leader_hwnd:
            self._switching = True
            try:
                self.logic.focus_window(self.logic.leader_hwnd)
                cycle_list = self.logic.get_cycle_list()
                leader_name = self.config.data.get("leader_name", "")
                for index, acc in enumerate(cycle_list):
                    if acc["name"] == leader_name:
                        self.current_idx = index
                        break
            finally:
                self._switching = False

    def next_char(self):
        cycle_list = self.logic.get_cycle_list()
        if not cycle_list:
            return
        self._switching = True
        try:
            self.current_idx = (self.current_idx + 1) % len(cycle_list)
            self.logic.focus_window(cycle_list[self.current_idx]["hwnd"])
        finally:
            self._switching = False

    def run_auto_trade_valider_scan(self, green_hwnd):
        """Vert détecté sur green_hwnd : cherche directement VALIDER sur les autres fenêtres et clique."""
        try:
            cycle_list = self.logic.get_cycle_list()
            if len(cycle_list) < 2:
                self.focus_leader()
                return

            clicked = False
            step_delay = float(self.config.data.get("auto_trade_step_delay", 0.5))

            for acc in cycle_list:
                hwnd_target = acc["hwnd"]
                if hwnd_target == green_hwnd:
                    continue

                # Focus direct sur la fenêtre esclave, sans touche Suivant
                self.logic.focus_window(hwnd_target)
                time.sleep(step_delay)

                if self.logic._check_valider_button_pixel_on_hwnd(hwnd_target, require_foreground=False):
                    self.logic.execute_trade_validate(
                        target_hwnd=hwnd_target, use_account_hotkey=False
                    )
                    clicked = True
                    break

            self.focus_leader()

            if not clicked:
                self.logic._notify_error(
                    "Auto-échange : bouton VALIDER non détecté. Recalibre 'Calibrer bouton VALIDER (relatif)' avec la popup d'échange ouverte."
                )
        except Exception as e:
            logger.error(f"Erreur dans run_auto_trade_valider_scan : {e}")
            try:
                self.focus_leader()
            except Exception:
                pass

    def prev_char(self):
        cycle_list = self.logic.get_cycle_list()
        if not cycle_list:
            return
        self._switching = True
        try:
            self.current_idx = (self.current_idx - 1) % len(cycle_list)
            self.logic.focus_window(cycle_list[self.current_idx]["hwnd"])
        finally:
            self._switching = False

    def quit_app(self, *args):
        logger.info("Fermeture de l'application...")
        if hasattr(self, 'tray_icon'):
            self.tray_icon.stop()
        sys.exit(0)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_as_admin():
    if getattr(sys, "frozen", False):
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv[1:]), None, 1
        )
    else:
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}" {params}', None, 1
        )


def handle_multiple_instances():
    mutex_name = "DOSTRUB_SINGLE_INSTANCE_MUTEX"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)

    if ctypes.windll.kernel32.GetLastError() == 183:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        rep = messagebox.askyesno(
            "Instance détectée",
            "Une instance de Dostrub est déjà en cours d'exécution !\n\nVoulez-vous fermer l'ancienne instance pour ouvrir celle-ci ?",
            parent=root,
        )

        if rep:
            hwnd = win32gui.FindWindow(None, "Dostrub")
            if hwnd:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    handle = ctypes.windll.kernel32.OpenProcess(1, False, pid)
                    ctypes.windll.kernel32.TerminateProcess(handle, 0)
                    ctypes.windll.kernel32.CloseHandle(handle)
                except Exception as e:
                    logger.debug(f"Erreur lors de la fermeture de l'ancienne instance : {e}")
            time.sleep(0.5)
            root.destroy()
        else:
            root.destroy()
            sys.exit(0)

    return mutex


def verify_sha256(file_path, expected_hash):
    if not expected_hash:
        return True
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest() == expected_hash

def check_and_update(app):
    try:
        # Ajout d'un paramètre date pour éviter le cache de GitHub
        response = requests.get(f"{VERSION_URL}?t={int(time.time())}", timeout=5)
        if response.status_code != 200:
            return  # Fichier version introuvable ou erreur serveur

        latest = response.text.strip()
        if latest == CURRENT_VERSION:
            return  # déjà à jour

        def _ask_and_update():
            parent = app.gui.root if app and hasattr(app, "gui") else None
            if not messagebox.askyesno(
                "Mise à jour disponible",
                f"Nouvelle version {latest} disponible !\n(Actuelle : {CURRENT_VERSION})\n\nMettre à jour maintenant ?",
                parent=parent
            ):
                return

            # Téléchargement
            exe_path = sys.executable if getattr(sys, "frozen", False) else None
            if not exe_path:
                return
            
            try:
                # Tentative de récupération du SHA256
                sha_response = requests.get(f"{DOWNLOAD_URL}.sha256", timeout=5)
                expected_sha = sha_response.text.strip().split()[0] if sha_response.status_code == 200 else None
                if not expected_sha:
                    logger.warning("Fichier SHA256 non trouvé sur le serveur. La vérification sera ignorée.")

                response_exe = requests.get(DOWNLOAD_URL, stream=True, timeout=30)
                if response_exe.status_code != 200:
                    return # Fichier exe introuvable

                tmp_path = exe_path + ".tmp"
                with open(tmp_path, "wb") as f:
                    shutil.copyfileobj(response_exe.raw, f)

                # Vérification SHA256
                if expected_sha:
                    if not verify_sha256(tmp_path, expected_sha):
                        messagebox.showerror("Erreur de sécurité", "La signature du fichier téléchargé ne correspond pas ! Mise à jour annulée.")
                        os.remove(tmp_path)
                        return
                    logger.info("Vérification SHA256 réussie.")

                # Script batch pour remplacer l'exe pendant qu'il tourne
                bat = f"""@echo off
timeout /t 2 /nobreak >nul
move /y "{tmp_path}" "{exe_path}"
start "" "{exe_path}"
"""
                bat_path = exe_path + "_update.bat"
                with open(bat_path, "w", encoding="utf-8") as f:
                    f.write(bat)

                subprocess.Popen(bat_path, shell=True)
                sys.exit(0)
            except Exception as e:
                logger.error(f"Erreur lors du téléchargement de la mise à jour : {e}")

        if app and hasattr(app, "gui") and app.gui.root:
            app.gui.root.after(0, _ask_and_update)
        else:
            _ask_and_update()

    except Exception as e:
        logger.error(f"Erreur lors de la vérification de mise à jour : {e}")


def _hide_console_on_windows():
    if sys.platform != "win32":
        return
    try:
        w = ctypes.windll.kernel32.GetConsoleWindow()
        if w:
            ctypes.windll.user32.ShowWindow(w, 0)
    except Exception:
        pass


def start_application():
    _hide_console_on_windows()

    if not is_admin():
        run_as_admin()
        sys.exit()

    try:
        _app_mutex = handle_multiple_instances()

        app = OrganizerApp()
        threading.Thread(target=check_and_update, args=(app,), daemon=True).start()
        app.gui.run()
    except Exception as e:
        logger.critical(f"Crash de l'application : {e}", exc_info=True)
        with open("crash_log.txt", "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        time.sleep(10)

if __name__ == "__main__":
    start_application()
