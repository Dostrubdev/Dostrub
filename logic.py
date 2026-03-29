import ctypes
import time
import threading
import logging
import os
import re
from ctypes import POINTER
import win32gui
import win32con
import win32api
import win32process
import keyboard
import pytesseract
from PIL import ImageGrab, ImageFilter, ImageEnhance
from constants import AZERTY_TO_SCAN

logger = logging.getLogger("Dostrub.Logic")

class BlockInputContext:
    """Conserve l'accès exclusif aux périphériques et force le déblocage après timeout."""
    def __init__(self, timeout=5.0):
        self.timeout = timeout
        self._stop_event = threading.Event()
        self._watchdog_thread = None

    def __enter__(self):
        ctypes.windll.user32.BlockInput(True)
        self._stop_event.clear()
        self._watchdog_thread = threading.Thread(target=self._watchdog, daemon=True)
        self._watchdog_thread.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._stop_event.set()
        ctypes.windll.user32.BlockInput(False)

    def _watchdog(self):
        if not self._stop_event.wait(self.timeout):
            logger.warning(f"Watchdog BlockInput déclenché après {self.timeout}s - déblocage forcé.")
            ctypes.windll.user32.BlockInput(False)




class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", POINTER(ctypes.c_ulong)),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", POINTER(ctypes.c_ulong)),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.c_ulong),
        ("wParamL", ctypes.c_ushort),
        ("wParamH", ctypes.c_ushort),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _anonymous_ = ("_input",)
    _fields_ = [("type", ctypes.c_ulong), ("_input", _INPUT_UNION)]


def screen_to_norm(x, y):
    sw = ctypes.windll.user32.GetSystemMetrics(78)
    sh = ctypes.windll.user32.GetSystemMetrics(79)
    ox = ctypes.windll.user32.GetSystemMetrics(76)
    oy = ctypes.windll.user32.GetSystemMetrics(77)
    return int((x - ox) * 65535 / sw), int((y - oy) * 65535 / sh)


class DofusLogic:
    def __init__(self, config):
        self.config = config
        self.all_accounts = []
        self.active_mode = "ALL"
        self.leader_hwnd = None
        self.last_sync_time = 0
        self.db = None
        self.scanner_running = True
        # Point 3 : flag anti-conflit groupe/échange
        self._group_invite_in_progress = False
        # Point 4 : nom OCR du destinataire en attente
        self._pending_trade_target = None
        # Auto-échange (callbacks depuis OrganizerApp)
        self._trade_focus_leader_fn = None
        self._trade_run_valider_scan_fn = None  # (green_hwnd) — enchaîne les vrais « Suivant »
        threading.Thread(target=self._trade_scanner_loop, daemon=True).start()
        self.error_callback = None

    def set_error_callback(self, callback):
        self.error_callback = callback

    def _notify_error(self, msg):
        if self.error_callback:
            self.error_callback(msg)

    def scan_slots(self):
        windows_trouvees = []

        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                if win32gui.GetClassName(hwnd) == "UnityWndClass":
                    titre = win32gui.GetWindowText(hwnd)
                    if titre.strip():
                        windows_trouvees.append((hwnd, titre))
            return True

        win32gui.EnumWindows(enum_windows_callback, None)

        nouveaux_comptes = []
        for hwnd, titre in windows_trouvees:
            titre_clean = titre.strip()
            if titre_clean.lower().startswith("dofus") or not titre_clean:
                continue
            parts = titre_clean.split(" - ")
            pseudo = parts[0].strip()
            classe = parts[1].strip() if len(parts) > 1 else "Inconnu"
            self.config.data["classes"][pseudo] = classe
            etat_actif = self.config.data["accounts_state"].get(pseudo, True)
            equipe = self.config.data["accounts_team"].get(pseudo, "Team 1")
            nouveaux_comptes.append(
                {
                    "name": pseudo,
                    "hwnd": hwnd,
                    "active": etat_actif,
                    "team": equipe,
                    "classe": classe,
                }
            )

        custom_order = self.config.data.get("custom_order", [])
        for acc in nouveaux_comptes:
            if acc["name"] not in custom_order:
                custom_order.append(acc["name"])

        if len(custom_order) > 50:
            active_names = [acc["name"] for acc in nouveaux_comptes]
            inactive = [n for n in custom_order if n not in active_names]
            while len(custom_order) > 50 and inactive:
                to_remove = inactive.pop(0)
                if to_remove in custom_order:
                    custom_order.remove(to_remove)

        self.config.data["custom_order"] = custom_order
        self.config.save()
        self.all_accounts = sorted(
            nouveaux_comptes, key=lambda x: custom_order.index(x["name"])
        )

        self.leader_hwnd = None
        leader_name = self.config.data.get("leader_name", "")
        for acc in self.all_accounts:
            if acc["name"] == leader_name:
                self.leader_hwnd = acc["hwnd"]

        return self.all_accounts

    def get_cycle_list(self):
        mode = self.config.data.get("current_mode", "ALL")
        valid_accounts = []
        for acc in self.all_accounts:
            if (
                win32gui.IsWindow(acc["hwnd"])
                and acc["active"]
                and (mode == "ALL" or acc["team"] == mode)
            ):
                valid_accounts.append(acc)
        return valid_accounts

    def _update_global_order_from_active(self, active_accs):
        order = self.config.data.get("custom_order", [])
        indices = []
        valid_names = []
        for acc in active_accs:
            if acc["name"] in order:
                indices.append(order.index(acc["name"]))
                valid_names.append(acc["name"])
        indices.sort()
        for i, name in zip(indices, valid_names):
            order[i] = name
        self.config.data["custom_order"] = order
        self.config.save()
        self.all_accounts.sort(key=lambda x: order.index(x["name"]))

    def set_account_position(self, name, new_index):
        active_accs = self.get_cycle_list()
        names = [a["name"] for a in active_accs]
        if name not in names:
            return
        idx = names.index(name)
        acc_to_move = active_accs.pop(idx)
        active_accs.insert(new_index, acc_to_move)
        self._update_global_order_from_active(active_accs)

    def move_account(self, name, direction):
        active_accs = self.get_cycle_list()
        names = [a["name"] for a in active_accs]
        if name not in names:
            return
        idx = names.index(name)
        new_idx = idx + direction
        if 0 <= new_idx < len(names):
            active_accs[idx], active_accs[new_idx] = (
                active_accs[new_idx],
                active_accs[idx],
            )
            self._update_global_order_from_active(active_accs)

    def toggle_account(self, name, is_active):
        for acc in self.all_accounts:
            if acc["name"] == name:
                acc["active"] = is_active
        self.config.data["accounts_state"][name] = is_active
        self.config.save()

    def change_team(self, name, new_team):
        for acc in self.all_accounts:
            if acc["name"] == name:
                acc["team"] = new_team
        self.config.data["accounts_team"][name] = new_team
        self.config.save()

    def set_mode(self, mode):
        self.config.data["current_mode"] = mode
        self.config.save()

    def set_leader(self, name):
        self.leader_hwnd = None
        self.config.data["leader_name"] = name
        self.config.save()
        for acc in self.all_accounts:
            if acc["name"] == name:
                self.leader_hwnd = acc["hwnd"]

    def close_account_window(self, name):
        for acc in self.all_accounts:
            if acc["name"] == name:
                hwnd = acc["hwnd"]
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    handle = ctypes.windll.kernel32.OpenProcess(1, False, pid)
                    ctypes.windll.kernel32.TerminateProcess(handle, 0)
                    ctypes.windll.kernel32.CloseHandle(handle)
                except Exception:
                    pass
                break

    def close_all_active_accounts(self):
        active_accs = self.get_cycle_list()
        for acc in active_accs:
            hwnd = acc["hwnd"]
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                handle = ctypes.windll.kernel32.OpenProcess(1, False, pid)
                ctypes.windll.kernel32.TerminateProcess(handle, 0)
                ctypes.windll.kernel32.CloseHandle(handle)
            except:
                pass

    def focus_window(self, hwnd):
        if not hwnd or not win32gui.IsWindow(hwnd):
            return
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            fg = win32gui.GetForegroundWindow()
            if fg == hwnd:
                return
            fg_tid, _ = win32process.GetWindowThreadProcessId(fg)
            to_tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            cur_tid = ctypes.windll.kernel32.GetCurrentThreadId()
            ctypes.windll.user32.AllowSetForegroundWindow(pid)
            if fg_tid and to_tid:
                ctypes.windll.user32.AttachThreadInput(cur_tid, fg_tid, True)
                ctypes.windll.user32.AttachThreadInput(cur_tid, to_tid, True)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
            if fg_tid and to_tid:
                ctypes.windll.user32.AttachThreadInput(cur_tid, to_tid, False)
                ctypes.windll.user32.AttachThreadInput(cur_tid, fg_tid, False)
            for _ in range(25):
                if win32gui.GetForegroundWindow() == hwnd:
                    break
                time.sleep(0.04)
        except Exception:
            pass

    def get_first_non_leader_hwnd(self):
        leader_name = self.config.data.get("leader_name")
        for acc in self.get_cycle_list():
            if acc["name"] != leader_name:
                return acc["hwnd"]
        return None


    def get_account_name_by_hwnd(self, hwnd):
        for acc in self.get_cycle_list():
            if acc["hwnd"] == hwnd:
                return acc["name"]
        return "Unknown"

    def get_relative_ratio_pos(self, hwnd=None):
        x_screen, y_screen = win32gui.GetCursorPos()
        target_hwnd = hwnd if hwnd else self.leader_hwnd
        if target_hwnd:
            try:
                client_pt = win32gui.ScreenToClient(target_hwnd, (x_screen, y_screen))
                rect = win32gui.GetClientRect(target_hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w > 0 and h > 0:
                    return (client_pt[0] / float(w), client_pt[1] / float(h))
            except:
                pass
        return (0.0, 0.0)

    def get_screen_coords_from_saved(self, hwnd, saved_pos):
        if not saved_pos or len(saved_pos) != 2:
            return None
        x, y = saved_pos

        if isinstance(x, float) and isinstance(y, float) and x <= 1.0 and y <= 1.0:
            try:
                rect = win32gui.GetClientRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w == 0 or h == 0:
                    return None
                client_x = int(x * w)
                client_y = int(y * h)
                return win32gui.ClientToScreen(hwnd, (client_x, client_y))
            except:
                return None
        else:
            try:
                return win32gui.ClientToScreen(hwnd, (int(x), int(y)))
            except:
                return None

    def get_client_coords_from_saved(self, hwnd, saved_pos):
        """Convertit un point calibré en coordonnées CLIENT (relatives à la fenêtre)."""
        if not saved_pos or len(saved_pos) != 2 or not hwnd:
            return None
        x, y = saved_pos
        try:
            rect = win32gui.GetClientRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            if w <= 0 or h <= 0:
                return None
            if isinstance(x, float) and isinstance(y, float) and x <= 1.0 and y <= 1.0:
                return int(x * w), int(y * h)
            return int(x), int(y)
        except Exception:
            return None

    def _hardware_key(self, scan_code):
        i_down = INPUT()
        i_down.type = 1
        i_down.ki.wScan = scan_code
        i_down.ki.dwFlags = 0x0008

        i_up = INPUT()
        i_up.type = 1
        i_up.ki.wScan = scan_code
        i_up.ki.dwFlags = 0x0008 | 0x0002

        ctypes.windll.user32.SendInput(1, ctypes.byref(i_down), ctypes.sizeof(INPUT))
        time.sleep(0.01)
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_up), ctypes.sizeof(INPUT))

    def _hardware_click(self, x, y):
        nx, ny = screen_to_norm(x, y)
        i_move = INPUT()
        i_move.type = 0
        i_move.mi.dx = nx
        i_move.mi.dy = ny
        i_move.mi.dwFlags = 0x0001 | 0x8000 | 0x4000
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_move), ctypes.sizeof(INPUT))
        time.sleep(0.03)

        i_down = INPUT()
        i_down.type = 0
        i_down.mi.dx = nx
        i_down.mi.dy = ny
        i_down.mi.dwFlags = 0x0002 | 0x8000 | 0x4000
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_down), ctypes.sizeof(INPUT))
        time.sleep(0.04)

        i_up = INPUT()
        i_up.type = 0
        i_up.mi.dx = nx
        i_up.mi.dy = ny
        i_up.mi.dwFlags = 0x0004 | 0x8000 | 0x4000
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_up), ctypes.sizeof(INPUT))

    def _fast_hardware_click(self, x, y):
        nx, ny = screen_to_norm(x, y)
        i_move = INPUT()
        i_move.type = 0
        i_move.mi.dx = nx
        i_move.mi.dy = ny
        i_move.mi.dwFlags = 0x0001 | 0x8000 | 0x4000
        i_down = INPUT()
        i_down.type = 0
        i_down.mi.dx = nx
        i_down.mi.dy = ny
        i_down.mi.dwFlags = 0x0002 | 0x8000 | 0x4000
        i_up = INPUT()
        i_up.type = 0
        i_up.mi.dx = nx
        i_up.mi.dy = ny
        i_up.mi.dwFlags = 0x0004 | 0x8000 | 0x4000

        speed = self.config.data.get("click_speed", "Rapide")

        delay_press = (
            0.01 if speed == "Rapide" else (0.04 if speed == "Moyen" else 0.06)
        )
        delay_after = 0.0 if speed == "Rapide" else (0.09 if speed == "Moyen" else 0.18)

        ctypes.windll.user32.SendInput(1, ctypes.byref(i_move), ctypes.sizeof(INPUT))
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_down), ctypes.sizeof(INPUT))
        time.sleep(delay_press)
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_up), ctypes.sizeof(INPUT))

        if delay_after > 0:
            time.sleep(delay_after)

    def _fast_hardware_right_click(self, x, y):
        nx, ny = screen_to_norm(x, y)
        i_move = INPUT()
        i_move.type = 0
        i_move.mi.dx = nx
        i_move.mi.dy = ny
        i_move.mi.dwFlags = 0x0001 | 0x8000 | 0x4000
        i_down = INPUT()
        i_down.type = 0
        i_down.mi.dx = nx
        i_down.mi.dy = ny
        i_down.mi.dwFlags = 0x0008 | 0x8000 | 0x4000
        i_up = INPUT()
        i_up.type = 0
        i_up.mi.dx = nx
        i_up.mi.dy = ny
        i_up.mi.dwFlags = 0x0010 | 0x8000 | 0x4000

        speed = self.config.data.get("click_speed", "Rapide")

        delay_press = (
            0.01 if speed == "Rapide" else (0.04 if speed == "Moyen" else 0.06)
        )
        delay_after = 0.0 if speed == "Rapide" else (0.09 if speed == "Moyen" else 0.18)

        ctypes.windll.user32.SendInput(1, ctypes.byref(i_move), ctypes.sizeof(INPUT))
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_down), ctypes.sizeof(INPUT))
        time.sleep(delay_press)
        ctypes.windll.user32.SendInput(1, ctypes.byref(i_up), ctypes.sizeof(INPUT))

        if delay_after > 0:
            time.sleep(delay_after)

    def broadcast_key(self, key_name):
        time.sleep(0.1)
        active_accs = self.get_cycle_list()
        if not active_accs:
            return
        with BlockInputContext():
            current_hwnd = win32gui.GetForegroundWindow()
            scan_code = AZERTY_TO_SCAN.get(key_name.lower())
            for acc in active_accs:
                self.focus_window(acc["hwnd"])
                time.sleep(0.1)
                if scan_code:
                    self._hardware_key(scan_code)
                else:
                    keyboard.send(key_name)
                time.sleep(0.02)
            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(current_hwnd)

    def execute_paste_enter(self):
        time.sleep(0.1)
        active_accs = self.get_cycle_list()
        if not active_accs:
            return
        original_fg_hwnd = win32gui.GetForegroundWindow()
        time.sleep(0.15)
        with BlockInputContext():
            for acc in active_accs:
                self.focus_window(acc["hwnd"])
                time.sleep(0.1)
                keyboard.send("ctrl+v")
                time.sleep(0.02)
                keyboard.send("enter")
                time.sleep(0.02)
            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(original_fg_hwnd)

    def execute_auto_zaap(self):
        active_accs = self.get_cycle_list()
        zaaps_pos = self.config.data["macro_positions"].get("zaaps", {})
        haven_key = self.config.data.get("game_haven_key", "h")

        if not active_accs:
            return
        for acc in active_accs:
            if acc["name"] not in zaaps_pos:
                self._notify_error(f"Votre Zaap ({acc['name']}) n'est pas calibré.")
                return

        original_fg_hwnd = win32gui.GetForegroundWindow()
        time.sleep(0.15)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")

        haven_scan = AZERTY_TO_SCAN.get(haven_key.lower())

        try:
            for acc in active_accs:
                self.focus_window(acc["hwnd"])
                time.sleep(0.15)
                if haven_scan:
                    self._hardware_key(haven_scan)
                else:
                    keyboard.send(haven_key)
                time.sleep(0.02)

            try:
                delai = float(self.config.data.get("zaap_delay", "1.8"))
            except ValueError:
                delai = 1.8

            time.sleep(delai)

            click_order = (
                active_accs[1:] + [active_accs[0]]
                if len(active_accs) > 1
                else active_accs
            )

            for index, acc in enumerate(click_order):
                coords = self.get_screen_coords_from_saved(
                    acc["hwnd"], zaaps_pos.get(acc["name"])
                )
                if not coords:
                    continue
                x_c, y_c = coords

                self.focus_window(acc["hwnd"])
                time.sleep(0.15)
                win32api.SetCursorPos((x_c, y_c))
                time.sleep(0.05)

                self._hardware_click(x_c, y_c)
                time.sleep(0.05)

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(original_fg_hwnd)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def sync_click_all(self):
        active_accs = self.get_cycle_list()
        if not active_accs:
            return
        current_x, current_y = win32api.GetCursorPos()
        hwnd_under_mouse = win32gui.WindowFromPoint((current_x, current_y))
        root_hwnd = win32gui.GetAncestor(hwnd_under_mouse, win32con.GA_ROOT)

        is_dofus = any(acc["hwnd"] == root_hwnd for acc in self.all_accounts)
        reference_hwnd = (
            root_hwnd
            if is_dofus
            else (self.leader_hwnd if self.leader_hwnd else active_accs[0]["hwnd"])
        )

        try:
            rel_x, rel_y = win32gui.ScreenToClient(
                reference_hwnd, (current_x, current_y)
            )
            ref_rect = win32gui.GetClientRect(reference_hwnd)
            ref_w, ref_h = ref_rect[2] - ref_rect[0], ref_rect[3] - ref_rect[1]
            if ref_w == 0 or ref_h == 0:
                return
            ratio_x, ratio_y = rel_x / float(ref_w), rel_y / float(ref_h)
        except:
            return

        original_fg_hwnd = win32gui.GetForegroundWindow()
        time.sleep(0.15)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")

        try:
            for acc in active_accs:
                hwnd = acc["hwnd"]
                try:
                    target_rect = win32gui.GetClientRect(hwnd)
                    t_w, t_h = (
                        target_rect[2] - target_rect[0],
                        target_rect[3] - target_rect[1],
                    )
                    if t_w == 0 or t_h == 0:
                        continue
                    client_x, client_y = int(ratio_x * t_w), int(ratio_y * t_h)
                    target_x, target_y = win32gui.ClientToScreen(
                        hwnd, (client_x, client_y)
                    )

                    self.focus_window(hwnd)
                    win32api.SetCursorPos((target_x, target_y))
                    self._fast_hardware_click(target_x, target_y)
                except:
                    pass

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(original_fg_hwnd)
            try:
                win32api.SetCursorPos((current_x, current_y))
            except:
                pass
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def sync_right_click_all(self):
        active_accs = self.get_cycle_list()
        if not active_accs:
            return
        current_x, current_y = win32api.GetCursorPos()
        hwnd_under_mouse = win32gui.WindowFromPoint((current_x, current_y))
        root_hwnd = win32gui.GetAncestor(hwnd_under_mouse, win32con.GA_ROOT)

        is_dofus = any(acc["hwnd"] == root_hwnd for acc in self.all_accounts)
        reference_hwnd = (
            root_hwnd
            if is_dofus
            else (self.leader_hwnd if self.leader_hwnd else active_accs[0]["hwnd"])
        )

        try:
            rel_x, rel_y = win32gui.ScreenToClient(
                reference_hwnd, (current_x, current_y)
            )
            ref_rect = win32gui.GetClientRect(reference_hwnd)
            ref_w, ref_h = ref_rect[2] - ref_rect[0], ref_rect[3] - ref_rect[1]
            if ref_w == 0 or ref_h == 0:
                return
            ratio_x, ratio_y = rel_x / float(ref_w), rel_y / float(ref_h)
        except:
            return

        original_fg_hwnd = win32gui.GetForegroundWindow()
        time.sleep(0.15)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")

        try:
            for acc in active_accs:
                hwnd = acc["hwnd"]
                try:
                    target_rect = win32gui.GetClientRect(hwnd)
                    t_w, t_h = (
                        target_rect[2] - target_rect[0],
                        target_rect[3] - target_rect[1],
                    )
                    if t_w == 0 or t_h == 0:
                        continue
                    client_x, client_y = int(ratio_x * t_w), int(ratio_y * t_h)
                    target_x, target_y = win32gui.ClientToScreen(
                        hwnd, (client_x, client_y)
                    )

                    self.focus_window(hwnd)
                    win32api.SetCursorPos((target_x, target_y))
                    self._fast_hardware_right_click(target_x, target_y)
                except:
                    pass

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(original_fg_hwnd)
            try:
                win32api.SetCursorPos((current_x, current_y))
            except:
                pass
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def execute_group_invite(self):
        leader = self.config.data.get("leader_name")
        chat_pos = self.config.data["macro_positions"].get("chat_position")

        if not self.leader_hwnd or not leader:
            self._notify_error("Décidez d'un chef pour inviter !")
            return
        if not chat_pos:
            self._notify_error("Votre Chat n'est pas calibré.")
            return

        coords = self.get_screen_coords_from_saved(self.leader_hwnd, chat_pos)
        if not coords:
            return
        x_c, y_c = coords

        active_accs = self.get_cycle_list()
        time.sleep(0.1)
        self.focus_window(self.leader_hwnd)
        time.sleep(0.2)

        # Point 3 : bloquer le scanner échange pendant l'invitation de groupe
        self._group_invite_in_progress = True
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        try:
            win32api.SetCursorPos((x_c, y_c))
            time.sleep(0.05)
            self._hardware_click(x_c, y_c)
            time.sleep(0.15)

            keyboard.send("ctrl+a")
            time.sleep(0.05)
            keyboard.send("backspace")
            time.sleep(0.1)

            for acc in active_accs:
                if acc["name"] == leader:
                    continue
                keyboard.write(f"/invite {acc['name']}")
                time.sleep(0.1)
                keyboard.send("enter")
                time.sleep(0.1)

            if self.config.data.get("auto_group_enabled", False):
                # Point 1 : délai suffisant pour que la pop-up apparaisse sur chaque compte
                time.sleep(1.5)
                group_pos = self.config.data.get("macro_positions", {}).get("group_accept_pos")
                for acc in active_accs:
                    if acc["name"] == leader:
                        continue
                    self.focus_window(acc["hwnd"])
                    time.sleep(0.25)
                    if group_pos:
                        coords_g = self.get_screen_coords_from_saved(acc["hwnd"], group_pos)
                        if coords_g:
                            sx, sy = coords_g
                            win32api.SetCursorPos((sx, sy))
                            self._hardware_click(sx, sy)
                    else:
                        keyboard.send("enter")
                    # Point 1 : attendre que le jeu traite l'action avant de passer au suivant
                    time.sleep(0.4)

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)

        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass
            # Point 3 : libérer le verrou après la fin de l'invitation
            self._group_invite_in_progress = False

    def execute_trade_accept(self, target_name=None):
        """Accepte un échange. Si target_name est fourni (Point 4), va directement
        sur la fenêtre du personnage concerné. Sinon fallback sur tous les comptes."""
        active_accs = self.get_cycle_list()
        if not active_accs:
            return

        current_hwnd = win32gui.GetForegroundWindow()
        leader = self.config.data.get("leader_name")

        # Point 4 : ciblage direct si on connaît le nom
        if target_name:
            target_acc = None
            for acc in active_accs:
                if acc["name"].lower() == target_name.lower():
                    target_acc = acc
                    break
            if target_acc:
                try:
                    ctypes.windll.user32.BlockInput(True)
                except:
                    pass
                try:
                    self.focus_window(target_acc["hwnd"])
                    time.sleep(0.2)
                    keyboard.send("enter")
                    time.sleep(0.1)
                    if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                        self.focus_window(self.leader_hwnd)
                    else:
                        self.focus_window(current_hwnd)
                except:
                    pass
                finally:
                    try:
                        ctypes.windll.user32.BlockInput(False)
                    except:
                        pass
                return
            # Si le nom OCR n'a pas matché, fallback sur tous les comptes

        # Fallback : boucle sur tous les comptes (comportement original)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        try:
            for acc in active_accs:
                if acc["name"] == leader:
                    continue
                self.focus_window(acc["hwnd"])
                time.sleep(0.15)
                keyboard.send("enter")
                time.sleep(0.05)

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
            else:
                self.focus_window(current_hwnd)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def _ocr_extract_sender_from_notif(self, sx, sy):
        try:
            import pytesseract
            import os
            from PIL import ImageGrab, ImageEnhance
            import re

            default_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ]
            for p in default_paths:
                if os.path.exists(p):
                    pytesseract.pytesseract.tesseract_cmd = p
                    break

            x1 = max(0, sx - 350)
            y1 = max(0, sy - 40)
            x2 = sx + 50
            y2 = sy + 40
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2))

            img = img.resize((img.width * 2, img.height * 2))
            img = ImageEnhance.Contrast(img).enhance(2.5)
            img = img.convert("L")

            custom_cfg = r"--oem 3 --psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789- "
            text = pytesseract.image_to_string(img, config=custom_cfg).strip()

            patterns = [
                r"^([A-Za-z0-9-]+)\s+souhaite",
                r"^([A-Za-z0-9-]+)\s+wants to trade",
                r"^([A-Za-z0-9-]+)\s+m",
            ]
            for pat in patterns:
                m = re.search(pat, text)
                if m:
                    return m.group(1)
            return None
        except Exception:
            return None

    def _check_green_pixel_on_hwnd(self, hwnd):
        try:
            from PIL import ImageGrab
            validate_pos = self.config.data.get("macro_positions", {}).get("trade_validate_pos")
            if not validate_pos or len(validate_pos) < 5 or not hwnd:
                return False
            if not win32gui.IsWindow(hwnd):
                return False
            rx, ry, exp_r, exp_g, exp_b = validate_pos
            client_coords = self.get_client_coords_from_saved(hwnd, [rx, ry])
            if not client_coords:
                return False
            cx, cy = client_coords
            sx, sy = win32gui.ClientToScreen(hwnd, (cx, cy))
            img = ImageGrab.grab(bbox=(sx, sy, sx + 1, sy + 1))
            r, g, b = img.getpixel((0, 0))[:3]
            tol = int(self.config.data.get("auto_trade_color_tolerance", 45))
            return (
                abs(r - exp_r) < tol
                and abs(g - exp_g) < tol
                and abs(b - exp_b) < tol
            )
        except Exception:
            return False

    def _check_valider_button_pixel_on_hwnd(self, hwnd, require_foreground=True):
        try:
            from PIL import ImageGrab

            tcp = self.config.data.get("macro_positions", {}).get("trade_validate_click_pos")
            if not tcp or len(tcp) < 5 or not hwnd:
                return False
            if not win32gui.IsWindow(hwnd):
                return False
            if require_foreground and win32gui.GetForegroundWindow() != hwnd:
                return False
            rx, ry, exp_r, exp_g, exp_b = tcp[0], tcp[1], tcp[2], tcp[3], tcp[4]
            coords = self.get_screen_coords_from_saved(hwnd, [rx, ry])
            if not coords:
                return False
            sx, sy = coords
            img = ImageGrab.grab(bbox=(sx, sy, sx + 1, sy + 1))
            r, g, b = img.getpixel((0, 0))[:3]
            tol = int(self.config.data.get("auto_trade_color_tolerance", 48))
            return (
                abs(r - exp_r) < tol
                and abs(g - exp_g) < tol
                and abs(b - exp_b) < tol
            )
        except Exception:
            return False

    def _run_trade_valider_window_scan(self, green_hwnd):
        if self._trade_run_valider_scan_fn:
            self._trade_run_valider_scan_fn(green_hwnd)
            return

        cycle_accs = self.get_cycle_list()
        n = len(cycle_accs)
        if n < 2:
            return

        start_idx = 0
        for i, acc in enumerate(cycle_accs):
            if acc["hwnd"] == green_hwnd:
                start_idx = i
                break

        clicked = False
        for step in range(1, n):
            idx = (start_idx + step) % n
            hwnd = cycle_accs[idx]["hwnd"]
            if hwnd == green_hwnd:
                continue
            self.focus_window(hwnd)
            time.sleep(0.2)
            if self._check_valider_button_pixel_on_hwnd(hwnd):
                self.execute_trade_validate(target_hwnd=hwnd, use_account_hotkey=False)
                clicked = True
                break

        if self._trade_focus_leader_fn:
            self._trade_focus_leader_fn()
        elif self.leader_hwnd:
            self.focus_window(self.leader_hwnd)

        if not clicked and len(
            self.config.data.get("macro_positions", {}).get("trade_validate_click_pos") or []
        ) >= 5:
            self._notify_error(
                "Auto-échange : aucune fenêtre avec le bouton Valider détecté. "
                "Recalibre « VALIDER » sur le centre du bouton (échange ouvert)."
            )

    def execute_trade_validate(self, target_hwnd=None, use_account_hotkey=True):
        mp = self.config.data.get("macro_positions", {})
        click_rel = mp.get("trade_validate_click_pos")
        legacy = mp.get("trade_validate_pos")

        if click_rel and len(click_rel) >= 2:
            saved = [click_rel[0], click_rel[1]]
        elif legacy and len(legacy) >= 2:
            saved = [legacy[0], legacy[1]]
        else:
            return

        hwnd = target_hwnd
        if not hwnd:
            hwnd = self.get_first_non_leader_hwnd() or self.leader_hwnd
        if not hwnd:
            return

        partner_name = self.get_account_name_by_hwnd(hwnd)
        key_str = (self.config.data.get("account_focus_hotkeys") or {}).get(partner_name, "")
        key_str = (key_str or "").strip()

        try:
            coords = self.get_screen_coords_from_saved(hwnd, saved)
            if not coords:
                return
            sx, sy = coords
        except Exception:
            return

        try:
            if use_account_hotkey:
                if key_str:
                    try:
                        keyboard.send(key_str)
                    except Exception:
                        self.focus_window(hwnd)
                    time.sleep(0.2)
                else:
                    self.focus_window(hwnd)
                    time.sleep(0.2)
            else:
                self.focus_window(hwnd)
                time.sleep(0.16)
            with BlockInputContext():
                win32api.SetCursorPos((sx, sy))
                time.sleep(0.06)
                self._hardware_click(sx, sy)
        except Exception:
            pass

    def _trade_scanner_loop(self):
        from PIL import ImageGrab
        import threading
        _in_trade_session = False
        _session_start = 0.0
        SESSION_TIMEOUT = 60.0

        while self.scanner_running:
            try:
                if not self.config.data.get("auto_trade_enabled", False):
                    _in_trade_session = False
                    time.sleep(1.0)
                    continue

                notif_pos = self.config.data.get("macro_positions", {}).get("trade_notif_pos")
                if not notif_pos or len(notif_pos) < 5:
                    time.sleep(1.0)
                    continue

                if not self.leader_hwnd:
                    time.sleep(1.0)
                    continue

                if self._group_invite_in_progress:
                    time.sleep(0.5)
                    continue

                if _in_trade_session:
                    if time.time() - _session_start > SESSION_TIMEOUT:
                        _in_trade_session = False
                        time.sleep(1.0)
                        continue

                    if not self.leader_hwnd:
                        _in_trade_session = False
                        continue

                    if self._check_green_pixel_on_hwnd(self.leader_hwnd):
                        self._run_trade_valider_window_scan(self.leader_hwnd)
                        _in_trade_session = False
                        time.sleep(1.5)
                    else:
                        time.sleep(0.3)

                    continue

                rx, ry, exp_r, exp_g, exp_b = notif_pos
                coords = self.get_screen_coords_from_saved(self.leader_hwnd, [rx, ry])
                if not coords:
                    time.sleep(1.0)
                    continue

                sx, sy = coords
                img = ImageGrab.grab(bbox=(sx, sy, sx + 1, sy + 1))
                r, g, b = img.getpixel((0, 0))[:3]

                pixel1_ok = abs(r - exp_r) < 20 and abs(g - exp_g) < 20 and abs(b - exp_b) < 20

                if pixel1_ok:
                    notif_pos2 = self.config.data.get("macro_positions", {}).get("trade_notif_pos2")
                    pixel2_ok = True
                    if notif_pos2 and len(notif_pos2) >= 5:
                        rx2, ry2, exp_r2, exp_g2, exp_b2 = notif_pos2
                        coords2 = self.get_screen_coords_from_saved(self.leader_hwnd, [rx2, ry2])
                        if coords2:
                            sx2, sy2 = coords2
                            img2 = ImageGrab.grab(bbox=(sx2, sy2, sx2 + 1, sy2 + 1))
                            r2, g2, b2 = img2.getpixel((0, 0))[:3]
                            pixel2_ok = abs(r2 - exp_r2) < 20 and abs(g2 - exp_g2) < 20 and abs(b2 - exp_b2) < 20

                    if pixel2_ok:
                        sender_name = self._ocr_extract_sender_from_notif(sx, sy)
                        self._pending_trade_target = sender_name

                        threading.Thread(
                            target=self.execute_trade_accept,
                            args=(sender_name,),
                            daemon=True
                        ).start()

                        _in_trade_session = True
                        _session_start = time.time()
                        time.sleep(1.5)  # laisser le temps à la fenêtre d'échange de s'ouvrir
                        continue

            except Exception:
                pass

            # ── Veille standard : 1 scan par seconde ────────────────
            time.sleep(1.0)


    def execute_treasure_hunt(self):
        chat_pos = self.config.data["macro_positions"].get("chat_position")
        if not chat_pos:
            self._notify_error("Votre Chat n'est pas calibré.")
            return

        current_hwnd = win32gui.GetForegroundWindow()
        known_hwnds = [acc["hwnd"] for acc in self.all_accounts]
        target_hwnd = current_hwnd
        if target_hwnd not in known_hwnds:
            if self.leader_hwnd:
                target_hwnd = self.leader_hwnd
                self.focus_window(target_hwnd)
                time.sleep(0.2)
            else:
                return

        coords = self.get_screen_coords_from_saved(target_hwnd, chat_pos)
        if not coords:
            return
        x_c, y_c = coords

        time.sleep(0.1)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        try:
            win32api.SetCursorPos((x_c, y_c))
            time.sleep(0.05)
            self._hardware_click(x_c, y_c)
            time.sleep(0.15)

            keyboard.send("ctrl+v")
            time.sleep(0.05)
            keyboard.send("enter")
            time.sleep(0.20)
            keyboard.send("enter")
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def execute_swap_xp_drop(self):
        pos = self.config.data["macro_positions"].get("xp_drop_button")
        if not pos:
            self._notify_error("Votre XP/Drop n'est pas calibré.")
            return

        active_accs = self.get_cycle_list()
        if not active_accs:
            return

        time.sleep(0.1)
        try:
            ctypes.windll.user32.BlockInput(True)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        try:
            for acc in active_accs:
                hwnd = acc["hwnd"]
                coords = self.get_screen_coords_from_saved(hwnd, pos)
                if not coords:
                    continue
                x_c, y_c = coords

                self.focus_window(hwnd)
                time.sleep(0.2)
                win32api.SetCursorPos((x_c, y_c))
                time.sleep(0.03)
                self._hardware_click(x_c, y_c)

            if self.config.data.get("return_to_leader", True) and self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
        except Exception as e:
            logger.debug(f"Erreur ignorée dans logic (except nu) : {e}")
        finally:
            try:
                ctypes.windll.user32.BlockInput(False)
            except:
                pass

    def sort_taskbar(self):
        active_accs = self.get_cycle_list()
        if not active_accs:
            return
        try:
            for acc in active_accs:
                win32gui.ShowWindow(acc["hwnd"], win32con.SW_HIDE)
            time.sleep(0.2)
            for acc in active_accs:
                win32gui.ShowWindow(acc["hwnd"], win32con.SW_SHOW)
                time.sleep(0.05)
            if self.leader_hwnd:
                self.focus_window(self.leader_hwnd)
        except Exception:
            pass

