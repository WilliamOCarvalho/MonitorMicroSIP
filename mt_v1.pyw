import os
import subprocess
import ctypes
from ctypes import wintypes
import traceback
from datetime import datetime

# Constantes de eventos de sessão
WTS_SESSION_LOCK = 7
WTS_SESSION_UNLOCK = 8

# Tipos adicionais não disponíveis em ctypes.wintypes
HCURSOR = wintypes.HANDLE
HICON = wintypes.HANDLE
HBRUSH = wintypes.HANDLE
WNDPROCTYPE = ctypes.WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

LOGFILE = os.path.join(os.path.dirname(__file__), "monitor_microsip.log")

def log(msg):
    try:
        with open(LOGFILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} - {msg}\n")
    except Exception:
        pass  # evitar falha no log

def is_process_running(process_name):
    try:
        tasks = subprocess.check_output('tasklist', shell=True).decode('cp1252', errors='ignore')
        return process_name.lower() in tasks.lower()
    except Exception as e:
        log(f"Erro ao verificar processos: {e}")
        return False

def start_program(programa, process_name):
    if not is_process_running(process_name):
        try:
            subprocess.Popen(programa)
            log(f"Programa iniciado: {programa}")
        except Exception as e:
            log(f"Erro ao iniciar programa: {e}")
    else:
        log(f"O programa já está em execução: {process_name}")

def stop_program(process_name):
    if is_process_running(process_name):
        try:
            os.system(f'taskkill /f /im {process_name}')
            log(f"Programa encerrado: {process_name}")
        except Exception as e:
            log(f"Erro ao encerrar o programa: {e}")
    else:
        log(f"O programa já está encerrado: {process_name}")

def session_change_callback(event_type, programa, process_name):
    try:
        if event_type == WTS_SESSION_LOCK:
            log("Sessão bloqueada.")
            stop_program(process_name)
        elif event_type == WTS_SESSION_UNLOCK:
            log("Sessão desbloqueada.")
            start_program(programa, process_name)
    except Exception as e:
        log(f"Erro no callback da sessão: {e}\n{traceback.format_exc()}")

def session_change_listener(programa, process_name):
    log("Monitorando eventos de bloqueio e desbloqueio...")

    Wtsapi32 = ctypes.WinDLL("Wtsapi32.dll")
    user32 = ctypes.WinDLL("User32.dll")

    WTSRegisterSessionNotification = Wtsapi32.WTSRegisterSessionNotification
    WTSRegisterSessionNotification.argtypes = [wintypes.HWND, wintypes.DWORD]
    WTSRegisterSessionNotification.restype = wintypes.BOOL

    WTSUnRegisterSessionNotification = Wtsapi32.WTSUnRegisterSessionNotification
    WTSUnRegisterSessionNotification.argtypes = [wintypes.HWND]
    WTSUnRegisterSessionNotification.restype = wintypes.BOOL

    class WNDCLASS(ctypes.Structure):
        _fields_ = [
            ("style", wintypes.UINT),
            ("lpfnWndProc", WNDPROCTYPE),
            ("cbClsExtra", wintypes.INT),
            ("cbWndExtra", wintypes.INT),
            ("hInstance", wintypes.HINSTANCE),
            ("hIcon", HICON),
            ("hCursor", HCURSOR),
            ("hbrBackground", HBRUSH),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]

    def wnd_proc(hwnd, msg, wparam, lparam):
        try:
            if msg == 0x02B1:  # WM_WTSSESSION_CHANGE
                session_change_callback(wparam, programa, process_name)
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)
        except Exception as e:
            log(f"Erro no wnd_proc: {e}\n{traceback.format_exc()}")
            return 0

    wnd_class = WNDCLASS()
    wnd_class.lpfnWndProc = WNDPROCTYPE(wnd_proc)
    wnd_class.lpszClassName = "SessionChangeListener"
    class_atom = user32.RegisterClassW(ctypes.byref(wnd_class))

    hwnd = user32.CreateWindowExW(0, class_atom, "SessionChangeListener", 0, 0, 0, 0, 0, None, None, None, None)
    if not hwnd:
        log(f"Erro ao criar janela: {ctypes.GetLastError()}")
        return

    if not WTSRegisterSessionNotification(hwnd, 1):
        log("Erro ao registrar notificações de sessão.")
        return

    log("Monitoramento iniciado.")

    try:
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    except Exception as e:
        log(f"Erro no loop de mensagens: {e}\n{traceback.format_exc()}")
    finally:
        WTSUnRegisterSessionNotification(hwnd)
        user32.DestroyWindow(hwnd)
        user32.UnregisterClassW(class_atom, None)
        log("Monitoramento encerrado.")

if __name__ == "__main__":
    userprofile = os.environ.get("USERPROFILE")
    programa = os.path.join(userprofile, "AppData", "Local", "MicroSIP", "microsip.exe")
    process_name = os.path.basename(programa)

    if not os.path.isfile(programa):
        log(f"Erro: O programa '{programa}' não foi encontrado.")
    else:
        start_program(programa, process_name)
        session_change_listener(programa, process_name)
