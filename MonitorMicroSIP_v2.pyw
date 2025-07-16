# MonitorMicroSIP_v2.pyw

import os
import subprocess
import ctypes
from ctypes import wintypes
import logging
from logging.handlers import TimedRotatingFileHandler
import time
import win32api
import win32con
import datetime

log_dir = r"C:\MicroSIPService"
log_file = os.path.join(log_dir, "servico.log")

if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
logger.handlers = []

handler = TimedRotatingFileHandler(
    log_file,
    when='W0',
    interval=1,
    backupCount=4,
    encoding='utf-8'
)
handler.setFormatter(logging.Formatter(
    '%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
))
logger.addHandler(handler)

WTS_SESSION_LOCK = 7
WTS_SESSION_UNLOCK = 8

HCURSOR = wintypes.HANDLE
HICON = wintypes.HANDLE
HBRUSH = wintypes.HANDLE
WNDPROCTYPE = ctypes.WINFUNCTYPE(ctypes.c_long, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

def get_user_profile_path():
    profile_path = os.environ.get("USERPROFILE", "")
    if profile_path:
        logging.info(f"Caminho do perfil do usuário obtido via USERPROFILE: {profile_path}")
    else:
        logging.error("Não foi possível obter USERPROFILE do ambiente.")
    return profile_path

def is_process_running(process_name):
    try:
        tasks = subprocess.check_output('tasklist', shell=True).decode('cp1252', errors='ignore')
        running = process_name.lower() in tasks.lower()
        logging.info(f"Verificando processo {process_name}: {'em execução' if running else 'não encontrado'}")
        return running
    except Exception as e:
        logging.error(f"Erro ao verificar processos: {e}")
        return False

def start_program(programa, process_name):
    try:
        if is_process_running(process_name):
            logging.info(f"{process_name} já em execução. Finalizando antes de reiniciar.")
            stop_program(process_name)
            time.sleep(2)
        subprocess.Popen(programa, shell=False)
        logging.info(f"Programa iniciado: {programa}")
    except Exception as e:
        logging.error(f"Erro ao iniciar programa {programa}: {e}")

def stop_program(process_name):
    if is_process_running(process_name):
        try:
            os.system(f'taskkill /f /im {process_name}')
            logging.info(f"Comando de encerramento enviado para: {process_name}")
            for _ in range(10):
                time.sleep(1)
                if not is_process_running(process_name):
                    logging.info("Processo encerrado com sucesso.")
                    break
                else:
                    logging.info("Aguardando encerramento do processo...")
        except Exception as e:
            logging.error(f"Erro ao encerrar programa {process_name}: {e}")
    else:
        logging.info(f"O programa já está encerrado: {process_name}")

def session_change_callback(event_type, programa, process_name):
    if event_type == WTS_SESSION_LOCK:
        logging.info("Sessão bloqueada.")
        stop_program(process_name)
    elif event_type == WTS_SESSION_UNLOCK:
        logging.info("Sessão desbloqueada.")
        start_program(programa, process_name)

def session_change_listener(programa, process_name):
    logging.info("Iniciando monitoramento de eventos de sessão...")

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
            if msg == 0x02B1:
                session_change_callback(wparam, programa, process_name)
            return user32.DefWindowProcW(
                wintypes.HWND(hwnd),
                wintypes.UINT(msg),
                wintypes.WPARAM(wparam),
                wintypes.LPARAM(lparam & 0xFFFFFFFFFFFFFFFF)
            )
        except Exception as e:
            logging.error(f"Erro no wnd_proc: {e}")
            return 0

    wnd_class = WNDCLASS()
    wnd_class.lpfnWndProc = WNDPROCTYPE(wnd_proc)
    wnd_class.lpszClassName = "SessionChangeListener"
    class_atom = user32.RegisterClassW(ctypes.byref(wnd_class))
    if not class_atom:
        logging.error(f"Erro ao registrar classe de janela: {ctypes.GetLastError()}")
        return False

    hwnd = user32.CreateWindowExW(0, class_atom, "SessionChangeListener", 0, 0, 0, 0, 0, None, None, None, None)
    if not hwnd:
        logging.error(f"Erro ao criar janela: {ctypes.GetLastError()}")
        return False

    if not WTSRegisterSessionNotification(hwnd, 1):
        logging.error(f"Erro ao registrar notificações de sessão: {ctypes.GetLastError()}")
        user32.DestroyWindow(hwnd)
        user32.UnregisterClassW(class_atom, None)
        return False

    logging.info("Monitoramento de sessão iniciado com sucesso.")

    try:
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    except KeyboardInterrupt:
        logging.info("Monitoramento interrompido pelo usuário.")
    except Exception as e:
        logging.error(f"Erro no loop de mensagens: {e}")
    finally:
        WTSUnRegisterSessionNotification(hwnd)
        user32.DestroyWindow(hwnd)
        user32.UnregisterClassW(class_atom, None)
        logging.info("Monitoramento de sessão encerrado.")
    return True

def main():
    profile_path = get_user_profile_path()
    programa = os.path.join(profile_path, "AppData", "Local", "MicroSIP", "microsip.exe")
    process_name = os.path.basename(programa)

    if not os.path.isfile(programa):
        logging.error(f"Erro: O programa '{programa}' não foi encontrado.")
        return

    logging.info(f"Iniciando serviço para gerenciar: {programa}")
    start_program(programa, process_name)

    success = session_change_listener(programa, process_name)
    if not success:
        logging.warning("Monitoramento de sessão falhou. Mantendo serviço ativo com verificação periódica.")
        while True:
            try:
                if not is_process_running(process_name):
                    logging.info("Programa não está em execução. Tentando reiniciar.")
                    start_program(programa, process_name)
                time.sleep(60)
            except KeyboardInterrupt:
                logging.info("Serviço interrompido pelo usuário.")
                break
            except Exception as e:
                logging.error(f"Erro no loop de verificação: {e}")
                time.sleep(60)

if __name__ == "__main__":
    try:
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        main()
    except Exception as e:
        logging.error(f"Erro fatal no serviço: {e}")
