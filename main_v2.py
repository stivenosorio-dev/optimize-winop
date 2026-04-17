"""
WinOptimizer Pro — Sistema de Mantenimiento y Optimización de Windows
Requiere ejecución como Administrador.
Compatible con Windows 10/11.
"""

import os
import sys
import subprocess
import winreg
import ctypes
import threading
import shutil
import glob
import time
import platform
from tkinter import (
    Tk, ttk, BooleanVar, StringVar, IntVar,
    Frame, Label, Button, Canvas, Scrollbar,
    messagebox
)
import tkinter as tk


# ─────────────────────────────────────────────
#  CONSTANTES Y PALETA VISUAL
# ─────────────────────────────────────────────
APP_TITLE    = "WinOptimizer Pro"
APP_W, APP_H = 1150, 800

CLR = {
    "bg":        "#0D0F14",
    "panel":     "#131720",
    "card":      "#1A1F2E",
    "card2":     "#161B28",
    "border":    "#252D3D",
    "accent":    "#4F8EF7",
    "accent2":   "#7B5EF8",
    "green":     "#3DDC84",
    "green_dim": "#1A6040",
    "yellow":    "#F7C948",
    "yellow_dim":"#5C4A10",
    "red":       "#F75A5A",
    "red_dim":   "#5C1A1A",
    "text":      "#E8EAF0",
    "muted":     "#6B7280",
    "highlight": "#1E2A40",
    "progress_bg": "#0D1520",
    "progress_fill": "#4F8EF7",
}

# Estados de ejecución de una sección
ST_IDLE     = "idle"       # Sin ejecutar
ST_RUNNING  = "running"    # En ejecución
ST_DONE     = "done"       # Completado con éxito
ST_ERROR    = "error"      # Terminado con errores
ST_PARTIAL  = "partial"    # Completado con algunos errores

FONT_TITLE   = ("Segoe UI", 20, "bold")
FONT_SUB     = ("Segoe UI", 11)
FONT_LABEL   = ("Segoe UI", 10)
FONT_SMALL   = ("Segoe UI", 9)
FONT_MONO    = ("Consolas", 9)
FONT_BTN     = ("Segoe UI Semibold", 10)
FONT_BADGE   = ("Segoe UI", 8, "bold")
FONT_STATUS  = ("Segoe UI Semibold", 9)


# ─────────────────────────────────────────────
#  UTILIDADES DEL SISTEMA
# ─────────────────────────────────────────────

def is_admin() -> bool:
    """Verifica si el proceso corre con privilegios de administrador."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_cmd(cmd, shell: bool = True) -> tuple[int, str]:
    """Ejecuta un comando y retorna (código_retorno, salida)."""
    try:
        result = subprocess.run(
            cmd, shell=shell, capture_output=True, text=True,
            timeout=120, encoding="utf-8", errors="ignore"
        )
        return result.returncode, (result.stdout + result.stderr).strip()
    except subprocess.TimeoutExpired:
        return -1, "Tiempo de espera agotado."
    except Exception as e:
        return -1, str(e)


def run_ps(script: str) -> tuple[int, str]:
    """Ejecuta un bloque de PowerShell y retorna (código, salida)."""
    cmd = ["powershell", "-NoProfile", "-NonInteractive",
           "-ExecutionPolicy", "Bypass", "-Command", script]
    return run_cmd(cmd, shell=False)


def get_system_info() -> dict:
    """Recopila especificaciones relevantes del sistema."""
    info = {}

    _, out = run_ps(
        "Get-CimInstance Win32_ComputerSystem | "
        "Select-Object -ExpandProperty TotalPhysicalMemory"
    )
    try:
        info["ram_mb"] = int(out.strip()) // (1024 * 1024)
    except Exception:
        info["ram_mb"] = 8192

    _, out = run_ps(
        "Get-PhysicalDisk | Sort-Object MediaType | "
        "Select-Object -First 1 -ExpandProperty DeviceID"
    )
    info["fast_disk"] = out.strip() or "C"

    _, out = run_ps(
        "Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name"
    )
    info["cpu"] = out.strip() or platform.processor()
    info["win_release"] = platform.release()
    info["win_ver"] = platform.version()
    return info


# ─────────────────────────────────────────────
#  LÓGICA DE MANTENIMIENTO ESTÁNDAR
# ─────────────────────────────────────────────

def get_startup_apps() -> list[dict]:
    """
    Lista apps de inicio de terceros ordenadas por impacto.
    Excluye controladores GPU, audio, Defender y servicios del SO.
    """
    EXCLUDED = {
        "nvidia", "amd", "radeon", "geforce", "realtek", "nahimic",
        "audiodg", "defender", "windowssecurity", "securityhealth",
        "msmpeng", "windefend", "sgrmbroker",
    }
    apps = []
    hives = [
        (winreg.HKEY_CURRENT_USER,
         r"Software\Microsoft\Windows\CurrentVersion\Run"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"Software\Microsoft\Windows\CurrentVersion\Run"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"),
    ]
    for hive, path in hives:
        try:
            key = winreg.OpenKey(hive, path)
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, i)
                    if not any(kw in name.lower().replace(" ", "") for kw in EXCLUDED):
                        apps.append({
                            "name": name, "cmd": value,
                            "hive": hive, "path": path,
                            "impact": _estimate_impact(name.lower()),
                        })
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception:
            continue
    apps.sort(key=lambda x: {"Alto": 0, "Medio": 1, "Bajo": 2}[x["impact"]])
    return apps


def _estimate_impact(name: str) -> str:
    """Heurística para catalogar impacto de inicio."""
    if any(k in name for k in {"update", "cloud", "onedrive", "dropbox",
                                 "googledrive", "teams", "discord", "slack",
                                 "zoom", "spotify", "steam"}):
        return "Alto"
    if any(k in name for k in {"chat", "messenger", "telegram", "skype",
                                 "epic", "uplay", "origin", "battle", "sync"}):
        return "Medio"
    return "Bajo"


def disable_startup_app(app: dict) -> bool:
    """Elimina la entrada de registro de inicio para la app indicada."""
    try:
        key = winreg.OpenKey(app["hive"], app["path"], 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, app["name"])
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def get_nonessential_services() -> list[dict]:
    """Lista servicios no esenciales en inicio automático."""
    TARGETS = {
        "Spooler":          "Cola de Impresión",
        "DiagTrack":        "Telemetría de Windows (DiagTrack)",
        "dmwappushservice": "Telemetría WAP Push",
        "SysMain":          "SuperFetch / SysMain",
        "WSearch":          "Windows Search (indexado)",
        "Fax":              "Servicio de Fax",
        "MapsBroker":       "Mapas descargados",
        "XblAuthManager":   "Xbox Live Auth Manager",
        "XblGameSave":      "Xbox Live Game Save",
        "XboxNetApiSvc":    "Xbox Live Networking",
        "RetailDemo":       "Modo demostración minorista",
        "wercplsupport":    "Informes de errores de Windows",
        "WerSvc":           "Servicio de informes de errores",
    }
    services = []
    for svc_name, friendly in TARGETS.items():
        ps = (
            f"$s = Get-Service -Name '{svc_name}' -ErrorAction SilentlyContinue; "
            f"if ($s) {{ $s.StartType }}"
        )
        _, out = run_ps(ps)
        if out.strip() in ("Automatic", "AutomaticDelayedStart"):
            services.append({
                "name": svc_name,
                "friendly": friendly,
                "start_type": out.strip(),
            })
    return services


def set_service_manual(service_name: str) -> bool:
    code, _ = run_ps(f"Set-Service -Name '{service_name}' -StartupType Manual")
    return code == 0


def perform_cleanup(tasks: list[str], step_cb, log_cb) -> dict[str, bool]:
    """
    Ejecuta tareas de limpieza.
    step_cb(current, total, label): actualiza la barra de progreso.
    log_cb(msg, ok): emite una línea al log con indicador de éxito.
    Retorna {task_id: bool} con el resultado de cada tarea.
    """
    TASK_LABELS = {
        "temp_files":           "Eliminando archivos temporales",
        "windows_update_cache": "Limpiando caché de actualizaciones",
        "windows_old":          "Eliminando carpeta Windows.old",
        "recycle_bin":          "Vaciando Papelera de Reciclaje",
        "event_logs":           "Limpiando registros de eventos",
        "cleanmgr":             "Ejecutando Liberador de espacio en disco",
        "prefetch":             "Limpiando Prefetch",
    }
    results = {}
    total = len(tasks)

    for idx, task in enumerate(tasks):
        label = TASK_LABELS.get(task, task)
        step_cb(idx, total, label)

        ok = True
        try:
            if task == "temp_files":
                _delete_glob(os.environ.get("TEMP", ""), "*")
                _delete_glob(
                    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Temp"), "*"
                )

            elif task == "windows_update_cache":
                run_cmd("net stop wuauserv", shell=True)
                _delete_glob(r"C:\Windows\SoftwareDistribution\Download", "*")
                run_cmd("net start wuauserv", shell=True)

            elif task == "windows_old":
                old = r"C:\Windows.old"
                if os.path.exists(old):
                    run_ps(
                        f"Remove-Item -Recurse -Force '{old}' "
                        "-ErrorAction SilentlyContinue"
                    )
                else:
                    log_cb("Windows.old no encontrado (ya fue limpiado)", True)
                    results[task] = True
                    continue

            elif task == "recycle_bin":
                run_ps("Clear-RecycleBin -Force -ErrorAction SilentlyContinue")

            elif task == "event_logs":
                run_ps(
                    "Get-WinEvent -ListLog * | ForEach-Object { "
                    "[System.Diagnostics.Eventing.Reader.EventLogSession]::"
                    "GlobalSession.ClearLog($_.LogName) }"
                )

            elif task == "cleanmgr":
                run_cmd(
                    r'reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion'
                    r'\Explorer\VolumeCaches\Update Cleanup" '
                    r"/v StateFlags0064 /t REG_DWORD /d 2 /f",
                    shell=True,
                )
                run_cmd("cleanmgr /sagerun:64", shell=True)

            elif task == "prefetch":
                _delete_glob(r"C:\Windows\Prefetch", "*.pf")

        except Exception as e:
            ok = False
            log_cb(f"{label}: error — {e}", False)

        results[task] = ok
        if ok:
            log_cb(f"{label}: completado", True)

    step_cb(total, total, "Limpieza finalizada")
    return results


def _delete_glob(directory: str, pattern: str) -> None:
    """Elimina archivos/carpetas que coincidan con el patrón."""
    if not directory or not os.path.exists(directory):
        return
    for item in glob.glob(os.path.join(directory, pattern)):
        try:
            if os.path.isfile(item) or os.path.islink(item):
                os.remove(item)
            elif os.path.isdir(item):
                shutil.rmtree(item, ignore_errors=True)
        except Exception:
            pass


# ─────────────────────────────────────────────
#  LÓGICA DE OPTIMIZACIÓN AVANZADA
# ─────────────────────────────────────────────

def create_restore_point() -> tuple[bool, str]:
    """Crea un punto de restauración antes de cambios críticos."""
    ps = (
        "Enable-ComputerRestore -Drive 'C:\\' -ErrorAction SilentlyContinue; "
        "Checkpoint-Computer -Description 'WinOptimizer Pro — Pre-Optimización' "
        "-RestorePointType 'MODIFY_SETTINGS'"
    )
    code, out = run_ps(ps)
    return code == 0, out


def check_bios_recommendations(sys_info: dict) -> list[dict]:
    """Genera recomendaciones de BIOS/UEFI basadas en el hardware detectado."""
    recs = []

    _, ram_speed = run_ps(
        "Get-CimInstance Win32_PhysicalMemory | "
        "Select-Object -ExpandProperty Speed | Select-Object -First 1"
    )
    recs.append({
        "id":     "xmp_docp",
        "title":  "Activar perfil XMP / DOCP",
        "desc":   (
            f"RAM: {sys_info.get('ram_mb','?')} MB | "
            f"Velocidad actual: {ram_speed.strip() or '?'} MHz. "
            "Habilitar XMP/DOCP en BIOS para operar a la velocidad anunciada."
        ),
        "guide":  "BIOS → AI Tweaker / Extreme Tweaker → XMP/DOCP → Perfil 1",
    })

    _, rebar = run_ps(
        "Get-ItemProperty 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers' "
        "-Name PcieLargePageEnabled -ErrorAction SilentlyContinue | "
        "Select-Object -ExpandProperty PcieLargePageEnabled"
    )
    recs.append({
        "id":     "resizable_bar",
        "title":  "Activar Resizable BAR (ReBAR)",
        "desc":   (
            "Mejora acceso de la CPU a la VRAM. "
            f"Estado actual: {rebar.strip() or 'No configurado'}."
        ),
        "guide":  "BIOS → Advanced PCIe → Resizable BAR / Above 4G Decoding → Enabled",
    })

    _, sb  = run_ps("Confirm-SecureBootUEFI -ErrorAction SilentlyContinue")
    _, tpm = run_ps(
        "(Get-WmiObject -Namespace 'root/cimv2/Security/MicrosoftTpm' "
        "-Class Win32_Tpm).IsEnabled_InitialValue"
    )
    recs.append({
        "id":     "secure_boot_tpm",
        "title":  "Verificar Secure Boot y TPM 2.0",
        "desc":   (
            f"Secure Boot: {sb.strip() or 'Desconocido'} | "
            f"TPM: {tpm.strip() or 'Desconocido'}. Requeridos por Windows 11."
        ),
        "guide":  "BIOS → Security → Secure Boot → Enabled | TPM Device → Firmware TPM",
    })
    return recs


def run_ddu_guide() -> str:
    return (
        "1. Descarga DDU desde guru3d.com/files/2/display-driver-uninstaller.html\n"
        "2. Reinicia en Modo Seguro: Configuración → Recuperación → Inicio avanzado\n"
        "3. Ejecuta DDU → selecciona GPU/fabricante → 'Clean and restart'\n"
        "4. Instala el controlador limpio desde el sitio oficial:\n"
        "   • NVIDIA : nvidia.com/drivers\n"
        "   • AMD    : amd.com/support\n"
        "   • Intel  : intel.com/content/www/us/en/download-center"
    )


def get_outdated_drivers() -> list[dict]:
    """Detecta drivers instalados y verifica actualizaciones disponibles vía Windows Update."""
    import json as _json

    # 1. Obtener drivers instalados
    ps_installed = r"""
    Get-CimInstance Win32_PnPSignedDriver |
      Where-Object { $_.DriverVersion -ne $null -and $_.DeviceName -ne $null } |
      Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer, DeviceClass |
      ConvertTo-Json -Compress
    """
    _, installed_json = run_ps(ps_installed)
    installed = []
    try:
        parsed = _json.loads(installed_json.strip())
        if isinstance(parsed, dict):
            parsed = [parsed]
        installed = parsed
    except Exception:
        pass

    # 2. Buscar actualizaciones de drivers vía Windows Update
    ps_updates = r"""
    try {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        $Result = $Searcher.Search("IsInstalled=0 and Type='Driver'")
        $updates = @()
        foreach ($u in $Result.Updates) {
            $updates += @{
                Title = $u.Title
                DriverModel = if ($u.DriverModel) { $u.DriverModel } else { '' }
                DriverVerDate = if ($u.DriverVerDate) { $u.DriverVerDate.ToString() } else { '' }
            }
        }
        $updates | ConvertTo-Json -Compress
    } catch {
        Write-Output '[]'
    }
    """
    _, updates_json = run_ps(ps_updates)
    available_updates = []
    try:
        parsed_upd = _json.loads(updates_json.strip())
        if isinstance(parsed_upd, dict):
            parsed_upd = [parsed_upd]
        available_updates = parsed_upd
    except Exception:
        pass

    # 3. Marcar drivers con actualizaciones disponibles
    update_titles = {u.get("Title", "").lower() for u in available_updates}
    drivers = []
    seen = set()
    for drv in installed:
        name = drv.get("DeviceName", "")
        if not name or name in seen:
            continue
        seen.add(name)
        has_update = any(
            name.lower() in title or
            drv.get("Manufacturer", "").lower() in title
            for title in update_titles
        )
        drivers.append({
            "name": name,
            "version": drv.get("DriverVersion", "?"),
            "date": str(drv.get("DriverDate", "?"))[:10],
            "manufacturer": drv.get("Manufacturer", "Desconocido"),
            "class": drv.get("DeviceClass", ""),
            "has_update": has_update,
        })

    # Ordenar: los que tienen update primero
    drivers.sort(key=lambda d: (0 if d["has_update"] else 1, d["name"]))

    return drivers, available_updates


def update_all_drivers(step_cb=None, log_cb=None) -> tuple[bool, str]:
    """Instala todas las actualizaciones de drivers disponibles vía Windows Update (oficiales)."""
    ps = r"""
    try {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        $Result = $Searcher.Search("IsInstalled=0 and Type='Driver'")
        if ($Result.Updates.Count -eq 0) {
            Write-Output 'NO_UPDATES'
            exit 0
        }
        Write-Output "FOUND:$($Result.Updates.Count)"
        $Downloader = $Session.CreateUpdateDownloader()
        $Downloader.Updates = $Result.Updates
        Write-Output 'DOWNLOADING'
        $DlResult = $Downloader.Download()
        Write-Output 'INSTALLING'
        $Installer = $Session.CreateUpdateInstaller()
        $Installer.Updates = $Result.Updates
        $InstResult = $Installer.Install()
        $success = 0
        $failed = 0
        for ($i = 0; $i -lt $Result.Updates.Count; $i++) {
            $title = $Result.Updates.Item($i).Title
            $code = $InstResult.GetUpdateResult($i).ResultCode
            if ($code -eq 2) {
                Write-Output "OK:$title"
                $success++
            } else {
                Write-Output "FAIL:$title"
                $failed++
            }
        }
        Write-Output "DONE:$success/$($success+$failed)"
        if ($InstResult.RebootRequired) {
            Write-Output 'REBOOT_REQUIRED'
        }
    } catch {
        Write-Output "ERROR:$($_.Exception.Message)"
    }
    """
    code, out = run_ps(ps)
    return code == 0, out.strip()


def apply_telemetry_registry() -> tuple[bool, str]:
    """Deshabilita telemetría de Windows mediante claves de registro."""
    ps = r"""
    $p = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
    New-Item -Path $p -Force | Out-Null
    Set-ItemProperty -Path $p -Name AllowTelemetry                   -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $p -Name DoNotShowFeedbackNotifications   -Value 1 -Type DWord -Force
    Set-Service -Name DiagTrack        -StartupType Disabled -ErrorAction SilentlyContinue
    Stop-Service -Name DiagTrack       -Force -ErrorAction SilentlyContinue
    Set-Service -Name dmwappushservice -StartupType Disabled -ErrorAction SilentlyContinue
    Stop-Service -Name dmwappushservice -Force -ErrorAction SilentlyContinue
    Write-Output 'OK'
    """
    code, out = run_ps(ps)
    return code == 0, out.strip()


def configure_pagefile(ram_mb: int, target_drive: str = "C:") -> tuple[bool, str]:
    """Configura pagefile de tamaño fijo recomendado según la RAM."""
    size_mb = int(min(max(ram_mb * 1.5, 4096), 16384))
    ps = f"""
    $cs = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
    $cs.AutomaticManagedPagefile = $False
    $cs.Put() | Out-Null
    $pf = Get-WmiObject -Query "SELECT * FROM Win32_PageFileSetting WHERE Name LIKE '{target_drive}%'"
    if ($pf) {{
        $pf.InitialSize = {size_mb}; $pf.MaximumSize = {size_mb}; $pf.Put() | Out-Null
    }} else {{
        Set-WmiInstance -Class Win32_PageFileSetting -Arguments @{{
            Name='{target_drive}\\pagefile.sys';
            InitialSize={size_mb}; MaximumSize={size_mb}
        }} | Out-Null
    }}
    Write-Output '{size_mb} MB configurados en {target_drive}'
    """
    code, out = run_ps(ps)
    return code == 0, f"{out.strip()} (requiere reinicio)"


def exclude_search_indexer(paths: list[str]) -> tuple[bool, str]:
    """Agrega exclusiones al indexador de Windows Search vía registro."""
    ps_lines = [
        "$base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows Search"
        "\\CrawlScopeManager\\Windows\\SystemIndex\\DefaultRules'"
    ]
    for i, p in enumerate(paths, 1):
        ps_lines.append(
            f"Set-ItemProperty -Path $base -Name 'ExcludedPath{i}' "
            f"-Value '{p}' -Type String -Force -ErrorAction SilentlyContinue"
        )
    ps_lines.append("Write-Output 'OK'")
    code, out = run_ps("\n".join(ps_lines))
    return code == 0, out.strip()


def apply_power_plan_ultimate() -> tuple[bool, str]:
    """Activa SCHEME_MIN y fija CPU al 100% mínimo y máximo."""
    ps = r"""
    powercfg -setactive SCHEME_MIN 2>$null
    if ($LASTEXITCODE -ne 0) {
        powercfg -duplicatescheme SCHEME_MIN 2>$null
        powercfg -setactive SCHEME_MIN 2>$null
    }
    $guid = (powercfg -getactivescheme).Split()[3]
    powercfg -setacvalueindex $guid SUB_PROCESSOR PROCTHROTTLEMIN 100
    powercfg -setacvalueindex $guid SUB_PROCESSOR PROCTHROTTLEMAX 100
    powercfg -setactive $guid
    Write-Output "Plan Máximo Rendimiento activado — CPU 100%"
    """
    code, out = run_ps(ps)
    return code == 0, out.strip()


def apply_visual_performance() -> tuple[bool, str]:
    """
    Ajusta efectos visuales a 'Mejor rendimiento', conservando
    sombras de ventanas, ClearType y miniaturas.
    """
    ps = r"""
    $vfx = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'
    Set-ItemProperty -Path $vfx -Name VisualFXSetting -Value 2 -Type DWord -Force
    $mask = [byte[]](0x90,0x12,0x00,0x80,0x10,0x00,0x00,0x00)
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name UserPreferencesMask -Value $mask -Type Binary -Force
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' `
        -Name ListviewShadow -Value 1 -Type DWord -Force
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\WindowMetrics' `
        -Name MinAnimate -Value 0 -Type String -Force
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name FontSmoothing     -Value 2 -Type String -Force
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name FontSmoothingType -Value 2 -Type DWord -Force
    Write-Output "Efectos visuales optimizados"
    """
    code, out = run_ps(ps)
    return code == 0, out.strip()


# ─────────────────────────────────────────────
#  WIDGET: PANEL DE PROGRESO POR SECCIÓN
# ─────────────────────────────────────────────

class ProgressPanel(Frame):
    """
    Panel de estado y progreso que se muestra en cada sección.
    Muestra: estado global, barra de progreso, tarea actual y log de pasos.
    """

    # Configuración visual por estado
    _STATE_CFG = {
        ST_IDLE:    {"icon": "○", "label": "Listo para ejecutar",   "color": CLR["muted"],  "bar": CLR["border"]},
        ST_RUNNING: {"icon": "◉", "label": "Ejecutando...",         "color": CLR["accent"], "bar": CLR["accent"]},
        ST_DONE:    {"icon": "✓", "label": "Completado con éxito",  "color": CLR["green"],  "bar": CLR["green"]},
        ST_ERROR:   {"icon": "✗", "label": "Error en la ejecución", "color": CLR["red"],    "bar": CLR["red"]},
        ST_PARTIAL: {"icon": "⚠", "label": "Completado con avisos", "color": CLR["yellow"], "bar": CLR["yellow"]},
    }

    def __init__(self, parent, section_name: str, **kwargs):
        super().__init__(parent, bg=CLR["card2"], **kwargs)
        self._section = section_name
        self._state   = ST_IDLE
        self._steps_ok    = 0
        self._steps_error = 0
        self._bar_value   = 0.0   # 0.0 – 1.0
        self._anim_id     = None  # ID de animación de la barra indeterminada
        self._anim_pos    = 0.0

        self._build()

    def _build(self):
        """Construye el widget internamente."""
        # ── Fila superior: icono de estado + texto + contador ──
        top = Frame(self, bg=CLR["card2"], padx=12, pady=8)
        top.pack(fill="x")

        self._icon_lbl = Label(top, text="○", font=("Segoe UI", 14),
                               bg=CLR["card2"], fg=CLR["muted"])
        self._icon_lbl.pack(side="left")

        mid = Frame(top, bg=CLR["card2"])
        mid.pack(side="left", padx=10, fill="x", expand=True)

        self._state_lbl = Label(mid, text="Listo para ejecutar",
                                font=FONT_STATUS, bg=CLR["card2"], fg=CLR["muted"])
        self._state_lbl.pack(anchor="w")

        self._task_lbl = Label(mid, text="",
                               font=FONT_SMALL, bg=CLR["card2"], fg=CLR["muted"])
        self._task_lbl.pack(anchor="w")

        self._counter_lbl = Label(top, text="",
                                  font=FONT_MONO, bg=CLR["card2"], fg=CLR["muted"])
        self._counter_lbl.pack(side="right")

        # ── Barra de progreso dibujada en Canvas ──
        bar_frame = Frame(self, bg=CLR["card2"], padx=12)
        bar_frame.pack(pady=(0, 4))
        bar_frame.pack(fill="x")

        self._bar_canvas = Canvas(bar_frame, height=6, bg=CLR["progress_bg"],
                                  highlightthickness=0, bd=0)
        self._bar_canvas.pack(fill="x")
        self._bar_rect_bg   = None
        self._bar_rect_fill = None
        self._bar_canvas.bind("<Configure>", self._redraw_bar)

        # ── Log de pasos ──
        self._log_frame = Frame(self, bg=CLR["card2"], padx=12)
        # Se mostrará dinámicamente en start_run()

        self._log_text = tk.Text(
            self._log_frame, height=4, bg=CLR["bg"], fg=CLR["text"],
            font=FONT_MONO, relief="flat", bd=0, wrap="word",
            state="disabled", padx=6, pady=4,
        )
        log_scroll = Scrollbar(self._log_frame, command=self._log_text.yview,
                               bg=CLR["panel"])
        self._log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")
        self._log_text.pack(fill="x")

        # Tags de color en el log
        self._log_text.tag_configure("ok",      foreground=CLR["green"])
        self._log_text.tag_configure("error",   foreground=CLR["red"])
        self._log_text.tag_configure("warn",    foreground=CLR["yellow"])
        self._log_text.tag_configure("info",    foreground=CLR["accent"])
        self._log_text.tag_configure("default", foreground=CLR["text"])

    def _redraw_bar(self, event=None):
        """Redibuja la barra de progreso al cambiar el tamaño."""
        w = self._bar_canvas.winfo_width()
        h = self._bar_canvas.winfo_height()
        if w < 2:
            return
        cfg   = self._STATE_CFG[self._state]
        fill  = max(2, int(w * self._bar_value))
        color = cfg["bar"]

        self._bar_canvas.delete("all")
        # Fondo
        self._bar_canvas.create_rectangle(0, 0, w, h, fill=CLR["progress_bg"],
                                          outline="", tags="bg")
        # Relleno
        if self._bar_value > 0 or self._state == ST_RUNNING:
            self._bar_canvas.create_rectangle(0, 0, fill, h,
                                              fill=color, outline="",
                                              tags="fill")

    def set_state(self, state: str):
        """Cambia el estado visual del panel."""
        self._state = state
        cfg = self._STATE_CFG[state]
        self._icon_lbl.configure(text=cfg["icon"], fg=cfg["color"])
        self._state_lbl.configure(text=cfg["label"], fg=cfg["color"])
        if state != ST_RUNNING:
            self._stop_indeterminate()
            self._task_lbl.configure(text="")
        self._redraw_bar()

    def set_progress(self, value: float, task_label: str = ""):
        """
        Actualiza la barra de progreso.
        value: 0.0 – 1.0. Si es -1 activa modo indeterminado.
        """
        if value < 0:
            self._start_indeterminate()
        else:
            self._stop_indeterminate()
            self._bar_value = max(0.0, min(1.0, value))
            self._redraw_bar()

        if task_label:
            self._task_lbl.configure(text=f"  {task_label}")

    def _start_indeterminate(self):
        """Animación de barra indeterminada (pulso de izquierda a derecha)."""
        if self._anim_id:
            return
        self._anim_pos = 0.0
        self._animate_indeterminate()

    def _animate_indeterminate(self):
        """Un tick de la animación indeterminada."""
        self._anim_pos = (self._anim_pos + 0.02) % 1.2
        w = self._bar_canvas.winfo_width()
        if w < 2:
            self._anim_id = self._bar_canvas.after(40, self._animate_indeterminate)
            return

        seg_w = int(w * 0.35)
        x0 = int((self._anim_pos - 0.35) * w)
        x1 = x0 + seg_w

        self._bar_canvas.delete("all")
        self._bar_canvas.create_rectangle(0, 0, w, 6, fill=CLR["progress_bg"],
                                          outline="")
        if x1 > 0 and x0 < w:
            self._bar_canvas.create_rectangle(
                max(0, x0), 0, min(w, x1), 6,
                fill=CLR["accent"], outline=""
            )
        self._anim_id = self._bar_canvas.after(30, self._animate_indeterminate)

    def _stop_indeterminate(self):
        """Detiene la animación indeterminada."""
        if self._anim_id:
            self._bar_canvas.after_cancel(self._anim_id)
            self._anim_id = None

    def update_counter(self, ok: int, errors: int, total: int):
        """Actualiza el contador de pasos completados."""
        self._steps_ok    = ok
        self._steps_error = errors
        parts = []
        if total > 0:
            parts.append(f"{ok + errors}/{total}")
        if ok > 0:
            parts.append(f"✓{ok}")
        if errors > 0:
            parts.append(f"✗{errors}")
        self._counter_lbl.configure(
            text="  ".join(parts),
            fg=CLR["red"] if errors > 0 else CLR["green"] if ok > 0 else CLR["muted"]
        )

    def log(self, msg: str, tag: str = "default"):
        """Añade una línea al log del panel con timestamp."""
        ts = time.strftime("%H:%M:%S")
        self._log_text.configure(state="normal")
        self._log_text.insert("end", f"[{ts}] ", "default")
        self._log_text.insert("end", f"{msg}\n", tag)
        self._log_text.configure(state="disabled")
        self._log_text.see("end")

    def reset(self):
        """Restablece el panel al estado inicial."""
        self._stop_indeterminate()
        self._bar_value = 0.0
        self._steps_ok = self._steps_error = 0
        self._counter_lbl.configure(text="")
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")
        self._log_frame.pack_forget()
        self.set_state(ST_IDLE)

    def start_run(self):
        """Prepara el panel para iniciar una ejecución."""
        self.reset()
        self._log_frame.pack(fill="x", pady=(2, 8))
        self.set_state(ST_RUNNING)
        self.set_progress(-1)  # Animación indeterminada al inicio


# ─────────────────────────────────────────────
#  INTERFAZ GRÁFICA PRINCIPAL
# ─────────────────────────────────────────────

class WinOptimizerApp:
    """Ventana principal de la aplicación."""

    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{APP_W}x{APP_H}")
        self.root.configure(bg=CLR["bg"])
        self.root.resizable(True, True)
        self.root.minsize(950, 700)

        self.sys_info: dict = {}
        self.startup_apps:           list[dict] = []
        self.nonessential_services:  list[dict] = []

        # Indica si una sección está corriendo (evita doble-click)
        self._running = False

        self._build_ui()
        self._load_data_async()

    # ── Construcción de la UI ──────────────────

    def _build_ui(self):
        self._build_sidebar()
        self._build_content_area()

    def _build_sidebar(self):
        sidebar = Frame(self.root, bg=CLR["panel"], width=230)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Logo
        hdr = Frame(sidebar, bg=CLR["panel"], pady=22)
        hdr.pack(fill="x")
        Label(hdr, text="⚡", font=("Segoe UI", 30),
              bg=CLR["panel"], fg=CLR["accent"]).pack()
        Label(hdr, text="WinOptimizer", font=("Segoe UI Semibold", 13),
              bg=CLR["panel"], fg=CLR["text"]).pack()
        Label(hdr, text="Pro", font=("Segoe UI", 10),
              bg=CLR["panel"], fg=CLR["accent2"]).pack()

        ttk.Separator(sidebar, orient="horizontal").pack(fill="x", padx=16, pady=6)

        # Navegación
        self._nav_btns: dict[str, Button] = {}
        self._nav_status: dict[str, Label] = {}   # indicador de estado por sección
        sections = [
            ("🚀", "Inicio Automático",    "startup"),
            ("⚙️", "Servicios",            "services"),
            ("🧹", "Limpieza",             "cleanup"),
            ("🔧", "BIOS / Controladores", "bios"),
            ("✨", "Soft Optimización",    "soft_optim"),
            ("🛡️", "Suprimir Telemetría", "telemetry"),
            ("📄", "Archivo de Paginación","pagefile"),
            ("⚡", "Plan de Energía",      "power"),
            ("🎨", "Efectos Visuales",     "visual"),
        ]
        self.current_section = StringVar(value="startup")
        for icon, label, key in sections:
            self._make_nav_btn(sidebar, icon, label, key)

        ttk.Separator(sidebar, orient="horizontal").pack(fill="x", padx=16, pady=6)

        # Info del sistema
        self._sys_label = Label(
            sidebar, text="Analizando sistema...",
            font=FONT_MONO, bg=CLR["panel"], fg=CLR["muted"],
            wraplength=210, justify="left", padx=14, pady=6
        )
        self._sys_label.pack(side="bottom", fill="x")

    def _make_nav_btn(self, parent, icon: str, label: str, key: str):
        """Crea un botón de navegación con su indicador de estado."""
        container = Frame(parent, bg=CLR["panel"])
        container.pack(fill="x", padx=6, pady=1)

        btn = Button(
            container, text=f"  {icon}  {label}",
            font=FONT_BTN, anchor="w",
            bg=CLR["panel"], fg=CLR["text"],
            activebackground=CLR["highlight"], activeforeground=CLR["accent"],
            relief="flat", bd=0, padx=10, pady=9, cursor="hand2",
            command=lambda k=key: self._switch_section(k)
        )
        btn.pack(side="left", fill="x", expand=True)

        # Pequeño círculo de estado (○ idle, ◉ running, ✓ done, ✗ error)
        status = Label(container, text="", font=("Segoe UI", 9),
                       bg=CLR["panel"], fg=CLR["muted"], width=2)
        status.pack(side="right", padx=(0, 8))

        self._nav_btns[key]    = btn
        self._nav_status[key]  = status

    def _update_nav_state(self, key: str, state: str):
        """Actualiza el indicador de estado del botón lateral."""
        cfg = ProgressPanel._STATE_CFG.get(state, {})
        lbl = self._nav_status.get(key)
        if lbl:
            lbl.configure(
                text=cfg.get("icon", ""),
                fg=cfg.get("color", CLR["muted"])
            )

    def _switch_section(self, key: str):
        """Cambia la sección activa."""
        self.current_section.set(key)
        # Actualizar encabezado
        TITLES = {
            "startup":    ("Inicio Automático",     "Gestión de programas de arranque"),
            "services":   ("Servicios del Sistema", "Servicios no esenciales"),
            "cleanup":    ("Limpieza del Sistema",  "Archivos temporales y residuos"),
            "bios":       ("BIOS / Controladores",  "Firmware, drivers y actualizaciones"),
            "soft_optim": ("Soft Optimización",     "Herramienta Chris Titus Tech WinUtil"),
            "telemetry":  ("Suprimir Telemetría",   "Desactivar rastreo y diagnóstico de Windows"),
            "pagefile":   ("Archivo de Paginación", "Memoria virtual y indexado"),
            "power":      ("Plan de Energía",       "Rendimiento del procesador"),
            "visual":     ("Efectos Visuales",      "Animaciones y transparencias"),
        }
        title, sub = TITLES.get(key, (key, ""))
        self._page_title.configure(text=title)
        self._page_sub.configure(text=sub)

        for k, btn in self._nav_btns.items():
            btn.configure(
                bg=CLR["highlight"] if k == key else CLR["panel"],
                fg=CLR["accent"]    if k == key else CLR["text"]
            )
        for k, frame in self._section_frames.items():
            if k == key:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()

    def _build_content_area(self):
        self._content = Frame(self.root, bg=CLR["bg"])
        self._content.pack(side="left", fill="both", expand=True)

        # ── Header ──
        hdr = Frame(self._content, bg=CLR["bg"], pady=14, padx=22)
        hdr.pack(fill="x")

        title_col = Frame(hdr, bg=CLR["bg"])
        title_col.pack(side="left", fill="y")

        self._page_title = Label(
            title_col, text="Inicio Automático", font=FONT_TITLE,
            bg=CLR["bg"], fg=CLR["text"]
        )
        self._page_title.pack(anchor="w")
        self._page_sub = Label(
            title_col, text="Gestión de programas de arranque",
            font=FONT_SMALL, bg=CLR["bg"], fg=CLR["muted"]
        )
        self._page_sub.pack(anchor="w")

        # Botón Ejecutar
        self._run_btn = Button(
            hdr, text="  ▶  Ejecutar seleccionados  ",
            font=FONT_BTN, bg=CLR["accent"], fg="white",
            activebackground="#3a7ae0", activeforeground="white",
            relief="flat", bd=0, padx=14, pady=9, cursor="hand2",
            command=self._run_current_section
        )
        self._run_btn.pack(side="right")

        ttk.Separator(self._content, orient="horizontal").pack(fill="x")

        # ── Contenedor de secciones ──
        self._sections_container = Frame(self._content, bg=CLR["bg"])
        self._sections_container.pack(fill="both", expand=True)

        self._section_frames: dict[str, Frame] = {}
        self._progress_panels: dict[str, ProgressPanel] = {}

        builders = {
            "startup":    self._build_startup_section,
            "services":   self._build_services_section,
            "cleanup":    self._build_cleanup_section,
            "bios":       self._build_bios_section,
            "soft_optim": self._build_soft_optim_section,
            "telemetry":  self._build_telemetry_section,
            "pagefile":   self._build_pagefile_section,
            "power":      self._build_power_section,
            "visual":     self._build_visual_section,
        }
        for key, builder in builders.items():
            frame = Frame(self._sections_container, bg=CLR["bg"])
            self._section_frames[key] = frame
            builder(frame)

        self._switch_section("startup")

    # ── Helpers de UI ─────────────────────────

    def _scrollable_frame(self, parent) -> tuple[Frame, Canvas]:
        """Frame con scroll vertical dentro del parent."""
        canvas = Canvas(parent, bg=CLR["bg"], highlightthickness=0)
        scrollbar = Scrollbar(parent, orient="vertical", command=canvas.yview)
        sf = Frame(canvas, bg=CLR["bg"])

        sf.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=sf, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units")
        )
        return sf, canvas

    def _card(self, parent, title: str = "", pad: int = 8) -> Frame:
        """Tarjeta con borde y fondo de panel."""
        outer = Frame(parent, bg=CLR["border"], pady=1)
        outer.pack(fill="x", padx=12, pady=4)
        inner = Frame(outer, bg=CLR["card"], padx=pad, pady=pad)
        inner.pack(fill="x")
        if title:
            Label(inner, text=title, font=("Segoe UI Semibold", 10),
                  bg=CLR["card"], fg=CLR["accent"]).pack(anchor="w", pady=(0, 4))
        return inner

    def _badge(self, parent, text: str, color: str) -> Label:
        return Label(parent, text=f" {text} ", font=FONT_BADGE,
                     bg=color, fg="white", relief="flat", padx=4)

    def _progress_panel(self, parent, key: str) -> ProgressPanel:
        """Crea e instala un ProgressPanel en el parent y lo registra."""
        # Borde superior del panel de progreso
        outer = Frame(parent, bg=CLR["border"], pady=1)
        outer.pack(fill="x", padx=14, pady=(8, 6))
        pp = ProgressPanel(outer, section_name=key)
        pp.pack(fill="x")
        self._progress_panels[key] = pp
        return pp

    # ── Construcción de secciones ──────────────

    def _build_startup_section(self, parent: Frame):
        self._startup_vars: dict[str, BooleanVar] = {}

        # Panel de progreso arriba
        self._progress_panel(parent, "startup")

        sf, _ = self._scrollable_frame(parent)

        info = self._card(sf, "ℹ️  Aplicaciones de inicio de terceros")
        Label(
            info,
            text=(
                "Programas cargados automáticamente al arrancar Windows, ordenados por impacto. "
                "Selecciona los que deseas deshabilitar.\n"
                "⚠  No se listan controladores de GPU, audio ni servicios de seguridad."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w")

        self._startup_list_frame = self._card(sf, "Aplicaciones detectadas")
        self._startup_placeholder = Label(
            self._startup_list_frame,
            text="Analizando programas de inicio...",
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"]
        )
        self._startup_placeholder.pack(pady=8)

    def _populate_startup_list(self):
        self._startup_placeholder.destroy()
        IMPACT_CLR = {"Alto": CLR["red"], "Medio": CLR["yellow"], "Bajo": CLR["muted"]}

        if not self.startup_apps:
            Label(self._startup_list_frame,
                  text="✓  No se detectaron aplicaciones de inicio de terceros.",
                  font=FONT_LABEL, bg=CLR["card"], fg=CLR["green"]).pack(pady=8)
            return

        for app in self.startup_apps:
            row = Frame(self._startup_list_frame, bg=CLR["card"], pady=5)
            row.pack(fill="x")
            var = BooleanVar(value=False)
            self._startup_vars[app["name"]] = var
            ttk.Checkbutton(row, variable=var).pack(side="left", padx=(0, 6))
            Label(row, text=app["name"], font=FONT_LABEL,
                  bg=CLR["card"], fg=CLR["text"], width=26, anchor="w").pack(side="left")
            self._badge(row, app["impact"], IMPACT_CLR.get(app["impact"], CLR["muted"])
                        ).pack(side="left", padx=6)
            cmd_short = (app["cmd"][:60] + "…") if len(app["cmd"]) > 60 else app["cmd"]
            Label(row, text=cmd_short, font=FONT_MONO,
                  bg=CLR["card"], fg=CLR["muted"]).pack(side="left", padx=8)
            ttk.Separator(self._startup_list_frame, orient="horizontal").pack(fill="x", pady=1)

    def _build_services_section(self, parent: Frame):
        self._service_vars: dict[str, BooleanVar] = {}
        self._progress_panel(parent, "services")
        sf, _ = self._scrollable_frame(parent)

        info = self._card(sf, "ℹ️  Servicios no esenciales en inicio automático")
        Label(
            info,
            text=(
                "Servicios configurados en inicio automático pero rara vez necesarios desde el arranque. "
                "Cambiarlos a 'Manual' permite que Windows los inicie solo cuando sean requeridos."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w")

        self._services_list_frame = self._card(sf, "Servicios detectados")
        self._svc_placeholder = Label(
            self._services_list_frame,
            text="Analizando servicios del sistema...",
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"]
        )
        self._svc_placeholder.pack(pady=8)

    def _populate_services_list(self):
        self._svc_placeholder.destroy()
        if not self.nonessential_services:
            Label(self._services_list_frame,
                  text="✓  Ningún servicio no esencial detectado en inicio automático.",
                  font=FONT_LABEL, bg=CLR["card"], fg=CLR["green"]).pack(pady=8)
            return

        for svc in self.nonessential_services:
            row = Frame(self._services_list_frame, bg=CLR["card"], pady=5)
            row.pack(fill="x")
            var = BooleanVar(value=False)
            self._service_vars[svc["name"]] = var
            ttk.Checkbutton(row, variable=var).pack(side="left", padx=(0, 6))
            Label(row, text=svc["friendly"], font=FONT_LABEL,
                  bg=CLR["card"], fg=CLR["text"], width=32, anchor="w").pack(side="left")
            Label(row, text=f"[{svc['name']}]", font=FONT_MONO,
                  bg=CLR["card"], fg=CLR["muted"]).pack(side="left", padx=6)
            self._badge(row, svc["start_type"], CLR["yellow"]).pack(side="left")
            Label(row, text=" → Manual", font=FONT_MONO,
                  bg=CLR["card"], fg=CLR["green"]).pack(side="left", padx=6)
            ttk.Separator(self._services_list_frame, orient="horizontal").pack(fill="x", pady=1)

    def _build_cleanup_section(self, parent: Frame):
        self._cleanup_vars: dict[str, BooleanVar] = {}
        self._progress_panel(parent, "cleanup")
        sf, _ = self._scrollable_frame(parent)

        TASKS = [
            ("temp_files",           "🗑️  Archivos temporales profundos  (%TEMP% y Windows\\Temp)"),
            ("windows_update_cache", "🔄  Caché de actualizaciones de Windows  (SoftwareDistribution)"),
            ("windows_old",          "📦  Instalación anterior del SO  (Windows.old)"),
            ("recycle_bin",          "♻️  Vaciar Papelera de Reciclaje"),
            ("event_logs",           "📋  Limpiar registros de eventos del sistema"),
            ("cleanmgr",             "💾  Liberador de espacio en disco  (incluye actualizaciones)"),
            ("prefetch",             "⚡  Limpiar Prefetch  (se regenera automáticamente)"),
        ]

        info = self._card(sf, "ℹ️  Tareas de limpieza del sistema")
        Label(
            info,
            text="Todas las tareas están seleccionadas por defecto. Desmarca las que no desees ejecutar.",
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780
        ).pack(anchor="w")

        card = self._card(sf, "Selecciona las tareas")
        for task_id, label in TASKS:
            var = BooleanVar(value=True)
            self._cleanup_vars[task_id] = var
            row = Frame(card, bg=CLR["card"], pady=5)
            row.pack(fill="x")
            ttk.Checkbutton(row, variable=var).pack(side="left", padx=(0, 6))
            Label(row, text=label, font=FONT_LABEL,
                  bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    def _build_bios_section(self, parent: Frame):
        self._bios_vars: dict[str, BooleanVar] = {}
        self._drivers_data: list[dict] = []
        self._driver_updates: list[dict] = []
        self._update_drivers_enabled = BooleanVar(value=False)
        self._progress_panel(parent, "bios")
        sf, _ = self._scrollable_frame(parent)

        # ── Card 1: Recomendaciones BIOS ──
        warn = self._card(sf, "⚠️  Configuración de BIOS/UEFI")
        Label(
            warn,
            text=(
                "Recomendaciones basadas en tu hardware. Los ajustes de BIOS requieren acceso manual al firmware.\n"
                "Al confirmar, se creará un PUNTO DE RESTAURACIÓN automáticamente antes de cualquier cambio."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["yellow"],
            wraplength=780, justify="left"
        ).pack(anchor="w")

        self._bios_list_frame = self._card(sf, "Recomendaciones para tu equipo")
        self._bios_placeholder = Label(
            self._bios_list_frame, text="Analizando hardware...",
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"]
        )
        self._bios_placeholder.pack(pady=8)

        # ── Card 2: DDU ──
        ddu = self._card(sf, "🖥️  Guía: Display Driver Uninstaller (DDU)")
        Label(ddu, text=run_ddu_guide(), font=FONT_MONO,
              bg=CLR["card"], fg=CLR["text"], justify="left").pack(anchor="w")

        # ── Card 3: Controladores del sistema ──
        drv_card = self._card(sf, "🔄  Controladores del Sistema")
        Label(
            drv_card,
            text=(
                "Verificación de drivers instalados contra actualizaciones disponibles vía Windows Update.\n"
                "Las actualizaciones son oficiales y provienen del fabricante a través de los servidores de Microsoft."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 8))

        # Resumen de drivers
        self._drivers_summary_frame = Frame(drv_card, bg=CLR["card"])
        self._drivers_summary_frame.pack(fill="x", pady=(0, 6))
        self._drivers_summary_lbl = Label(
            self._drivers_summary_frame,
            text="🔍  Escaneando controladores del sistema...",
            font=FONT_STATUS, bg=CLR["card"], fg=CLR["muted"]
        )
        self._drivers_summary_lbl.pack(anchor="w")

        # Lista de drivers
        self._drivers_list_frame = Frame(drv_card, bg=CLR["card"])
        self._drivers_list_frame.pack(fill="x")

        # Checkbox para actualizar
        self._drivers_action_frame = Frame(drv_card, bg=CLR["card"])
        self._drivers_action_frame.pack(fill="x", pady=(8, 0))
        row_upd = Frame(self._drivers_action_frame, bg=CLR["card"], pady=4)
        row_upd.pack(fill="x")
        ttk.Checkbutton(row_upd, variable=self._update_drivers_enabled).pack(side="left", padx=(0, 6))
        Label(row_upd, text="Actualizar todos los drivers desactualizados (vía Windows Update oficial)",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    def _populate_bios_list(self):
        self._bios_placeholder.destroy()
        for rec in check_bios_recommendations(self.sys_info):
            var = BooleanVar(value=False)
            self._bios_vars[rec["id"]] = var
            row = Frame(self._bios_list_frame, bg=CLR["card"], pady=7)
            row.pack(fill="x")
            ttk.Checkbutton(row, variable=var).pack(side="left", anchor="n", padx=(0, 8))
            col = Frame(row, bg=CLR["card"])
            col.pack(side="left", fill="x", expand=True)
            Label(col, text=rec["title"], font=("Segoe UI Semibold", 10),
                  bg=CLR["card"], fg=CLR["accent"]).pack(anchor="w")
            Label(col, text=rec["desc"], font=FONT_LABEL,
                  bg=CLR["card"], fg=CLR["text"],
                  wraplength=700, justify="left").pack(anchor="w")
            Label(col, text=f"📋 {rec['guide']}", font=FONT_MONO,
                  bg=CLR["card"], fg=CLR["green"],
                  wraplength=700, justify="left").pack(anchor="w", pady=(2, 0))
            ttk.Separator(self._bios_list_frame, orient="horizontal").pack(fill="x", pady=3)

    def _populate_drivers_list(self):
        """Rellena la lista de controladores con su estado de actualización."""
        for child in self._drivers_list_frame.winfo_children():
            child.destroy()

        if not self._drivers_data:
            Label(self._drivers_list_frame,
                  text="⚠  No se pudieron detectar controladores.",
                  font=FONT_LABEL, bg=CLR["card"], fg=CLR["yellow"]).pack(pady=4)
            return

        outdated_drivers = [d for d in self._drivers_data if d["has_update"]]
        total_count = len(self._drivers_data)
        outdated_count = len(outdated_drivers)

        if outdated_count > 0:
            self._drivers_summary_lbl.configure(
                text=f"⚠  {outdated_count} driver(s) requieren actualización (de {total_count} analizados)",
                fg=CLR["yellow"]
            )
        else:
            self._drivers_summary_lbl.configure(
                text=f"✓  Todos los {total_count} controladores están en su última versión",
                fg=CLR["green"]
            )
            Label(self._drivers_list_frame,
                  text="No hay actualizaciones de controladores pendientes.",
                  font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"]).pack(pady=4)
            return

        # Mostrar solo los desactualizados
        for drv in outdated_drivers:
            row = Frame(self._drivers_list_frame, bg=CLR["card"], pady=2)
            row.pack(fill="x")

            Label(row, text="⬆️", font=("Segoe UI", 10),
                  bg=CLR["card"], fg=CLR["yellow"], width=3).pack(side="left")

            info_col = Frame(row, bg=CLR["card"])
            info_col.pack(side="left", fill="x", expand=True)
            Label(info_col, text=drv["name"], font=FONT_LABEL,
                  bg=CLR["card"], fg=CLR["text"], anchor="w").pack(anchor="w")
            detail = f"{drv['manufacturer']}  •  v{drv['version']}  •  {drv['date']}"
            Label(info_col, text=detail, font=FONT_MONO,
                  bg=CLR["card"], fg=CLR["muted"], anchor="w").pack(anchor="w")

            self._badge(row, "Actualización Oficial (WHQL)", CLR["accent"]).pack(side="right", padx=6)
            ttk.Separator(self._drivers_list_frame, orient="horizontal").pack(fill="x", pady=1)

    # ── Soft Optimización (Chris Titus Tech) ──

    def _build_soft_optim_section(self, parent: Frame):
        self._soft_optim_enabled = BooleanVar(value=True)
        self._progress_panel(parent, "soft_optim")
        sf, _ = self._scrollable_frame(parent)

        # Una sola tarjeta principal
        card = self._card(sf, "✨  Chris Titus Tech — WinUtil")

        # Texto introducción
        Label(
            card,
            text=(
                "Herramienta open-source extremadamente popular para optimizar Windows.\n"
                "Permite desinstalar bloatware UWP, deshabilitar servicios innecesarios y aplicar tweaks."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 6))

        # Contenedor dividido 2 columnas: Lista a la izquierda, Info/Comando a la derecha
        split_frame = Frame(card, bg=CLR["card"])
        split_frame.pack(fill="x", pady=4)

        left_col = Frame(split_frame, bg=CLR["card"])
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_col = Frame(split_frame, bg=CLR["card"])
        right_col.pack(side="left", fill="both")

        # Lista izquierda
        Label(left_col, text="¿Qué puedes hacer?", font=("Segoe UI Semibold", 9),
              bg=CLR["card"], fg=CLR["muted"]).pack(anchor="w")
        features_list = [
            "🗑️  Desinstalar bloatware UWP",
            "⚙️  Deshabilitar telemetría y rastreo",
            "🚀  Aplicar tweaks de rendimiento locales",
            "🔧  Configurar actualizaciones de Windows",
            "🛡️  Instalar software esencial rápidamente"
        ]
        for feat in features_list:
            Label(left_col, text=feat, font=FONT_SMALL,
                  bg=CLR["card"], fg=CLR["text"]).pack(anchor="w", pady=1)

        # Info derecha
        Label(right_col, text="Ejecución mediante terminal", font=("Segoe UI Semibold", 9),
              bg=CLR["card"], fg=CLR["muted"]).pack(anchor="w")
        Label(right_col, text="iwr -useb christitus.com/win | iex",
              font=("Consolas", 9), bg=CLR["card"], fg=CLR["accent"]).pack(anchor="w", pady=(2, 6))

        warn_frame = Frame(right_col, bg=CLR["highlight"], padx=6, pady=4)
        warn_frame.pack(fill="x")
        Label(warn_frame, text="⚠  IMPORTANTE: Solo usar perfil 'Desktop' o 'Laptop'.\nNUNCA usar el perfil 'Minimal'.",
              font=FONT_SMALL, bg=CLR["highlight"], fg=CLR["yellow"], justify="left").pack()

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=8)

        # Confirmación
        row = Frame(card, bg=CLR["card"], pady=2)
        row.pack(fill="x")
        ttk.Checkbutton(row, variable=self._soft_optim_enabled).pack(side="left", padx=(0, 6))
        Label(row, text="Abrir Chris Titus Tech WinUtil en una ventana de PowerShell nueva",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    # ── Suprimir Telemetría ──

    def _build_telemetry_section(self, parent: Frame):
        self._telemetry_enabled = BooleanVar(value=True)
        self._progress_panel(parent, "telemetry")
        sf, _ = self._scrollable_frame(parent)

        # Card unificada
        card = self._card(sf, "🛡️  Supresión de Telemetría de Windows")

        # Intro y seguridad apilada
        desc = Frame(card, bg=CLR["card"])
        desc.pack(fill="x", pady=(0, 6))
        Label(
            desc,
            text=(
                "Desactiva el envío de datos de diagnóstico y telemetría de Windows.\n"
                "✓ Operación 100% local (sin red)  |  ✓ Punto de restauración automático  |  ✓ Totalmente reversible"
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"],
            wraplength=780, justify="left"
        ).pack(anchor="w")

        # Grid para ítems
        details = Frame(card, bg=CLR["highlight"], padx=8, pady=4)
        details.pack(fill="x", pady=4)
        Label(details, text="📋 Ajustes a aplicar (vía Registro y Servicios):", font=("Segoe UI Semibold", 9),
              bg=CLR["highlight"], fg=CLR["muted"]).pack(anchor="w", pady=(0, 4))

        items = [
            ("DiagTrack", "Detiene Servicio de Telemetría"),
            ("dmwappushservice", "Detiene enrutamiento WAP"),
            ("AllowTelemetry = 0", "Nivel telemetría: Seguridad"),
            ("Notificaciones = 0", "Oculta pop-ups de feedback"),
        ]

        grid_frame = Frame(details, bg=CLR["highlight"])
        grid_frame.pack(fill="x")
        for i, (name, dinfo) in enumerate(items):
            row = i // 2
            col = i % 2
            cell = Frame(grid_frame, bg=CLR["highlight"])
            cell.grid(row=row, column=col, sticky="ew", padx=(0, 15), pady=2)
            grid_frame.columnconfigure(col, weight=1)
            
            Label(cell, text="🔒", font=("Segoe UI", 9), bg=CLR["highlight"], fg=CLR["green"]).pack(side="left")
            Label(cell, text=name, font=("Segoe UI Semibold", 9), bg=CLR["highlight"], fg=CLR["accent"]).pack(side="left", padx=4)
            Label(cell, text=f"— {dinfo}", font=FONT_SMALL, bg=CLR["highlight"], fg=CLR["muted"]).pack(side="left")

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=8)

        # Checkbox confirmación
        row = Frame(card, bg=CLR["card"], pady=2)
        row.pack(fill="x")
        ttk.Checkbutton(row, variable=self._telemetry_enabled).pack(side="left", padx=(0, 6))
        Label(row, text="Confirmar: Suprimir telemetría de Windows de forma segura",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    def _build_pagefile_section(self, parent: Frame):
        self._pagefile_enabled = BooleanVar(value=True)
        self._indexer_enabled  = BooleanVar(value=False)
        self._progress_panel(parent, "pagefile")
        sf, _ = self._scrollable_frame(parent)

        pf = self._card(sf, "📄  Archivo de Paginación (Pagefile)")
        Label(
            pf,
            text=(
                "Se configurará un tamaño fijo recomendado según tu RAM para evitar que "
                "Windows redimensione el pagefile en tiempo real, reduciendo la E/S del disco."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 8))
        self._pagefile_info_label = Label(
            pf, text="Calculando...", font=("Segoe UI Semibold", 11),
            bg=CLR["card"], fg=CLR["green"]
        )
        self._pagefile_info_label.pack(anchor="w")
        row = Frame(pf, bg=CLR["card"], pady=6)
        row.pack(fill="x")
        ttk.Checkbutton(row, variable=self._pagefile_enabled).pack(side="left", padx=(0, 6))
        Label(row, text="Aplicar configuración de pagefile recomendada",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

        idx = self._card(sf, "🔍  Exclusiones del Indexador de Windows Search")
        Label(
            idx,
            text=(
                "Excluir carpetas de proyectos o caché de desarrollo evita que "
                "SearchIndexer.exe consuma CPU rastreando cambios en el código."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 8))
        row2 = Frame(idx, bg=CLR["card"], pady=4)
        row2.pack(fill="x")
        ttk.Checkbutton(row2, variable=self._indexer_enabled).pack(side="left", padx=(0, 6))
        Label(row2, text="Excluir carpetas de desarrollo del indexado",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")
        Label(
            idx,
            text="Rutas: node_modules, .git, __pycache__, dist, build, .gradle, .idea, .vs",
            font=FONT_MONO, bg=CLR["card"], fg=CLR["muted"]
        ).pack(anchor="w", pady=(4, 0))

    def _build_power_section(self, parent: Frame):
        self._power_enabled = BooleanVar(value=True)
        self._progress_panel(parent, "power")
        sf, _ = self._scrollable_frame(parent)

        card = self._card(sf, "⚡  Plan de Energía — Máximo Rendimiento")
        Label(
            card,
            text=(
                "Activa el plan SCHEME_MIN y fija el procesador al 100% mínimo y máximo, "
                "eliminando el throttling dinámico de Windows.\n"
                "Recomendado para PC de escritorio o laptop conectada a corriente."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 8))
        row = Frame(card, bg=CLR["card"], pady=4)
        row.pack(fill="x")
        ttk.Checkbutton(row, variable=self._power_enabled).pack(side="left", padx=(0, 6))
        Label(row, text="Activar SCHEME_MIN y CPU al 100% mínimo / máximo",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    def _build_visual_section(self, parent: Frame):
        self._visual_enabled = BooleanVar(value=True)
        self._progress_panel(parent, "visual")
        sf, _ = self._scrollable_frame(parent)

        card = self._card(sf, "🎨  Efectos Visuales — Mejor Rendimiento")
        Label(
            card,
            text=(
                "Configura Windows en 'Mejor rendimiento' conservando solo:\n"
                "  • Sombras debajo de las ventanas\n"
                "  • Fuentes suavizadas (ClearType)\n"
                "  • Miniaturas en lugar de iconos genéricos\n\n"
                "Animaciones, transparencias y efectos Aero se desactivan."
            ),
            font=FONT_LABEL, bg=CLR["card"], fg=CLR["muted"],
            wraplength=780, justify="left"
        ).pack(anchor="w", pady=(0, 8))
        row = Frame(card, bg=CLR["card"], pady=4)
        row.pack(fill="x")
        ttk.Checkbutton(row, variable=self._visual_enabled).pack(side="left", padx=(0, 6))
        Label(row, text="Aplicar configuración de efectos visuales para rendimiento",
              font=FONT_LABEL, bg=CLR["card"], fg=CLR["text"]).pack(side="left")

    # ── Carga asíncrona de datos ───────────────

    def _load_data_async(self):
        def _work():
            self.sys_info              = get_system_info()
            self.startup_apps          = get_startup_apps()
            self.nonessential_services = get_nonessential_services()
            self.root.after(0, self._on_data_loaded)
            # Escaneo de drivers en segundo plano (puede tardar)
            try:
                drivers, updates = get_outdated_drivers()
            except Exception:
                drivers, updates = [], []
            self._drivers_data    = drivers
            self._driver_updates  = updates
            self.root.after(0, self._on_drivers_loaded)
        threading.Thread(target=_work, daemon=True).start()

    def _on_data_loaded(self):
        info   = self.sys_info
        cpu    = info.get("cpu", "?")[:30]
        ram_mb = info.get("ram_mb", 8192)
        rec_mb = int(min(max(ram_mb * 1.5, 4096), 16384))
        drive  = info.get("fast_disk", "C")

        self._sys_label.configure(
            text=f"CPU: {cpu}\nRAM: {ram_mb} MB\nWindows {info.get('win_release','?')}"
        )
        self._pagefile_info_label.configure(
            text=f"Recomendado: {rec_mb} MB ({rec_mb // 1024} GB) — Disco: {drive}:\\"
        )
        self._populate_startup_list()
        self._populate_services_list()
        self._populate_bios_list()

        pp = self._progress_panels.get("startup")
        if pp:
            pp.log(
                f"Sistema listo: {cpu} | {ram_mb} MB RAM | Win {info.get('win_release','')}",
                "info"
            )

    def _on_drivers_loaded(self):
        """Llamado desde el hilo principal una vez terminado el escaneo de drivers."""
        try:
            self._populate_drivers_list()
        except Exception:
            pass

    # ── Dispatcher de ejecución ────────────────

    def _run_current_section(self):
        """Lanza la ejecución de la sección activa en un hilo secundario."""
        if self._running:
            messagebox.showwarning(
                "En progreso",
                "Ya hay una operación en curso. Espera a que termine."
            )
            return
        section = self.current_section.get()
        runners = {
            "startup":    self._run_startup,
            "services":   self._run_services,
            "cleanup":    self._run_cleanup,
            "bios":       self._run_bios,
            "soft_optim": self._run_soft_optim,
            "telemetry":  self._run_telemetry,
            "pagefile":   self._run_pagefile,
            "power":      self._run_power,
            "visual":     self._run_visual,
        }
        runner = runners.get(section)
        if not runner:
            return

        self._running = True
        self._run_btn.configure(
            state="disabled", text="  ⏳ Ejecutando...",
            bg=CLR["muted"]
        )
        threading.Thread(target=self._wrap_runner(runner, section), daemon=True).start()

    def _wrap_runner(self, runner, section: str):
        """Envuelve el runner para gestionar estado antes y después."""
        def _execute():
            try:
                runner()
            finally:
                self._running = False
                self.root.after(0, self._restore_run_btn)
        return _execute

    def _restore_run_btn(self):
        self._run_btn.configure(
            state="normal", text="  ▶  Ejecutar seleccionados  ",
            bg=CLR["accent"]
        )

    # ── Helpers de progreso ────────────────────

    def _pp(self, key: str) -> ProgressPanel | None:
        """Atajo para acceder al ProgressPanel de una sección."""
        return self._progress_panels.get(key)

    def _ui(self, fn):
        """Ejecuta fn en el hilo principal de la UI."""
        self.root.after(0, fn)

    def _pp_log(self, key: str, msg: str, ok: bool | None = None):
        """
        Emite un mensaje en el ProgressPanel de la sección key.
        ok=True → verde, ok=False → rojo, ok=None → default
        """
        tag = "ok" if ok is True else "error" if ok is False else "default"
        pp = self._pp(key)
        if pp:
            self._ui(lambda m=msg, t=tag: pp.log(m, t))

    def _pp_step(self, key: str, current: int, total: int, label: str):
        """Actualiza barra de progreso y texto de tarea."""
        pp = self._pp(key)
        if not pp:
            return
        frac = current / total if total > 0 else 0.0
        self._ui(lambda f=frac, l=label: pp.set_progress(f, l))
        self._ui(lambda c=current, e=0, t=total: pp.update_counter(c, e, t))

    def _pp_finish(self, key: str, errors: int, total: int):
        """Cierra la ejecución del ProgressPanel con el estado correcto."""
        pp = self._pp(key)
        if not pp:
            return
        ok_count = total - errors
        state = ST_DONE if errors == 0 else (ST_PARTIAL if ok_count > 0 else ST_ERROR)
        self._ui(lambda: pp.set_progress(1.0))
        self._ui(lambda: pp.set_state(state))
        self._ui(lambda c=ok_count, e=errors, t=total: pp.update_counter(c, e, t))
        self._update_nav_state(key, state)

    # ── Runners por sección ────────────────────

    def _run_startup(self):
        key = "startup"
        pp  = self._pp(key)
        selected = [n for n, v in self._startup_vars.items() if v.get()]

        if not selected:
            self._ui(lambda: pp.log("⚠  No se seleccionó ninguna aplicación.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log(f"Deshabilitando {len(selected)} aplicación(es)…", "info"))

        app_map = {a["name"]: a for a in self.startup_apps}
        errors = 0

        for i, name in enumerate(selected, 1):
            self._ui(lambda i=i, n=name: pp.set_progress(
                (i - 1) / len(selected), f"Procesando: {n}"
            ))
            ok = disable_startup_app(app_map[name])
            if not ok:
                errors += 1
            self._pp_log(key, f"{'✓' if ok else '✗'}  {name}", ok)

        self._ui(lambda: pp.set_progress(1.0, "Completado"))
        if errors == 0:
            self._ui(lambda: pp.log("✓  Todos los cambios aplicados. Reinicia Windows.", "ok"))
        else:
            self._ui(lambda e=errors: pp.log(
                f"⚠  {e} elemento(s) no pudieron deshabilitarse.", "warn"
            ))
        self._pp_finish(key, errors, len(selected))

    def _run_services(self):
        key      = "services"
        pp       = self._pp(key)
        selected = [n for n, v in self._service_vars.items() if v.get()]

        if not selected:
            self._ui(lambda: pp.log("⚠  No se seleccionó ningún servicio.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log(f"Cambiando {len(selected)} servicio(s) a Manual…", "info"))
        errors = 0

        for i, name in enumerate(selected, 1):
            self._ui(lambda i=i, n=name: pp.set_progress(
                (i - 1) / len(selected), f"Configurando: {name}"
            ))
            ok = set_service_manual(name)
            if not ok:
                errors += 1
            self._pp_log(key, f"{'✓' if ok else '✗'}  {name} → Manual", ok)

        self._pp_finish(key, errors, len(selected))

    def _run_cleanup(self):
        key   = "cleanup"
        pp    = self._pp(key)
        tasks = [tid for tid, v in self._cleanup_vars.items() if v.get()]

        if not tasks:
            self._ui(lambda: pp.log("⚠  No se seleccionó ninguna tarea.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda n=len(tasks): pp.log(f"Iniciando {n} tarea(s) de limpieza…", "info"))

        ok_count = [0]
        err_count = [0]

        def step_cb(current, total, label):
            frac = current / total if total > 0 else 0.0
            self._ui(lambda f=frac, l=label: pp.set_progress(f, l))
            self._ui(lambda c=ok_count[0], e=err_count[0], t=total:
                     pp.update_counter(c, e, t))

        def log_cb(msg, ok):
            tag = "ok" if ok else "warn"
            if not ok:
                err_count[0] += 1
            else:
                ok_count[0] += 1
            self._ui(lambda m=msg, t=tag: pp.log(m, t))

        results = perform_cleanup(tasks, step_cb, log_cb)
        errors  = sum(1 for v in results.values() if not v)
        self._pp_finish(key, errors, len(tasks))

    def _run_bios(self):
        key      = "bios"
        pp       = self._pp(key)
        selected = [bid for bid, v in self._bios_vars.items() if v.get()]
        run_driver_update = self._update_drivers_enabled.get()

        if not selected and not run_driver_update:
            self._ui(lambda: pp.log(
                "⚠  No se seleccionó ninguna recomendación ni actualización de drivers.", "warn"
            ))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log("🔒  Creando punto de restauración del sistema…", "info"))
        self._ui(lambda: pp.set_progress(-1, "Creando punto de restauración…"))

        ok, msg = create_restore_point()
        if ok:
            self._pp_log(key, "✓  Punto de restauración creado.", True)
        else:
            self._pp_log(key, f"⚠  No se pudo crear el punto de restauración: {msg}", False)
            confirm = [False]
            event   = threading.Event()

            def _ask():
                confirm[0] = messagebox.askyesno(
                    "Advertencia",
                    "No se pudo crear el punto de restauración.\n¿Continuar de todas formas?"
                )
                event.set()

            self._ui(_ask)
            event.wait()
            if not confirm[0]:
                self._ui(lambda: pp.set_state(ST_IDLE))
                return

        errors = 0
        total_tasks = len(selected) + (1 if run_driver_update else 0)

        # ── Recomendaciones BIOS ──
        for i, bid in enumerate(selected, 1):
            self._ui(lambda i=i, b=bid: pp.set_progress(
                i / total_tasks, f"Revisando: {b}"
            ))
            self._pp_log(
                key,
                f"ℹ  '{bid}': ajuste manual requerido en BIOS. Consulta la guía.",
                None
            )
            time.sleep(0.3)

        if selected:
            self._ui(lambda: pp.log(
                "✓  Revisión BIOS completada. Accede al BIOS con DEL/F2 al reiniciar.", "ok"
            ))

        # ── Actualización de drivers ──
        if run_driver_update:
            self._ui(lambda: pp.log("🔄  Iniciando actualización de drivers vía Windows Update…", "info"))
            self._ui(lambda n=total_tasks: pp.set_progress(
                len(selected) / n if n > 0 else 0.5, "Actualizando drivers…"
            ))

            ok2, out = update_all_drivers()
            lines = out.splitlines()
            for line in lines:
                if line.startswith("OK:"):
                    self._pp_log(key, f"✓  {line[3:]}", True)
                elif line.startswith("FAIL:"):
                    self._pp_log(key, f"✗  {line[5:]}", False)
                    errors += 1
                elif line == "NO_UPDATES":
                    self._pp_log(key, "✓  No hay actualizaciones de drivers pendientes.", True)
                elif line == "REBOOT_REQUIRED":
                    self._pp_log(key, "⚠  Se requiere reiniciar para aplicar algunos drivers.", None)
                elif line.startswith("ERROR:"):
                    self._pp_log(key, f"✗  {line[6:]}", False)
                    errors += 1
                elif line.startswith("DONE:"):
                    self._pp_log(key, f"✓  Drivers instalados: {line[5:]}", True)

            # Refrescar lista de drivers en UI
            try:
                drivers, updates = get_outdated_drivers()
                self._drivers_data   = drivers
                self._driver_updates = updates
                self._ui(self._on_drivers_loaded)
            except Exception:
                pass

        self._pp_finish(key, errors, total_tasks)

    def _run_soft_optim(self):
        key = "soft_optim"
        pp  = self._pp(key)

        if not self._soft_optim_enabled.get():
            self._ui(lambda: pp.log("⚠  Opción no marcada para ejecutar.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log("🔒  Creando punto de restauración…", "info"))
        self._ui(lambda: pp.set_progress(-1, "Creando punto de restauración…"))

        ok, _ = create_restore_point()
        self._pp_log(key,
                     "✓  Punto de restauración creado." if ok
                     else "⚠  Punto de restauración no disponible. Continúa con precaución.",
                     ok)

        self._pp_log(key, "✨  Lanzando Chris Titus Tech WinUtil en PowerShell…", None)
        self._ui(lambda: pp.set_progress(0.5, "Abriendo WinUtil…"))

        try:
            subprocess.Popen(
                ["powershell", "-NoExit", "-Command",
                 "iwr -useb https://christitus.com/win | iex"],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            self._pp_log(key,
                         "✓  PowerShell abierto. Selecciona 'Desktop' o 'Laptop' en WinUtil.",
                         True)
            self._pp_log(key, "ℹ  NUNCA uses el perfil 'Minimal'.", None)
            errors = 0
        except Exception as e:
            self._pp_log(key, f"✗  No se pudo lanzar PowerShell: {e}", False)
            errors = 1

        self._pp_finish(key, errors, 1)

    def _run_telemetry(self):
        key = "telemetry"
        pp  = self._pp(key)

        if not self._telemetry_enabled.get():
            self._ui(lambda: pp.log("⚠  Opción no marcada para ejecutar.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log("🔒  Creando punto de restauración…", "info"))
        self._ui(lambda: pp.set_progress(-1, "Creando punto de restauración…"))

        ok, _ = create_restore_point()
        self._pp_log(key,
                     "✓  Punto de restauración creado." if ok
                     else "⚠  Punto de restauración no disponible. Continúa con precaución.",
                     ok)

        self._pp_log(key, "🛡️  Aplicando claves de registro para suprimir telemetría…", None)
        self._ui(lambda: pp.set_progress(0.4, "Deshabilitando DiagTrack…"))

        ok2, out = apply_telemetry_registry()
        if ok2:
            self._pp_log(key, "✓  DiagTrack detenido y deshabilitado.", True)
            self._pp_log(key, "✓  dmwappushservice detenido y deshabilitado.", True)
            self._pp_log(key, "✓  Clave AllowTelemetry = 0 aplicada en políticas de grupo.", True)
            self._pp_log(key, "✓  DoNotShowFeedbackNotifications = 1 aplicado.", True)
            self._pp_log(key, "ℹ  Reinicia para confirmar todos los cambios.", None)
        else:
            self._pp_log(key, f"✗  Error aplicando telemetría: {out}", False)

        self._pp_finish(key, 0 if ok2 else 1, 1)

    def _run_pagefile(self):
        key    = "pagefile"
        pp     = self._pp(key)
        tasks  = []
        if self._pagefile_enabled.get():
            tasks.append("pagefile")
        if self._indexer_enabled.get():
            tasks.append("indexer")

        if not tasks:
            self._ui(lambda: pp.log("⚠  No se seleccionó ninguna opción.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        errors = 0
        total  = len(tasks)

        for i, task in enumerate(tasks, 1):
            frac = (i - 1) / total

            if task == "pagefile":
                self._ui(lambda f=frac: pp.set_progress(f, "Configurando pagefile…"))
                ram   = self.sys_info.get("ram_mb", 8192)
                drive = self.sys_info.get("fast_disk", "C") + ":"
                ok, msg = configure_pagefile(ram, drive)
                if not ok:
                    errors += 1
                self._pp_log(key, f"{'✓' if ok else '✗'}  Pagefile: {msg}", ok)

            elif task == "indexer":
                self._ui(lambda f=frac: pp.set_progress(f, "Aplicando exclusiones del indexador…"))
                dev_paths = [
                    r"C:\node_modules", r"%USERPROFILE%\.gradle",
                    r"%USERPROFILE%\.m2", r"%USERPROFILE%\AppData\Local\Temp",
                ]
                ok, msg = exclude_search_indexer(dev_paths)
                if not ok:
                    errors += 1
                self._pp_log(key, f"{'✓' if ok else '✗'}  Indexador: {msg}", ok)

        self._pp_finish(key, errors, total)

    def _run_power(self):
        key = "power"
        pp  = self._pp(key)

        if not self._power_enabled.get():
            self._ui(lambda: pp.log("⚠  Ningún ajuste seleccionado.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log("Activando plan de energía SCHEME_MIN…", "info"))
        self._ui(lambda: pp.set_progress(-1, "Configurando plan de energía…"))

        ok, msg = apply_power_plan_ultimate()
        self._pp_log(key, f"{'✓' if ok else '✗'}  {msg}", ok)
        self._pp_finish(key, 0 if ok else 1, 1)

    def _run_visual(self):
        key = "visual"
        pp  = self._pp(key)

        if not self._visual_enabled.get():
            self._ui(lambda: pp.log("⚠  Ningún ajuste seleccionado.", "warn"))
            return

        self._ui(lambda: pp.start_run())
        self._ui(lambda: pp.log("Optimizando efectos visuales…", "info"))
        self._ui(lambda: pp.set_progress(-1, "Aplicando configuración…"))

        ok, msg = apply_visual_performance()
        self._pp_log(key, f"{'✓' if ok else '✗'}  {msg}", ok)
        if ok:
            self._pp_log(key,
                         "ℹ  Cierra sesión o reinicia para que los efectos tomen efecto.", None)
        self._pp_finish(key, 0 if ok else 1, 1)


# ─────────────────────────────────────────────
#  PUNTO DE ENTRADA
# ─────────────────────────────────────────────

def main():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit(0)

    root = Tk()

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TCheckbutton",
                    background=CLR["card"],
                    foreground=CLR["text"],
                    focuscolor=CLR["card"])
    style.configure("TSeparator",  background=CLR["border"])
    style.configure("TScrollbar",
                    background=CLR["panel"],
                    troughcolor=CLR["bg"],
                    arrowcolor=CLR["muted"])

    WinOptimizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()