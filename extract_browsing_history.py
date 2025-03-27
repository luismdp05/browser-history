import os
import sqlite3
import pandas as pd
import datetime
from rich.console import Console
from rich.prompt import Prompt
from rich.progress import track
from rich.panel import Panel
from time import sleep
import glob  # Para buscar el perfil de Firefox automáticamente

console = Console()

# Diccionario con rutas de historial y emojis por navegador
BROWSERS = {
    "Google Chrome": {
        "path": os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data\Default\History"),
        "icon": "🌙"
    },
    "Microsoft Edge": {
        "path": os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\History"),
        "icon": "🟩"
    },
    "Firefox": {
        "path": os.path.expandvars(r"%APPDATA%\Mozilla\Firefox\Profiles"),
        "icon": "🔥"
    },
    "Brave": {
        "path": os.path.expandvars(r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data\Default\History"),
        "icon": "🦁"
    },
    "Opera": {
        "path": os.path.expandvars(r"%APPDATA%\Opera Software\Opera Stable\History"),
        "icon": "🎭"
    },
    "Vivaldi": {
        "path": os.path.expandvars(r"%LOCALAPPDATA%\Vivaldi\User Data\Default\History"),
        "icon": "🚀"
    },
    "Yandex Browser": {
        "path": os.path.expandvars(r"%LOCALAPPDATA%\Yandex\YandexBrowser\User Data\Default\History"),
        "icon": "🇷🇺"
    }
}

def show_menu():
    """Muestra el menú con el formato de tabla."""
    menu_text = """
 ╔══════════════════════════════════════════════════════════════════════╗
 ║                      🌐 EXTRACTOR DE HISTORIAL 🌐                    ║
 ╠══════════════════════════════════════════════════════════════════════╣
 ║ Selecciona el navegador del que deseas extraer el historial:         ║
 ║                                                                      ║
 ║ 1. 🌍 Google Chrome                                                  ║
 ║ 2. 🔵 Microsoft Edge                                                 ║
 ║ 3. 🦊 Firefox                                                        ║
 ║ 4. 🦁 Brave                                                          ║
 ║ 5. 🎭 Opera                                                          ║
 ║ 6. 🚀 Vivaldi                                                        ║
 ║ 7. 🇷🇺 Yandex Browser                                                 ║
 ║ 0. ❌ Cancelar y salir                                               ║
 ╚══════════════════════════════════════════════════════════════════════╝
"""
    console.print(menu_text, style="bold cyan")

def extract_history(db_path, browser):
    """Extrae el historial de navegación para navegadores basados en Chromium."""
    try:
        temp_db_path = f"{browser.replace(' ', '_').lower()}_history_temp.sqlite"
        with open(db_path, "rb") as src, open(temp_db_path, "wb") as dst:
            dst.write(src.read())

        conn = sqlite3.connect(temp_db_path)
        cursor = conn.cursor()

        query = "SELECT url, title, last_visit_time FROM urls ORDER BY last_visit_time DESC"
        cursor.execute(query)
        history = cursor.fetchall()

        history_formatted = []
        for entry in track(history, description=f"[bold cyan]⏳ Procesando historial de {browser}...[/bold cyan]"):
            url, title, timestamp = entry
            last_visit_time = datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=timestamp)
            history_formatted.append((url, title, last_visit_time))

        cursor.close()
        conn.close()
        os.remove(temp_db_path)

        return history_formatted
    except Exception as e:
        console.print(f"[bold red]❌ {str(e)}[/bold red]")
        return []

def get_firefox_history_path():
    """Obtiene la ruta correcta del historial de Firefox buscando el archivo places.sqlite en todos los perfiles."""
    base_path = BROWSERS["Firefox"]["path"]
    if not os.path.exists(base_path):
        return None
    # Obtener todas las carpetas dentro del directorio de perfiles
    profiles = [d for d in glob.glob(os.path.join(base_path, "*")) if os.path.isdir(d)]
    if not profiles:
        return None
    # Recorrer cada perfil y devolver la ruta si se encuentra places.sqlite
    for profile in profiles:
        candidate = os.path.join(profile, "places.sqlite")
        if os.path.exists(candidate):
            return candidate
    return None

def extract_firefox_history(db_path):
    """Extrae el historial de Firefox."""
    try:
        temp_db_path = "firefox_history_temp.sqlite"
        with open(db_path, "rb") as src, open(temp_db_path, "wb") as dst:
            dst.write(src.read())  # Copiar para evitar bloqueo

        conn = sqlite3.connect(temp_db_path)
        cursor = conn.cursor()

        query = """
            SELECT url, title, visit_date / 1000000 AS timestamp
            FROM moz_places
            JOIN moz_historyvisits ON moz_places.id = moz_historyvisits.place_id
            ORDER BY visit_date DESC
        """
        cursor.execute(query)
        history = cursor.fetchall()

        history_formatted = []
        for entry in track(history, description=f"[bold cyan]⏳ Procesando historial de Firefox...[/bold cyan]"):
            url, title, timestamp = entry
            last_visit_time = datetime.datetime(1970, 1, 1) + datetime.timedelta(seconds=timestamp)
            history_formatted.append((url, title, last_visit_time))

        cursor.close()
        conn.close()
        os.remove(temp_db_path)

        return history_formatted
    except Exception as e:
        console.print(f"[bold red]❌ {str(e)}[/bold red]")
        return []

def save_to_excel(history, browser):
    """Guarda el historial en un archivo Excel en la carpeta 'history'."""
    try:
        # Crear la carpeta "history" si no existe
        output_dir = "history"
        os.makedirs(output_dir, exist_ok=True)
        
        df = pd.DataFrame(history, columns=["URL", "Title", "Last Visit Time"])
        file_name = f"{browser.replace(' ', '_')}_history.xlsx"
        file_path = os.path.join(output_dir, file_name)
        df.to_excel(file_path, index=False)
        console.print(f"[bold green]✔ {BROWSERS[browser]['icon']} Historial de {browser} guardado en '{file_path}' ✅[/bold green]")
    except Exception as e:
        console.print(f"[bold red]❌ {str(e)}[/bold red]")

def main():
    """Función principal que permite al usuario seleccionar el navegador."""
    while True:
        show_menu()

        try:
            choice = int(Prompt.ask("\n🔹 [bold yellow]Ingresa el número correspondiente[/bold yellow]"))

            if choice == 0:
                console.print("\n[bold red]🚪 Saliendo del programa... ¡Hasta luego! 👋[/bold red]")
                sleep(1)
                break

            browser = list(BROWSERS.keys())[choice - 1]

            if browser == "Firefox":
                db_path = get_firefox_history_path()
                if not db_path or not os.path.exists(db_path):
                    console.print(f"[bold red]⚠ No se encontró el historial de {browser}.[/bold red]\n")
                    continue
                console.print(f"\n[bold blue]🔍 {BROWSERS[browser]['icon']} Extrayendo historial de {browser}...[/bold blue]")
                sleep(1)
                history = extract_firefox_history(db_path)
            else:
                db_path = BROWSERS[browser]["path"]
                if not os.path.exists(db_path):
                    console.print(f"[bold red]⚠ No se encontró el historial de {browser}.[/bold red]\n")
                    continue
                console.print(f"\n[bold blue]🔍 {BROWSERS[browser]['icon']} Extrayendo historial de {browser}...[/bold blue]")
                sleep(1)
                history = extract_history(db_path, browser)

            if history:
                save_to_excel(history, browser)
            else:
                console.print(f"[bold red]❌ No se pudo extraer el historial de {browser}.[/bold red]")

        except (ValueError, IndexError):
            console.print("[bold red]❌ Selección no válida. Intenta de nuevo.[/bold red]")

        again = Prompt.ask("\n🔄 [bold yellow]¿Quieres extraer otro historial? (s/n)[/bold yellow]", choices=["s", "n"], default="n")
        if again.lower() != "s":
            console.print("\n[bold red]🚪 Saliendo del programa... ¡Hasta luego! 👋[/bold red]")
            sleep(1)
            break

if __name__ == "__main__":
    main()
