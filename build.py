import os
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILD_DIR = ROOT / "build"
DIST_DIR = ROOT / "dist"

def clean_old_builds():
    print("Limpiando builds anteriores...")
    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)
        print("  - Eliminado /build")
    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
        print("  - Eliminado /dist")
    for spec in ROOT.glob("*.spec"):
        spec.unlink()
        print(f"  - Eliminado {spec.name}")

def run_pyinstaller():
    print("Compilando con PyInstaller...")
    main_script = str(ROOT / "src" / "main.py")
    
    # Argumentos principales segun requerimientos
    args = [
        "pyinstaller",
        "--noconfirm",
        "--onefile",
        "--noconsole",
        "--name", "Data-Travel",
        "--collect-all", "customtkinter",
        "--collect-all", "rapidfuzz",
        "--collect-all", "gspread",
        "--hidden-import", "google-api-python-client",
        "--hidden-import", "google.auth",
        "--hidden-import", "google.oauth2",
        main_script
    ]
    
    try:
        result = subprocess.run(args, cwd=str(ROOT))
        if result.returncode == 0:
            print("\nCompilacion exitosa!")
            exe_path = DIST_DIR / "Data-Travel.exe"
            if exe_path.exists():
                print(f"Ejecutable generado en: {exe_path}")
        else:
            print(f"\nError en la compilacion (codigo {result.returncode})")
    except FileNotFoundError:
        print("\nPyInstaller no encontrado. Ejecuta 'pip install pyinstaller' primero.")

if __name__ == "__main__":
    clean_old_builds()
    run_pyinstaller()
