import pyautogui
import pyperclip
import time
import datetime
import subprocess
import sys
import os
import pygetwindow as gw
import win32gui # type: ignore
import win32con # type: ignore
import glob
import logging
import traceback
import getpass
import keyring

#pip install -r requirements.txt

# ============================================================
# LOGGING SETUP
# ============================================================
log_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(log_dir, f"rotina860_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# CONFIG
# ============================================================

# Winthor AppController path (same for all users)
WINTHOR_LNK = r"C:\Program Files (x86)\GraphOn\AppController\AppController.exe"

if os.path.exists(WINTHOR_LNK):
    print(f"  ✓ Found AppController: {WINTHOR_LNK}")
else:
    print(f"  ✗ AppController.exe not found at: {WINTHOR_LNK}")
    sys.exit(1)

# Get credentials from Windows Credential Manager (keyring)
CREDENTIAL_SERVICE = "Rotina860-Winthor"

try:
    current_user = getpass.getuser()

    # Try to get password from credential manager
    WINTHOR_PASSWORD = keyring.get_password(CREDENTIAL_SERVICE, current_user)

    if WINTHOR_PASSWORD:
        print(f"  ✓ Found stored credentials for: {current_user}")
    else:
        # No stored credentials - prompt user to save them
        print(f"\n  No stored credentials found for {current_user}")
        print("  Enter your Winthor password to save it securely:")
        print("  Senha (will not be displayed): ", end='', flush=True)
        WINTHOR_PASSWORD = getpass.getpass("")

        # Save to credential manager for future use
        try:
            keyring.set_password(CREDENTIAL_SERVICE, current_user, WINTHOR_PASSWORD)
            print("  ✓ Password saved securely in Windows Credential Manager")
        except Exception as e:
            logger.warning(f"Could not save password to credential manager: {e}")
            print("  ⚠ Password will not be saved (credential manager error)")

except Exception as e:
    logger.error(f"Credential manager error: {e}")
    print(f"  ✗ Error accessing credential manager: {e}")
    sys.exit(1)

# Save folder mirrors script location but on V: drive
if getattr(sys, 'frozen', False):
    # Running as .exe - get actual exe location for file operations
    # Images are in temp folder, but we want to save files where the exe is
    IMAGES = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable)
    script_path = os.path.dirname(sys.executable)  # Directory where .exe is located
else:
    # Running as script - use script directory
    script_path = os.path.dirname(os.path.abspath(__file__))
    IMAGES = script_path

# Map V: to C: (V: is mapped to C: drive)
path_without_drive = os.path.splitdrive(script_path)[1]
SAVE_FOLDER = "C:" + path_without_drive  # Use C: directly instead of V:

# ============================================================
# CONFIDENCE LEVELS PER IMAGE
# ============================================================
CONFIDENCE = {
    "btn_exportar.png":  0.7,
    "btn_run.png":       0.75,
    "rotina_txtbx.png":  0.7,
    "btn_entrar.png":    0.9,
    "icon_winthor.png":  0.55,
    "btn_conectar.png":  0.7,
    "field_arquivo.png": 0.7,
    "field_senha.png":   0.5,
    "btn_save.png":      0.6,
    "btn_fechar.png":    0.6,
    "winthor_azul.png": 0.55,
}

# ============================================================
# SQL QUERIES
# ============================================================
SQL_QUERIES = {
    "bp": "SELECT CODPROD, DESCRICAO, CUSTOREP, DTULTALTCUSTOREP FROM PCPRODUT",
    "0":  "select p.codprod, p.descricao, e.qtpedida, e.codfilial, e.dtultpedcompra, p.dtultaltcom from pcprodut p inner join pcest e on (p.codprod = e.codprod) where e.qtpedida > 0",
    "f5": "SELECT p.codprod, p.descricao, e.CODFILIAL, e.QTPENDENTE from pcprodut p inner join pcest e on (p.codprod = e.codprod) WHERE e.CODFILIAL=5",
    "f2": "SELECT p.codprod, p.descricao, e.CODFILIAL, e.QTPENDENTE from pcprodut p inner join pcest e on (p.codprod = e.codprod) WHERE e.CODFILIAL=2",
    "f1": "SELECT p.codprod, p.descricao, e.CODFILIAL, e.QTPENDENTE from pcprodut p inner join pcest e on (p.codprod = e.codprod) WHERE e.CODFILIAL=1",
    "f":  "SELECT p.codprod, p.descricao, e.CODFILIAL, e.QTPENDENTE from pcprodut p inner join pcest e on (p.codprod = e.codprod) WHERE e.CODFILIAL IN (1,2,5)"
}

FILE_NAMES = {
    "bp": "BASE PRODUTOS",
    "0":  "PEDIDO COMPRAS",
    "f5": "QTD PEDIDA_RESERVADA filial5",
    "f2": "QTD PEDIDA_RESERVADA filial2",
    "f1": "QTD PEDIDA_RESERVADA filial1",
    "f":  "TODAS FILIAIS"
}

# Ask which query to run
print("Qual SQL?")
for key in SQL_QUERIES:
    print(f"  {key}: {FILE_NAMES[key]}")
print("  all: RUN ALL QUERIES (bp, 0, f, bp)")

choice = input("Escolhas (bp, 0, f5, f2, f1, f, all): ").strip()

valid_choices = list(SQL_QUERIES.keys()) + ["all"]
if choice not in valid_choices:
    print("Invalido")
    sys.exit(1)

# Determine run sequence

# ============================================================
# HELPERS
# ============================================================

def is_visible(image_name):
    confidence = CONFIDENCE.get(image_name, 0.7)
    image_path = os.path.join(IMAGES, image_name)
    try:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)
        return location is not None
    except pyautogui.ImageNotFoundException:
        return False

def find_and_click(image_name, timeout=30, double=False):
    images = image_name if isinstance(image_name, list) else [image_name]
    start = time.time()
    while time.time() - start < timeout:
        for img in images:
            confidence = CONFIDENCE.get(img, 0.7)
            image_path = os.path.join(IMAGES, img)
            try:
                location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if location:
                    if double:
                        pyautogui.doubleClick(location)
                    else:
                        pyautogui.click(location)
                    print(f"  ✓ Found and clicked: {img}")
                    return location
            except pyautogui.ImageNotFoundException:
                pass
        time.sleep(0.5)
    print(f"  ✗ Could not find: {image_name} after {timeout}s")
    sys.exit(1)

def find_and_type(image_name, text, timeout=30, secret=False):
    confidence = CONFIDENCE.get(image_name, 0.7)
    image_path = os.path.join(IMAGES, image_name)
    start = time.time()
    while time.time() - start < timeout:
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if location:
                pyautogui.click(location)
                time.sleep(0.3)
                pyautogui.hotkey("ctrl", "a")
                pyperclip.copy(text)
                pyautogui.hotkey("ctrl", "v")
                if not secret:
                    print(f"  ✓ Typed into: {image_name}")
                else:
                    print(f"  ✓ Typed (hidden) into: {image_name}")
                return location
        except pyautogui.ImageNotFoundException:
            pass
        time.sleep(0.5)
    print(f"  ✗ Could not find: {image_name} after {timeout}s")
    sys.exit(1)

def detect_current_state():
    checks = [
        ("btn_exportar.png", "export_dialog"),
        ("btn_run.png",      "rotina_860"),
        ("rotina_txtbx.png", "main_screen"),
        ("btn_entrar.png",   "login"),
        (["icon_winthor.png", "winthor_azul.png"], "totvs_cloud"),
        ("btn_conectar.png", "conectar"),
    ]
    for image, state in checks:
        if isinstance(image, list):
            result = any(is_visible(img) for img in image)
            print(f"  checking {image}: {result}")
        else:
            result = is_visible(image)
            print(f"  checking {image}: {result}")
        if result:
            print(f"  ✓ Detected current state: {state}")
            return state
    print("  ? No state detected")
    return "unknown"

def maximize_860(timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        hwnd = win32gui.FindWindow(None, "860 - Consulta SQL")
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            print("  ✓ 860 window maximized")
            return True
        time.sleep(1)
    print("  ✗ Could not maximize 860 window")
    return False

def run_query_and_export(sql_key):
    """Runs steps 6-10 for a given query key. Assumes rotina_860 state."""
    sql_query = SQL_QUERIES[sql_key]
    date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(SAVE_FOLDER, f"{FILE_NAMES[sql_key]}_{date_str}.xls")

    print(f"\n--- Running {FILE_NAMES[sql_key]} ---")

    # STEP 6: Clear SQL area and paste new query
    print("Step 6: Pasting SQL query...")
    pyautogui.click(screen_width // 2, screen_height // 2)
    pyautogui.click(screen_width // 2, screen_height // 2)
    time.sleep(2)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.5)
    pyperclip.copy(sql_query)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(5)

    # STEP 7: Run query
    print("Step 7: Running query...")
    pyautogui.press("f9")
    time.sleep(10)
    state = detect_current_state()
    print(state)

    # STEP 8: Click bottom floppy disk
    if state == "rotina_860":
        maximize_860()
        print("Step 8: Opening export dialog...")
        pyautogui.click(screen_width // 2, 50)
        time.sleep(3)
        bottom_left = (0, screen_height // 2, screen_width // 2, screen_height // 2)
        location = pyautogui.locateOnScreen(
            os.path.join(IMAGES, "btn_save.png"), confidence=0.6, region=bottom_left
        )
        pyautogui.click(location)
        time.sleep(5)
        state = detect_current_state()
        print(state)

    # STEP 9 + 10: Type filename and export
    if state == "export_dialog":
        print("Step 9: Setting filename...")
        logger.info("Step 9: Setting filename...")
        pyautogui.press("tab")
        time.sleep(2)
        pyautogui.press("delete")
        time.sleep(2)
        find_and_type("field_arquivo.png", filename, timeout=10)
        time.sleep(5)

        print("Step 10: Exporting...")
        logger.info("Step 10: Exporting...")
        find_and_click("btn_exportar.png", timeout=10)
        time.sleep(4)
        print(f"  [SUCCESS] Exported: {filename}")
        logger.info(f"Successfully exported: {filename}")

        # Close export dialog
        print("Closing export dialog...")
        logger.info("Closing export dialog...")
        find_and_click("btn_fechar.png", timeout=10)
        time.sleep(2)
    else:
        print(f"  [RETRY] Unexpected state at step 9/10: {state} - retrying...")
        logger.warning(f"Unexpected state at step 9/10: {state} - query may not have completed")
        # Don't exit - let the main loop retry

def get_todays_file(prefix):
    today = datetime.datetime.now().strftime("%Y%m%d")
    for drive in ["V:", "C:"]:
        print(f"  Checking drive: {drive}")
        folder = drive + path_without_drive
        pattern = os.path.join(folder, f"{prefix}_{today}_*.xls")
        print(f"  Looking for: {pattern}")
        files = glob.glob(pattern)
        print(f"  Found: {files}")
        if files:
            latest = max(files, key=os.path.getmtime)
            print(f"  ✓ Today's file already exists: {os.path.basename(latest)}")
            return latest
    return None

if choice == "all":
    # Run all queries: bp, 0, f
    run_sequence = ["bp", "0", "f"]
    print("  Running all queries: bp, 0, f")
elif choice == "f":
    existing = get_todays_file(FILE_NAMES["f"])
    if existing:
        print(f"  Skipping f — already exported today")
        run_sequence = []
    else:
        run_sequence = ["f"]
else:
    run_sequence = [choice]

print(f"  run_sequence: {run_sequence}")

# ============================================================
# MAIN FLOW
# ============================================================
try:
    logger.info("=" * 80)
    logger.info("ROTINA 860 STARTED")
    logger.info(f"Choice: {choice}")
    logger.info(f"Run sequence: {run_sequence}")
    logger.info("=" * 80)

    # Skip opening Winthor entirely if all files already exist
    if choice == "f" and not run_sequence:
        logger.info("All files already exported today, skipping to merge...")
        print("All files already exported today, skipping to merge...")
        import merge_filiais
        merge_filiais.run(SAVE_FOLDER)
        logger.info("✅ All done!")
        print("\n✅ All done!")
        sys.exit(0)

    print(f"  run_sequence: {run_sequence}")
    logger.info(f"run_sequence: {run_sequence}")

    screen_width, screen_height = pyautogui.size()
    logger.info(f"Screen size: {screen_width}x{screen_height}")

    # Keep looping through states until we reach rotina_860
    max_attempts = 10
    attempt = 0

    while attempt < max_attempts:
        attempt += 1
        logger.info(f"Attempt {attempt}: Checking state...")
        time.sleep(3)
        state = detect_current_state()
        print(state)
        logger.info(f"Current state: {state}")

        # STEP 1: Open Winthor
        if state == "unknown":
            time.sleep(5)
            print("Step 1: Opening Winthor...")
            logger.info("Step 1: Opening Winthor...")
            subprocess.Popen(["cmd", "/c", "start", "", WINTHOR_LNK], shell=True)
            time.sleep(10)
            continue

        # STEP 2: Click Conectar
        if state == "conectar":
            print("Step 2: Clicking Conectar...")
            logger.info("Step 2: Clicking Conectar...")
            find_and_click("btn_conectar.png", timeout=15)
            time.sleep(10)
            continue

        # STEP 3: Double click WinThor icon
        if state == "totvs_cloud":
            print("Step 3: Double clicking WinThor icon...")
            logger.info("Step 3: Double clicking WinThor icon...")
            find_and_click(["icon_winthor.png", "winthor_azul.png"], timeout=15, double=True)
            time.sleep(10)
            continue

        # STEP 4: Type password and click Entrar
        if state == "login":
            print("Step 4: Logging in...")
            logger.info("Step 4: Logging in...")
            time.sleep(4)
            pyperclip.copy(WINTHOR_PASSWORD)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.3)
            pyautogui.press("enter")
            time.sleep(12)
            continue

        # STEP 5: Open Rotina 860
        if state == "main_screen":
            print("Step 5: Opening Rotina 860...")
            logger.info("Step 5: Opening Rotina 860...")
            pyperclip.copy("860")
            time.sleep(2)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(5)
            continue

        # STEPS 6-10: Run each query in sequence
        if state == "rotina_860":
            logger.info("Reached rotina_860 state, starting queries...")
            for sql_key in run_sequence:
                logger.info(f"Running query: {sql_key}")
                run_query_and_export(sql_key)

            # Run merge based on choice
            logger.info(f"Starting merge for choice: {choice}")
            import merge_filiais

            if choice == "all":
                print("\nMerging all queries into master file...")
                logger.info("Merging all queries: bp, 0, f...")
                merge_filiais.run_bp(SAVE_FOLDER)
                merge_filiais.run_pedidos(SAVE_FOLDER)
                merge_filiais.run(SAVE_FOLDER, "TODAS FILIAIS")
            elif choice == "f":
                print("\nMerging all filiais into master file...")
                logger.info("Merging all filiais into master file...")
                merge_filiais.run(SAVE_FOLDER, "TODAS FILIAIS")
            elif choice == "f1":
                print("\nMerging Filial 1 into master file...")
                logger.info("Merging Filial 1 into master file...")
                merge_filiais.run(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial1")
            elif choice == "f2":
                print("\nMerging Filial 2 into master file...")
                logger.info("Merging Filial 2 into master file...")
                merge_filiais.run(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial2")
            elif choice == "f5":
                print("\nMerging Filial 5 into master file...")
                logger.info("Merging Filial 5 into master file...")
                merge_filiais.run(SAVE_FOLDER, "QTD PEDIDA_RESERVADA filial5")
            elif choice == "bp":
                print("\nMerging Base Produtos into master file...")
                logger.info("Merging Base Produtos into master file...")
                merge_filiais.run_bp(SAVE_FOLDER)
            elif choice == "0":
                print("\nMerging Pedidos Compra into master file...")
                logger.info("Merging Pedidos Compra into master file...")
                merge_filiais.run_pedidos(SAVE_FOLDER)

            print("\n[DONE] All done!")
            logger.info("[DONE] Rotina 860 completed successfully!")
            print("\nPress any key to exit...")
            input()
            break

        if state == "export_dialog":
            print("  ℹ️  In export dialog, waiting for state change...")
            logger.info("In export_dialog, checking again...")
            continue

        print(f"  ⏳ Waiting for state change (attempt {attempt}/{max_attempts})...")
        logger.info(f"Waiting for state change (attempt {attempt}/{max_attempts})...")

    if attempt >= max_attempts:
        print(f"  ✗ Could not reach rotina_860 state after {max_attempts} attempts")
        logger.error(f"Could not reach rotina_860 state after {max_attempts} attempts")
        print("\nPress any key to exit...")
        input()
        sys.exit(1)

except Exception as e:
    error_msg = f"ERROR: {str(e)}\n{traceback.format_exc()}"
    logger.error(error_msg)
    print(f"\n❌ {error_msg}")
    print(f"\n📋 Log file saved to: {log_file}")
    print("\nPress any key to exit...")
    input()
    sys.exit(1)
