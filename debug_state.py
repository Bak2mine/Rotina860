import pyautogui
import os
import time

IMAGES = r"C:\Users\andre\Desktop\Zeon"
time.sleep(5)

images = [
    "btn_exportar.png",
    "btn_run.png",
    "btn_rotina860.png",
    "btn_entrar.png",
    "icon_winthor.png",
    "btn_conectar.png",
    "field_arquivo.png",
    "field_senha.png",
]

for image in images:
    for conf in [0.8, 0.7, 0.6, 0.5, 0.4, 0.3]:
        try:
            location = pyautogui.locateOnScreen(
                os.path.join(IMAGES, image), confidence=conf
            )
            if location:
                print(f"  MATCH: {image} at confidence {conf}")
                break
        except pyautogui.ImageNotFoundException:
            pass
    else:
        print(f"  NO MATCH: {image} even at 0.3")