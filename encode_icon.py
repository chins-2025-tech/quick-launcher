import base64

# --- 設定 ---
ICON_FILE_PATH = "icon.png"

# --- 実行 ---
try:
    with open(ICON_FILE_PATH, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')

    print("--- 以下をPythonコードに貼り付けてください ---")
    print(f"ICON_BASE64 = \"\"\"\\\n{encoded_string}\\\n\"\"\"")
    print("-----------------------------------------")

except FileNotFoundError:
    print(f"エラー: '{ICON_FILE_PATH}' が見つかりません。")
    