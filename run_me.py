import requests
import os
import time
import sys
# --- רשימת הקבצים המלאה למשיכה מה-GitHub שלך ---
FILES_TO_UPDATE = {
    "run_me.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/run_me.py", # הוא מעדכן את עצמו!
    "main.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/main.py",
    "vectorizer_ops.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/vectorizer_ops.py",
    "server.js": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/server.js",
    "prepare_print.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/prepare_print.py",
    "package.json": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/package.json",
    "package-lock.json": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/package-lock.json", # חדש!
    "requirements.txt": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/requirements.txt", # חדש!
    "sidebar_logic.jsx": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/sidebar_logic.jsx",
    "sidebar_numorder_logic.jsx": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/sidebar_numorder_logic.jsx",
    "variable_print.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/variable_print.py",
    "illustrator_ops.py": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/illustrator_ops.py",
    "run_merge_batch.jsx": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/run_merge_batch.jsx",
    "config.example.json": "https://raw.githubusercontent.com/Paintzart/Auto_Print_System/refs/heads/main/config.example.json",
}
def update_files():
    print("--- בודק עדכונים מהענן (GitHub) ---")
    updated_count = 0
    for filename, url in FILES_TO_UPDATE.items():
        try:
            print(f"בודק: {filename}...")
            response = requests.get(url)
            if response.status_code == 200:
                current_content = ""
                if os.path.exists(filename):
                    with open(filename, "r", encoding="utf-8") as f:
                        current_content = f.read()
                if response.text.strip() != current_content.strip():
                    with open(filename, "w", encoding="utf-8", newline='\n') as f:
                        f.write(response.text)
                    print(f"✅ עודכן: {filename}")
                    updated_count += 1
                else:
                    print(f"⚡ {filename} מעודכן.")
            else:
                print(f"⚠️ שגיאה בהורדת {filename} (קוד {response.status_code})")
        except Exception as e:
            print(f"❌ שגיאה בעדכון {filename}: {e}")
    if updated_count > 0:
        print(f"\n--- סיימנו! {updated_count} קבצים עודכנו. ---")
    else:
        print("\n--- הכל מעודכן. ---")
def run_software():
    print("\n🚀 מפעיל את שרת האוטומציה...")
    if os.path.exists("server.js"):
        os.system("node server.js")
    else:
        print("❌ שגיאה: server.js חסר!")
        input("לחץ על Enter...")
if __name__ == "__main__":
    try:
        update_files()
        time.sleep(1)
        run_software()
    except Exception as e:
        print(f"Error: {e}")
        input()
