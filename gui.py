import streamlit as st
import os
import shutil

# --- ייבוא הפונקציות המעודכנות ---
# יש לוודא שהקבצים נמצאים באותה תיקייה או בנתיב שפייתון יודע למצוא
try:
    from print_automation import run_illustrator_split 
except ImportError:
    st.error("הקובץ print_automation.py חסר!")
    def run_illustrator_split(*args): return []

try:
    from photoshop_automation import run_photoshop_action
except ImportError:
    st.error("הקובץ photoshop_automation.py חסר!")
    def run_photoshop_action(*args): yield "DONE", "קובץ חסר"

# --- יבוא נתונים נוספים ---
try:
    # אם יש קובץ main.py או config.py שמכיל את EXTENDED_COLOR_MAP
    from main import EXTENDED_COLOR_MAP 
except ImportError:
    # ברירת מחדל אם הקובץ לא קיים
    EXTENDED_COLOR_MAP = {"שחור": "#000000", "לבן": "#FFFFFF", "אדום": "#FF0000"}

# --- הגדרות ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")
if not os.path.exists(UPLOAD_DIR): os.makedirs(UPLOAD_DIR)

def save_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

# ==========================================
# UI
# ==========================================
st.set_page_config(page_title="מערכת הדמיות ודפוס", layout="wide", page_icon="🖨️")
st.title("🖨️ מערכת ניהול הדמיות וקבצי דפוס")

tab1, tab2 = st.tabs(["👕 יצירת הדמיה", "✂️ פיצול והכנה לדפוס (מלא)"])

# --- טאב 1 (הדמיות) - נשאר ללא שינוי ---
with tab1:
    col1, col2, col3 = st.columns(3)
    with col1: order_id = st.text_input("מספר הזמנה", value="1001", key="sim_order_id")
    with col2: product_type_heb = st.selectbox("סוג מוצר", ["חולצה", "סווטשירט", "קפוצון", "קפוצון עם רוכסן"])
    with col3: product_color = st.selectbox("צבע המוצר", list(EXTENDED_COLOR_MAP.keys()))
    st.markdown("---")
    st.info("לשונית זו מפעילה את הקובץ main.py")

# --- טאב 2 (התהליך המאוחד) ---
with tab2:
    st.header("✨ תהליך אוטומטי מלא: אילוסטרייטור + פוטושופ")
    st.caption("התהליך כולל: פיצול קבצים, ניקוי שכבות, בדיקת צבע, ויצירת ערוץ ספוט לבן.")
    
    # 1. קלטים בסיסיים
    col_input, col_file = st.columns(2)
    with col_input:
        split_order_id = st.text_input("מספר הזמנה", value="", key="split_order_id")
    with col_file:
        source_pdf = st.file_uploader("העלה קובץ PDF/AI מקור", type=['pdf', 'ai'], key="source_pdf")
    
    st.markdown("---")
    
    # 2. בחירת הגדרות לוגו (לפני שמתחילים!)
    st.subheader("⚙️ הגדרות לוגו (עבור ספוט לבן)")
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        contract_choice = st.radio("בחר עובי לוגו:", 
                                   ["לוגו רגיל/עבה (כיווץ 2px)", "לוגו דק/עדין (כיווץ 1px)"], 
                                   index=0)
    
    # המרת הבחירה למספר
    contract_px = 2 if "2px" in contract_choice else 1

    st.markdown("---")

    # 3. כפתור ההפעלה
    if st.button("🚀 בצע תהליך מלא (Illustrator + Photoshop)", type="primary"):
        if not split_order_id or not source_pdf:
            st.error("נא להזין מספר הזמנה ולהעלות קובץ.")
        else:
            temp_pdf_path = save_uploaded_file(source_pdf)
            
            # איזור תצוגה
            st.info("מתחיל תהליך... נא לא לגעת במקלדת ובעכבר.")
            main_progress = st.progress(0)
            status_text = st.empty()
            
            final_folder = None
            files_list = []
            
            try:
                # ==========================
                # שלב א': אילוסטרייטור
                # ==========================
                status_text.text("🟠 שלב 1/2: מפעיל אילוסטרייטור (פיצול וניקוי)...")
                
                ill_runner = run_illustrator_split(temp_pdf_path, split_order_id)
                
                for data in ill_runner:
                    if isinstance(data[0], str) and data[0] == "DONE":
                        final_folder, files_list = data[1]
                    else:
                        # עדכון פרוגרס בר (0% עד 50% מהתהליך הכולל)
                        prog, txt = data
                        main_progress.progress(int(prog * 0.5 * 100)) 
                        status_text.text(f"Illustrator: {txt}")

                # ==========================
                # שלב ב': פוטושופ
                # ==========================
                if files_list: # רק אם אילוסטרייטור יצר קבצים
                    status_text.text("🔵 שלב 2/2: מפעיל פוטושופ (יצירת ספוט לבן)...")
                    
                    # *********** שימו לב: כאן נכנסת רשימת הקבצים המלאה ***********
                    ps_runner = run_photoshop_action(files_list, contract_px) 
                    
                    for data in ps_runner:
                        if isinstance(data[0], str) and data[0] == "DONE":
                            pass # סיימנו
                        else:
                            # עדכון פרוגרס בר (50% עד 100% מהתהליך הכולל)
                            prog, txt = data
                            combined_prog = 0.5 + (prog * 0.5)
                            main_progress.progress(combined_prog)
                            status_text.text(f"Photoshop: {txt}")
                    
                    # סיום מוצלח
                    main_progress.progress(100)
                    st.balloons()
                    st.success(f"✅ התהליך הושלם בהצלחה!")
                    if final_folder:
                        st.write(f"📂 הקבצים נשמרו בתיקייה: `{final_folder}`")
                    st.write(f"📄 קבצים שטופלו: {', '.join([os.path.basename(f) for f in files_list])}")
                
                else:
                    st.warning("אילוסטרייטור סיים אך לא נוצרו קבצים (אולי השכבות היו ריקות?), ולכן פוטושופ לא הופעל.")

            except Exception as e:
                st.error(f"❌ שגיאה במהלך התהליך: {e}")
            
            # ניקוי קובץ זמני
            if os.path.exists(temp_pdf_path):
                try: os.remove(temp_pdf_path)
                except: pass