import streamlit as st
import win32com.client
import os

def clean_and_save_part(app, source_path, save_path, target_layer_name, target_artboard_name):
    # פתיחת הקובץ
    try:
        doc = app.Open(source_path)
    except Exception as e:
        st.error(f"❌ שגיאה בפתיחת קובץ: {e}")
        return False

    try:
        # ============================================================
        # שלב 1: זיהוי השכבה ובדיקה האם היא באמת מלאה
        # ============================================================
        target_layer = None
        
        # חיפוש השכבה (לולאה מ-1 עד הסוף - אילוסטרייטור עובד עם 1-based ב-COM)
        for i in range(1, doc.Layers.Count + 1):
            lyr = doc.Layers(i)
            # השוואת שמות מנורמלת (בלי רווחים, אותיות קטנות)
            if lyr.Name.strip().lower() == target_layer_name.strip().lower():
                target_layer = lyr
                break
        
        # אם השכבה לא קיימת - סגור ודלג
        if target_layer is None:
            # st.warning(f"דילוג: שכבה {target_layer_name} לא נמצאה.")
            doc.Close(2)
            return False

        # --- הבדיקה הקשוחה לתוכן ---
        # אנחנו סופרים כמה אובייקטים יש בשכבה.
        # הערה: אם יש לך קווי עזר (Guides) או תיבות טקסט ריקות - זה ייחשב כפריט!
        item_count = target_layer.PageItems.Count
        
        if item_count == 0:
            st.warning(f"⚠️ מדלג על {target_layer_name}: השכבה קיימת אך ריקה (0 פריטים).")
            doc.Close(2)
            return False
            
        # ============================================================
        # שלב 2: מחיקת כל השכבות האחרות
        # ============================================================
        # רצים מהסוף להתחלה כדי לא לשבור אינדקסים
        for i in range(doc.Layers.Count, 0, -1):
            lyr = doc.Layers(i)
            # אם זו לא השכבה שלנו - למחוק!
            if lyr.Name.strip().lower() != target_layer_name.strip().lower():
                try:
                    lyr.Delete()
                except:
                    # אם אי אפשר למחוק (נעול וכו') - לפחות להסתיר
                    lyr.Visible = False
            else:
                # את השכבה שלנו לוודא שרואים
                lyr.Visible = True
                lyr.Locked = False

        # ============================================================
        # שלב 3: מחיקת כל הדפים (Artboards) המיותרים
        # ============================================================
        
        # א. מציאת האינדקס של הדף הרצוי
        target_ab_index = -1
        # משתמשים ב-range של פייתון (0 based) אבל הגישה לאובייקט תהיה (i+1)
        for i in range(doc.Artboards.Count):
            # שים לב: השימוש ב-(i+1) קריטי ב-Win32Com מול אילוסטרייטור
            ab = doc.Artboards(i+1) 
            if target_artboard_name.lower() in ab.Name.lower():
                target_ab_index = i # שומרים אינדקס 0-based ללוגיקה שלנו
                break
        
        if target_ab_index == -1:
            st.error(f"❌ לא נמצא עמוד בשם שמכיל: {target_artboard_name}")
            doc.Close(2)
            return False

        # ב. הופכים את הדף הרצוי לפעיל (כדי שלא נמחק אותו בטעות ושאילוסטרייטור יתפקס עליו)
        doc.Artboards.SetActiveArtboardIndex(target_ab_index)
        
        # ג. מחיקת כל השאר
        # רצים מהסוף להתחלה (חשוב מאוד במחיקת דפים!)
        initial_count = doc.Artboards.Count
        for i in range(initial_count, 0, -1):
            # האינדקס הנוכחי מול האינדקס הרצוי (צריך המרה כי שמרנו 0-based למעלה)
            # אבל באילוסטרייטור האינדקס הפעיל הוא 0-based... מבלבל, לכן נבדוק לפי שם או זהות
            
            # דרך בטוחה יותר: בדיקה אם זה הדף הפעיל
            is_active = (i - 1 == doc.Artboards.GetActiveArtboardIndex())
            
            if not is_active:
                try:
                    doc.Artboards(i).Delete()
                except Exception as e:
                    # אי אפשר למחוק אם נשאר רק עמוד אחד
                    pass

        # ============================================================
        # שלב 4: שמירה
        # ============================================================
        pdf_options = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
        pdf_options.PDFPreset = "[Press Quality]"
        
        # מכיוון שמחקנו את כל הדפים האחרים, נשאר רק עמוד מס' 1
        pdf_options.ArtboardRange = "1"
        
        # יצירת תיקייה
        folder = os.path.dirname(save_path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        doc.SaveAs(save_path, pdf_options)
        
        # st.success(f"✅ נשמר: {os.path.basename(save_path)}")
        
    except Exception as e:
        st.error(f"❌ שגיאה כללית בקובץ {target_layer_name}: {e}")
        doc.Close(2)
        return False
    
    doc.Close(2) # סגירה ללא שמירה על המקור
    return True

def process_order(source_file_path, order_number):
    # הגדרות נתיבים
    base_output_path = r"C:\Users\Yarde\Documents\Auto_Print_Output"
    
    # יצירת הנתיב: Output -> Last 4 Digits -> קבצי הדפסה
    last_4 = str(order_number)[-4:]
    order_folder = os.path.join(base_output_path, last_4)
    print_files_folder = os.path.join(order_folder, "קבצי הדפסה")
    
    if not os.path.exists(print_files_folder):
        os.makedirs(print_files_folder)
        print(f"Created folder: {print_files_folder}")

    # התחברות לאילוסטרייטור פעם אחת
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        app.UserInteractionLevel = -1 # ביטול חלונות קופצים (אזהרות פונטים וכו')
    except:
        print("Error: Illustrator not open.")
        return

    # הגדרת העבודות (סיומת, שם שכבה, שם משטח עבודה)
    # שימי לב: בשרוולים - יש לנו את אותו Artboard אבל שכבות שונות!
    jobs = [
        {"suffix": "PF", "layer": "Print_Front", "artboard": "Print_Front"},
        {"suffix": "PB", "layer": "Print_Back", "artboard": "Print_Back"},
        {"suffix": "PL", "layer": "Print_Left_Sleeve", "artboard": "Print_Sleeves"},
        {"suffix": "PR", "layer": "Print_Right_Sleeve", "artboard": "Print_Sleeves"},
    ]

    print(f"Processing Order {last_4}...")

    for job in jobs:
        filename = f"{last_4}_{job['suffix']}.pdf"
        full_save_path = os.path.join(print_files_folder, filename)
        
        # קריאה לפונקציה שעושה את העבודה המלוכלכת
        clean_and_save_part(
            app, 
            source_file_path, 
            full_save_path, 
            job["layer"], 
            job["artboard"]
        )

    # החזרת התראות למצב רגיל
    app.UserInteractionLevel = 1 
    print("Done.")

# --- אזור הרצה ---
# הזיני כאן את הנתיב לקובץ ה-PDF/AI שלך שאת רוצה לבדוק
my_input_file = r"C:\Users\Yarde\Downloads\Order_Example.pdf" # דוגמה
my_order_num = "109876"

# process_order(my_input_file, my_order_num)