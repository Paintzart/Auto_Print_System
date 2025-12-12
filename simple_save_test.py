import os
import win32com.client
import sys

# הגדרת קידוד כדי לראות עברית ולוגים כמו שצריך
sys.stdout.reconfigure(encoding='utf-8')

def save_active_doc_as_pdf(output_path):
    """
    פונקציה לשמירת המסמך הפעיל באילוסטרייטור כ-PDF באיכות דפוס.
    מבוססת על הפתרון שנבדק ונמצא תקין (ללא ViewPdfAfterSaving).
    
    Args:
        output_path (str): הנתיב המלא לשמירת הקובץ (כולל .pdf)
    
    Returns:
        bool: True אם הצליח, False אם נכשל
    """
    try:
        # 1. התחברות לאילוסטרייטור
        app = win32com.client.GetActiveObject("Illustrator.Application")
        
        if app.Documents.Count == 0:
            print("X שגיאה: לא נמצאו מסמכים פתוחים לשמירה.")
            return False

        doc = app.ActiveDocument
        
        # 2. יצירת אובייקט הגדרות PDF
        pdf_options = win32com.client.Dispatch("Illustrator.PDFSaveOptions")
        
        # 3. הגדרת איכות דפוס (Press Quality)
        try:
            # הדרך המועדפת: שימוש ב-Preset מוכן של אדובי
            pdf_options.PDFPreset = "[Press Quality]"
            print("V הוגדר Preset: [Press Quality]")
        except Exception:
            # גיבוי למקרה שה-Preset לא נמצא (נדיר, אבל ליתר ביטחון)
            print("! אזהרה: לא נמצא Preset, משתמש בהגדרות ידניות.")
            pdf_options.Compatibility = 7  # Acrobat 8
            pdf_options.PreserveEditability = False
            pdf_options.Optimization = True

        # *** קריטי: לא לגעת ב-ViewPdfAfterSaving כי זה גורם לקריסה ***
        
        # 4. ביצוע השמירה
        print(f"-> שומר קובץ ל: {output_path}")
        doc.SaveAs(output_path, pdf_options)
        
        print("V השמירה בוצעה בהצלחה.")
        return True

    except Exception as e:
        print(f"X שגיאה קריטית בשמירה: {e}")
        return False

# --- דוגמה לשימוש (Main) ---
if __name__ == "__main__":
    print("--- בדיקת פונקציית השמירה החדשה ---")
    
    # נתיב זמני לבדיקה על שולחן העבודה
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    test_file = os.path.join(desktop, "final_test_print_ready.pdf")
    
    # הרצת הפונקציה
    success = save_active_doc_as_pdf(test_file)
    
    if success:
        print("\nבדיקה עברה בהצלחה! אפשר להעתיק את הפונקציה לקוד הראשי.")
    else:
        print("\nהבדיקה נכשלה.")
        
    input("\nלחץ Enter לסיום...")