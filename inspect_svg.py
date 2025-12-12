import win32com.client
import os

# --- הגדרות ---
# שימי כאן את הנתיב המדויק לקובץ ה-SVG הבעייתי שיצרת
# (תמצאי אותו בתיקיית Test_Images שלך)
SVG_PATH = r"C:\Users\yarde\OneDrive\Desktop\Auto_Print_System\Test_Images\logo_text2.svg" 

def inspect_svg():
    print(f"--- Inspecting SVG Structure: {SVG_PATH} ---")
    
    app = win32com.client.Dispatch("Illustrator.Application")
    app.UserInteractionLevel = -1 
    
    # פותחים מסמך זמני נקי
    doc = app.Documents.Add()
    
    try:
        # מכניסים את ה-SVG
        imported_group = doc.ActiveLayer.GroupItems.CreateFromFile(SVG_PATH)
        
        count = imported_group.PageItems.Count
        print(f"\nFound {count} items inside the SVG Group.")
        print("-" * 50)
        
        # עוברים על כל הפריטים ומדפיסים נתונים עליהם
        for i in range(1, count + 1):
            item = imported_group.PageItems(i)
            
            # בדיקת סוג
            info = f"Item {i}: Type={item.TypeName}"
            
            # בדיקת גודל (שטח)
            try:
                area = int(item.Width * item.Height)
                info += f", Area={area}"
            except: pass
            
            # בדיקת צבע (אם אפשרי)
            try:
                if item.TypeName == "PathItem":
                    if item.Filled:
                        c = item.FillColor
                        if c.TypeName == "RGBColor":
                            info += f", Color=RGB({int(c.Red)},{int(c.Green)},{int(c.Blue)})"
                        else:
                            info += f", ColorType={c.TypeName}"
                    else:
                        info += ", No Fill"
            except: pass
            
            print(info)
            
        print("-" * 50)
        print("Please copy this list and send it to Gemini!")
        
        # סוגרים את המסמך הזמני בלי לשמור
        doc.Close(2) # 2 = Do not save

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_svg()