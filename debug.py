import win32com.client
import os

def inspect_file():
    # 1. הגדרת הנתיב (כמו שעשינו במיין)
    base_dir = os.getcwd()
    template_path = os.path.join(base_dir, 'Simulations', 'Short.ai')
    
    print("--- מתחיל בדיקה ---")
    
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        app.UserInteractionLevel = -1
        
        # פתיחת הקובץ
        print(f"Opening: {template_path}")
        doc = app.Open(template_path)
        
        # 2. ניסיון למצוא את השכבה
        try:
            sim_layer = doc.Layers("Simulation")
            print("V Found layer: 'Simulation'")
        except:
            print("X ERROR: Could not find layer 'Simulation'")
            return

        # 3. ניסיון למצוא את הקבוצה
        try:
            shirt_group = sim_layer.GroupItems("Shirt")
            print("V Found group: 'Shirt'")
            
            # 4. הדפסת מה שיש בפנים - זה החלק החשוב!
            count = shirt_group.PageItems.Count
            print(f"\n--- Items inside 'Shirt' group ({count} items) ---")
            
            for i in range(1, count + 1):
                item = shirt_group.PageItems(i)
                # נדפיס את השם ואת הסוג של כל פריט
                try:
                    name = item.Name
                except:
                    name = "(No Name)"
                
                print(f"Item {i}: Name='{name}', Type='{item.TypeName}'")
                
        except:
            print("X ERROR: Could not find group 'Shirt' inside 'Simulation'")
            # הדפסה של מה כן יש בשכבה כדי שנבין
            print("Items in Simulation layer:")
            for item in sim_layer.PageItems:
                print(f" - {item.Name} ({item.TypeName})")

    except Exception as e:
        print(f"General Error: {e}")

if __name__ == "__main__":
    inspect_file()