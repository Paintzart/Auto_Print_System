import streamlit as st

import os
import json
import shutil

# --- תיקון: שינוי השם למילון החדש ---

from main import process_order, EXTENDED_COLOR_MAP 



# --- הגדרת נתיב זמני דינמי ---

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")



# יצירת התיקייה הזמנית אם היא לא קיימת

if not os.path.exists(UPLOAD_DIR):

    os.makedirs(UPLOAD_DIR)



def save_uploaded_file(uploaded_file):

    if uploaded_file is not None:

        file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)

        with open(file_path, "wb") as f:

            f.write(uploaded_file.getbuffer())

        return file_path

    return None



def get_color_value(selection):

    # התיקון: מחזירים את המילה "צבעוני" (מחרוזת) ולא None

    if selection == "צבעוני (ללא שינוי)":

        return "צבעוני" 

    return selection



st.set_page_config(page_title="מערכת הדמיות", layout="wide", page_icon="👕")

st.title("👕 מערכת הדמיות אוטומטית")



# אזור עליון - פרטי הזמנה

col1, col2, col3 = st.columns(3)

with col1:

    order_id = st.text_input("מספר הזמנה", value="1001")

with col2:

    product_type_heb = st.selectbox("סוג מוצר", ["חולצה", "סווטשירט", "קפוצון", "קפוצון עם רוכסן"])

with col3:

    # --- תיקון: שימוש במילון המורחב ---

    shirt_colors = list(EXTENDED_COLOR_MAP.keys())

    product_color = st.selectbox("צבע המוצר", shirt_colors)



prod_type_map = {"חולצה": "Shirt", "סווטשירט": "Sweater", "קפוצון": "Hoodie", "קפוצון עם רוכסן": "Zippered Hoodie"}

product_type = prod_type_map[product_type_heb]



st.markdown("---")



def create_input_section(title, key_prefix, size_options):

    st.subheader(title)

    exists = st.checkbox(f"יש הדפסה ב{title}?", key=f"{key_prefix}_exists")

    

    if exists:

        c1, c2, c3 = st.columns(3)

        with c1:

            size = st.selectbox("גודל / סוג", size_options, key=f"{key_prefix}_size")

        with c2:

            # --- תיקון: שימוש במילון המורחב ---

            color_options = ["צבעוני (ללא שינוי)"] + list(EXTENDED_COLOR_MAP.keys())

            color = st.selectbox("צבע ההדפס", color_options, key=f"{key_prefix}_color")

        with c3:

            uploaded_file = st.file_uploader(f"העלאת קובץ", type=['jpg', 'jpeg', 'png', 'svg'], key=f"{key_prefix}_file")

        

        return {

            'exists': True,

            'size': size,

            'color': get_color_value(color),

            'file': uploaded_file

        }

    else:

        return {'exists': False}



# הגדרת האזורים

front_data = create_input_section("צד קידמי", "F", ["סמל כיס", "A4", "A3"])

back_data = create_input_section("צד אחורי", "B", ["A4", "A3"])

rs_data = create_input_section("שרוול ימין", "RS", ["9 ס\"מ"])

ls_data = create_input_section("שרוול שמאל", "LS", ["9 ס\"מ"])



st.markdown("---")



if st.button("🚀 צור הדמיה והדפסה", type="primary"):

    if not order_id:

        st.error("חובה להזין מספר הזמנה")

    else:

        def map_category(ui_size):

            if ui_size == "סמל כיס": return "Pocket"

            if ui_size == "9 ס\"מ": return "Sleeve"

            return ui_size



        # בניית אובייקט ההזמנה

        order_obj = {

            'order_id': order_id,

            'product_type': product_type,

            'product_color_hebrew': product_color,

            'front': {

                'exists': front_data['exists'],

                'file': save_uploaded_file(front_data.get('file')),

                'category': map_category(front_data.get('size')),

                'prefix': 'F',

                'label': 'size_Front', 'heb': 'קידמי',

                'req_color_hebrew': front_data.get('color')

            },

            'back': {

                'exists': back_data['exists'],

                'file': save_uploaded_file(back_data.get('file')),

                'category': map_category(back_data.get('size')),

                'prefix': 'B',

                'label': 'size_Back', 'heb': 'אחורי',

                'req_color_hebrew': back_data.get('color')

            },

            'right_sleeve': {

                'exists': rs_data['exists'],

                'file': save_uploaded_file(rs_data.get('file')),

                'category': 'Sleeve',

                'prefix': 'RS',

                'label': 'size_RS', 'heb': 'שרוול ימין',

                'req_color_hebrew': rs_data.get('color')

            },

            'left_sleeve': {

                'exists': ls_data['exists'],

                'file': save_uploaded_file(ls_data.get('file')),

                'category': 'Sleeve',

                'prefix': 'LS',

                'label': 'size_LS', 'heb': 'שרוול שמאל',

                'req_color_hebrew': ls_data.get('color')

            }

        }



        # בדיקת תקינות

        valid = True

        for key in ['front', 'back', 'right_sleeve', 'left_sleeve']:

            if order_obj[key]['exists'] and not order_obj[key]['file']:

                st.error(f"חסר קובץ עבור {order_obj[key]['heb']}")

                valid = False

        

        if valid:

            with st.spinner('מעבד את ההזמנה...'):

                try:

                    process_order(order_obj)
                    st.balloons()
                    st.success(f"✅ ההזמנה {order_id} בוצעה בהצלחה!")
                    
                    # --- התיקון: חישוב השם הקצר גם לתצוגה ---
                    short_id_display = str(order_id)[-4:]
                    # ----------------------------------------

                    # הצגת הנתיב החדש
                    try:
                        with open('config.json', 'r', encoding='utf-8') as f:
                            config = json.load(f)
                            root_save_folder = config.get('save_folder_path', "Documents/Auto_Print_Output")
                    except:
                        root_save_folder = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output")
                    
                    # שימוש בשם הקצר בנתיב שהמשתמש רואה
                    final_save_path = os.path.join(root_save_folder, short_id_display)
                    
                    st.info(f"הקובץ נשמר בתיקייה: {final_save_path}")
                    # הצגת הנתיב החדש (ללא תאריך)

                    save_path = os.path.join(os.path.expanduser("~"), "Documents", "Auto_Print_Output", order_id)

                    st.info(f"הקובץ נשמר בתיקייה: {save_path}")

                    

                    # ניקוי תיקייה זמנית

                    if os.path.exists(UPLOAD_DIR):

                        shutil.rmtree(UPLOAD_DIR)

                        os.makedirs(UPLOAD_DIR)

                except Exception as e:

                    st.error(f"שגיאה: {e}")