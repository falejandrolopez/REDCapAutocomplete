from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import re

# ---------------- CONFIG ----------------
URL = "https://rcinvestigacion.ms.gba.gov.ar/surveys/?s=LXXRMH3EDDH9XKAT"
EXCEL_FILE = "plantilla_encuesta_v2.xlsx"

# ---------------- OPTIMIZED FUNCTIONS ----------------

def wait_and_interact(driver, by_selector, selector_value, timeout=10):
    """Actively waits for an element to be fully clickable."""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by_selector, selector_value))
    )

def write_fixed_input(driver, name_attribute, value):
    """Safely writes to inputs that have a fixed NAME attribute."""
    try:
        if pd.isna(value) or str(value).strip() == "":
            return
        field = wait_and_interact(driver, By.NAME, name_attribute)
        field.clear()
        field.send_keys(str(value))
    except Exception as e:
        print(f"[-] Error writing to fixed field (NAME='{name_attribute}'): {e}")

def write_redcap_date(driver, name_attribute, value):
    """Converts any Excel/Pandas date format to REDCap's DD/MM/YYYY format."""
    try:
        if pd.isna(value) or str(value).strip() == "":
            return
            
        if isinstance(value, pd.Timestamp) or hasattr(value, 'strftime'):
            formatted_date = value.strftime("%d/%m/%Y")
        else:
            val_str = str(value).strip()
            if re.match(r'\d{4}-\d{2}-\d{2}', val_str):
                parts = val_str.split('-')
                formatted_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
            elif re.match(r'\d{2}-\d{2}-\d{4}', val_str):
                formatted_date = val_str.replace('-', '/')
            else:
                formatted_date = val_str

        print(f"[*] Typing adapted date: {formatted_date} into '{name_attribute}'")
        field = wait_and_interact(driver, By.NAME, name_attribute)
        field.clear()
        field.send_keys(formatted_date)
        field.send_keys(Keys.TAB) 
    except Exception as e:
        print(f"[-] Error writing date to (NAME='{name_attribute}'): {e}")

def select_static_dropdown(driver, name_attribute, value):
    """Handles standard dropdown menus by visible text or value matching."""
    try:
        if pd.isna(value) or str(value).strip() == "":
            return
            
        clean_value = str(value).strip()
        print(f"[*] Attempting to select '{clean_value}' in field '{name_attribute}'...")

        select_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, name_attribute))
        )
        select = Select(select_elem)
        
        target_value_upper = clean_value.upper()
        for option in select.options:
            if target_value_upper in option.text.upper():
                try:
                    select.select_by_visible_text(option.text)
                    print(f"[+] Success! Selected '{option.text}' in '{name_attribute}' by text.")
                    return
                except:
                    pass

        try:
            select.select_by_value(clean_value)
            print(f"[+] Success! Selected ID '{clean_value}' in '{name_attribute}' by numeric value.")
            return
        except:
            pass

        print(f"[-] Could not find or select '{clean_value}' in dropdown NAME='{name_attribute}'")
    except Exception as e:
        print(f"[-] Error interacting with dropdown NAME='{name_attribute}': {e}")

def select_redcap_radio(driver, base_name, value):
    """Handles REDCap radio buttons by constructing the expected element name suffix."""
    try:
        if pd.isna(value) or str(value).strip() == "":
            return
            
        val_str = str(value).strip().upper()
        if val_str in ["SI", "TRUE"]: 
            clean_value = "1"
        elif val_str in ["NO", "FALSE"]: 
            clean_value = "0"
        else:
            try:
                clean_value = str(int(float(value))).strip()
            except ValueError:
                clean_value = val_str

        radio_name = f"{base_name}___radio"
        xpath = f"//input[@name='{radio_name}' and @value='{clean_value}']"
        
        radio_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].click();", radio_elem)
        time.sleep(0.1)
    except Exception as e:
        print(f"[-] Error selecting radio button for '{base_name}': {e}")     

def select_redcap_checkbox(driver, base_name, checkbox_code, value):
    """Handles REDCap checkboxes via custom attributes."""
    try:
        if pd.isna(value) or str(value).strip().upper() not in ["SI", "1", "TRUE"]:
            return
            
        checkbox_name = f"__chkn__{base_name}"
        str_code = str(checkbox_code).strip()
        xpath = f"//input[@type='checkbox' and @name='{checkbox_name}' and @code='{str_code}']"
        
        checkbox_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        if not checkbox_elem.is_selected():
            driver.execute_script("arguments[0].click();", checkbox_elem)
            time.sleep(0.1)
    except Exception as e:
        print(f"[-] Error checking box '{base_name}' with code '{checkbox_code}': {e}")          

# ---------------- MAIN SCRIPT ----------------

df = pd.read_excel(EXCEL_FILE)

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.page_load_strategy = 'eager'

driver = webdriver.Chrome(options=chrome_options)

for i, row in df.iterrows():
    print(f"\n=========================================")
    print(f" LOADING RECORD {i+1} OF {len(df)}")
    print(f"=========================================")
    driver.get(URL)
    time.sleep(2)  
    
    # [STOPWATCH] Initial time tracking
    start_time = time.time() 
    
    # -------- FIXED FIELDS --------
    write_fixed_input(driver, "nom_1", row["efector"])
    write_fixed_input(driver, "cod_1", row["codigo"])

    # -------- MUNICIPALITY --------
    select_static_dropdown(driver, "municipio", row["municipio"])
    time.sleep(1) 
    
    write_redcap_date(driver, "fecha_1", row["fecha"])

    # -------- SHIFT & ZONE --------
    select_redcap_radio(driver, "turno", row["turno_2"])
    select_redcap_radio(driver, "zona", row["zona_2"])
    select_redcap_radio(driver, "po", row["pueblos_orig_2"])

    # -------- PERSONAL DATA --------
    write_fixed_input(driver, "apellido", row["apellido"])
    write_fixed_input(driver, "nombre_s", row["nombre"])
    write_fixed_input(driver, "dni", row["dni"])
    write_redcap_date(driver, "fecha_de_nacimiento", row["fecha_nacimiento"])
    write_fixed_input(driver, "edad", row["edad"])
    
    select_static_dropdown(driver, "g_nero", row["genero"])
    select_static_dropdown(driver, "a_o_escolar_en_curso", row["anio_curso"])    
    write_fixed_input(driver, "direcci_n_calle_y_n_mero", row["direccion"])
    write_fixed_input(driver, "localidad", row["localidad"])    
    select_static_dropdown(driver, "cobertura_m_dica", row["cobertura_medica"])
    select_static_dropdown(driver, "auh", row["auh"])
    write_fixed_input(driver, "tel_fono_de_contacto", row["telefono"])

    # -------- HEALTH STATUS --------
    select_static_dropdown(driver, "control_de_salud_en_el_lti", row["control_salud"])

    # -------- MEDICAL HISTORY --------
    select_redcap_radio(driver, "broncoespasmos_a_repetici2", row["broncoespasmo_2"])
    select_redcap_radio(driver, "cadiopat_as_cong_nitas_o_d2", row["cardiopatias_2"])
    select_redcap_radio(driver, "convulsiones2", row["convulsiones_2"])
    select_redcap_radio(driver, "diabetes2", row["diabetes_2"])
    select_redcap_radio(driver, "usa_anteojos_o_tiene_indic", row["anteojos_2"])
    select_redcap_radio(driver, "muerte_s_bita_o_antes_de_l", row["muerte_subita_familiar_2"])
    select_redcap_radio(driver, "tuvo_internaciones_y_o_cir", row["internaciones_2"])

    # -------- MEDICATION --------
    select_static_dropdown(driver, "toma_medicaci_n_actualment", row["medicacion"])
    write_fixed_input(driver, "cuales", row["cuales"])
    write_fixed_input(driver, "existen", row["otros_antecedentes"])

    # -------- CLINICAL EXAMINATION --------
    select_static_dropdown(driver, "cardiovascular1", row["cardiovascular"])
    write_fixed_input(driver, "observac_card", row["orb_card"])
    select_static_dropdown(driver, "respiratorio2", row["respiratorio"])
    write_fixed_input(driver, "observac_resp_2", row["orb_resp"])
    select_static_dropdown(driver, "osteoarticular3", row["osteoarticular"])
    write_fixed_input(driver, "observac_oste_4", row["orb_osteo"])
    select_static_dropdown(driver, "parttes_blandas4", row["piel"])
    write_fixed_input(driver, "observac_pb_5", row["orb_piel"])
    select_static_dropdown(driver, "abdominal4", row["abdomen"])
    write_fixed_input(driver, "observac_abd_3", row["orb_abdo"])
    select_static_dropdown(driver, "del_habla5", row["habla"])
    write_fixed_input(driver, "observac_habla_4", row["orb_habla"])
    write_fixed_input(driver, "descripci_n_de_hallazgos_p", row["otras_alteraciones"])
    write_fixed_input(driver, "peso_kg", row["peso"])
    write_fixed_input(driver, "talla_cm", row["talla"])

    # -------- VACCINATION --------
    select_redcap_radio(driver, "esquema_completo_de_vacuna", row["esquema_vacunas_2"])
    select_redcap_radio(driver, "decide_vacunarse", row["decide_vacunarse_2"])
    select_static_dropdown(driver, "vacunas_que_se_apliacan_tr", row["triple_viral"])
    select_static_dropdown(driver, "vacunas_que_se_apliacan_tr_2", row["triple_bacteriana"])
    select_static_dropdown(driver, "vacunas_que_se_apliacan_tr_3", row["ipv"])
    select_static_dropdown(driver, "vacunas_que_se_apliacan_tr_4", row["varicela"])
    select_static_dropdown(driver, "vacunas_que_se_apliacan_tr_5", row["covid"])
    write_fixed_input(driver, "otras_vacunas", row["otras_vacunas"])
    
    # -------- OPHTHALMOLOGY --------
    select_redcap_radio(driver, "agudeza_visual_ojo_derecho", row["agudeza_od"])
    select_redcap_radio(driver, "agudeza_visual_ojo_derecho_2", row["agudeza_oi"])
    select_redcap_radio(driver, "alteraciones_del_ojo_exter", row["ojo_externo_2"])
    write_fixed_input(driver, "cuales2", row["ojo_externo_cuales"])
    select_redcap_checkbox(driver, "oftalmolog_a_indicar_todas", 1, row["fondo_ojo"])
    select_redcap_checkbox(driver, "oftalmolog_a_indicar_todas", 2, row["programa_vpa"])
    
    # -------- DENTISTRY --------
    select_redcap_radio(driver, "acciones_promopreventivas", row["acciones_promoprev_2"])
    select_redcap_radio(driver, "aplicacion_de_fl_or", row["fluor_2"])
    select_redcap_radio(driver, "alteraciones_del_crecimien", row["maxilares_2"])
    write_fixed_input(driver, "oltras_alteraciones_odont", row["otras_odonto"])
    
    # -------- FINAL DETAILS --------
    write_fixed_input(driver, "observaciones_indicar_toda", row["observaciones"])
    
    # [STOPWATCH] Tallying time spent filling out the form
    end_time = time.time()
    total_time = end_time - start_time
    
    # =========================================================================
    # ASSISTANT MODE: MANUAL SUBMISSION PAUSE
    # =========================================================================
    current_url = driver.current_url
    
    # Smooth scroll up a bit before printing terminal alerts
    driver.execute_script("window.scrollTo(0, 300);") 
    
    print("\n[!] Form field data populated automatically.")
    print(f"[⏱️] The script took {total_time:.2f} seconds to autofill the form.")
    print("[>] Please select the school manually on the browser and click SUBMIT...")
    
    try:
        # Actively wait for user interaction to submit and the page URL to change
        WebDriverWait(driver, 300).until(
            EC.url_changes(current_url)
        )
        print(f"[+] Perfect! Detected manual submission for record {i+1}.")
        time.sleep(1) 
        
    except Exception:
        print("[-] Timeout: Took too long to detect form submission.")
        input(">>> If you already submitted it, press ENTER here to move to the next row...")

# --- SAFETY ANCHOR ---
print("\n[+] Data loop execution completed successfully.")
input(">>> Press ENTER in this console window to close the Chrome browser instance permanently...")
driver.quit()
