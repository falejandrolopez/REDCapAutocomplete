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
ARCHIVO_EXCEL = "plantilla_encuesta_v2.xlsx"

# ---------------- FUNCIONES OPTIMIZADAS ----------------

def esperar_e_interactuar(driver, por_que, selector, tiempo=10):
    """Espera activamente a que un elemento sea completamente interactuable"""
    return WebDriverWait(driver, tiempo).until(
        EC.element_to_be_clickable((por_que, selector))
    )

def escribir_input_fijo(driver, name_atributo, valor):
    """Escribe de forma segura en los campos que tienen un atributo NAME fijo"""
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return
        campo = esperar_e_interactuar(driver, By.NAME, name_atributo)
        campo.clear()
        campo.send_keys(str(valor))
    except Exception as e:
        print(f"[-] Error al escribir en el campo fijo (NAME='{name_atributo}'): {e}")

def escribir_fecha_redcap(driver, name_atributo, valor):
    """Convierte cualquier formato de fecha de Excel/Pandas al formato DD/MM/YYYY de REDCap"""
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return
            
        if isinstance(valor, pd.Timestamp) or hasattr(valor, 'strftime'):
            fecha_formateada = valor.strftime("%d/%m/%Y")
        else:
            val_str = str(valor).strip()
            if re.match(r'\d{4}-\d{2}-\d{2}', val_str):
                partes = val_str.split('-')
                fecha_formateada = f"{partes[2]}/{partes[1]}/{partes[0]}"
            elif re.match(r'\d{2}-\d{2}-\d{4}', val_str):
                fecha_formateada = val_str.replace('-', '/')
            else:
                fecha_formateada = val_str

        print(f"[*] Escribiendo fecha adaptada: {fecha_formateada} en '{name_atributo}'")
        campo = esperar_e_interactuar(driver, By.NAME, name_atributo)
        campo.clear()
        campo.send_keys(fecha_formateada)
        campo.send_keys(Keys.TAB) 
    except Exception as e:
        print(f"[-] Error al escribir la fecha en (NAME='{name_atributo}'): {e}")

def seleccionar_dropdown_estatico(driver, name_atributo, valor):
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return
            
        valor_limpio = str(valor).strip()
        print(f"[*] Intentando cargar '{valor_limpio}' en el campo '{name_atributo}'...")

        select_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, name_atributo))
        )
        select = Select(select_elem)
        
        valor_buscado_up = valor_limpio.upper()
        for opcion in select.options:
            if valor_buscado_up in opcion.text.upper():
                try:
                    select.select_by_visible_text(opcion.text)
                    print(f"[+] ¡Éxito! Se seleccionó '{opcion.text}' en '{name_atributo}' por texto.")
                    return
                except:
                    pass

        try:
            select.select_by_value(valor_limpio)
            print(f"[+] ¡Éxito! Se seleccionó el ID '{valor_limpio}' en '{name_atributo}' por valor numérico.")
            return
        except:
            pass

        print(f"[-] No se pudo seleccionar '{valor_limpio}' en el select NAME='{name_atributo}'")
    except Exception as e:
        print(f"[-] Error al interactuar con el select NAME='{name_atributo}': {e}")

def seleccionar_radio_redcap(driver, name_base, valor):
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return
            
        val_str = str(valor).strip().upper()
        if val_str in ["SI", "TRUE"]: valor_limpio = "1"
        elif val_str in ["NO", "FALSE"]: valor_limpio = "0"
        else:
            try:
                valor_limpio = str(int(float(valor))).strip()
            except ValueError:
                valor_limpio = val_str

        name_radio = f"{name_base}___radio"
        xpath = f"//input[@name='{name_radio}' and @value='{valor_limpio}']"
        
        radio_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].click();", radio_elem)
        time.sleep(0.1)
    except Exception as e:
        print(f"[-] Error al marcar el radio para '{name_base}': {e}")     

def seleccionar_checkbox_redcap(driver, name_base, codigo_casilla, valor):
    try:
        if pd.isna(valor) or str(valor).strip().upper() not in ["SI", "1", "TRUE"]:
            return
            
        name_checkbox = f"__chkn__{name_base}"
        str_code = str(codigo_casilla).strip()
        xpath = f"//input[@type='checkbox' and @name='{name_checkbox}' and @code='{str_code}']"
        
        checkbox_elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        if not checkbox_elem.is_selected():
            driver.execute_script("arguments[0].click();", checkbox_elem)
            time.sleep(0.1)
    except Exception as e:
        print(f"[-] Error al tildar el checkbox '{name_base}' con código '{codigo_casilla}': {e}")          

# ---------------- SCRIPT PRINCIPAL ----------------

df = pd.read_excel(ARCHIVO_EXCEL)

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.page_load_strategy = 'eager'

driver = webdriver.Chrome(options=chrome_options)

for i, row in df.iterrows():
    print(f"\n=========================================")
    print(f" CARGANDO REGISTRO {i+1} DE {len(df)}")
    print(f"=========================================")
    driver.get(URL)
    time.sleep(2)  
    
    # [CRONÓMETRO] Tiempo inicial
    tiempo_inicio = time.time() 
    
    # -------- FIJOS --------
    escribir_input_fijo(driver, "nom_1", row["efector"])
    escribir_input_fijo(driver, "cod_1", row["codigo"])

    # -------- MUNICIPIO --------
    seleccionar_dropdown_estatico(driver, "municipio", row["municipio"])
    time.sleep(1) 
    
    escribir_fecha_redcap(driver, "fecha_1", row["fecha"])

    # -------- TURNO ZONA --------
    seleccionar_radio_redcap(driver, "turno", row["turno_2"])
    seleccionar_radio_redcap(driver, "zona", row["zona_2"])
    seleccionar_radio_redcap(driver, "po", row["pueblos_orig_2"])

    # -------- DATOS PERSONALES --------
    escribir_input_fijo(driver, "apellido", row["apellido"])
    escribir_input_fijo(driver, "nombre_s", row["nombre"])
    escribir_input_fijo(driver, "dni", row["dni"])
    escribir_fecha_redcap(driver, "fecha_de_nacimiento", row["fecha_nacimiento"])
    escribir_input_fijo(driver, "edad", row["edad"])
    
    seleccionar_dropdown_estatico(driver, "g_nero", row["genero"])
    seleccionar_dropdown_estatico(driver, "a_o_escolar_en_curso", row["anio_curso"])    
    escribir_input_fijo(driver, "direcci_n_calle_y_n_mero", row["direccion"])
    escribir_input_fijo(driver, "localidad", row["localidad"])    
    seleccionar_dropdown_estatico(driver, "cobertura_m_dica", row["cobertura_medica"])
    seleccionar_dropdown_estatico(driver, "auh", row["auh"])
    escribir_input_fijo(driver, "tel_fono_de_contacto", row["telefono"])

    # -------- SALUD --------
    seleccionar_dropdown_estatico(driver, "control_de_salud_en_el_lti", row["control_salud"])

    # -------- ANTECEDENTES --------
    seleccionar_radio_redcap(driver, "broncoespasmos_a_repetici2", row["broncoespasmo_2"])
    seleccionar_radio_redcap(driver, "cadiopat_as_cong_nitas_o_d2", row["cardiopatias_2"])
    seleccionar_radio_redcap(driver, "convulsiones2", row["convulsiones_2"])
    seleccionar_radio_redcap(driver, "diabetes2", row["diabetes_2"])
    seleccionar_radio_redcap(driver, "usa_anteojos_o_tiene_indic", row["anteojos_2"])
    seleccionar_radio_redcap(driver, "muerte_s_bita_o_antes_de_l", row["muerte_subita_familiar_2"])
    seleccionar_radio_redcap(driver, "tuvo_internaciones_y_o_cir", row["internaciones_2"])

    # -------- MEDICACION --------
    seleccionar_dropdown_estatico(driver, "toma_medicaci_n_actualment", row["medicacion"])
    escribir_input_fijo(driver, "cuales", row["cuales"])
    escribir_input_fijo(driver, "existen", row["otros_antecedentes"])

    # -------- EXAMEN CLÍNICO --------
    seleccionar_dropdown_estatico(driver, "cardiovascular1", row["cardiovascular"])
    escribir_input_fijo(driver, "observac_card", row["orb_card"])
    seleccionar_dropdown_estatico(driver, "respiratorio2", row["respiratorio"])
    escribir_input_fijo(driver, "observac_resp_2", row["orb_resp"])
    seleccionar_dropdown_estatico(driver, "osteoarticular3", row["osteoarticular"])
    escribir_input_fijo(driver, "observac_oste_4", row["orb_osteo"])
    seleccionar_dropdown_estatico(driver, "parttes_blandas4", row["piel"])
    escribir_input_fijo(driver, "observac_pb_5", row["orb_piel"])
    seleccionar_dropdown_estatico(driver, "abdominal4", row["abdomen"])
    escribir_input_fijo(driver, "observac_abd_3", row["orb_abdo"])
    seleccionar_dropdown_estatico(driver, "del_habla5", row["habla"])
    escribir_input_fijo(driver, "observac_habla_4", row["orb_habla"])
    escribir_input_fijo(driver, "descripci_n_de_hallazgos_p", row["otras_alteraciones"])
    escribir_input_fijo(driver, "peso_kg", row["peso"])
    escribir_input_fijo(driver, "talla_cm", row["talla"])

    # -------- VACUNACION --------
    seleccionar_radio_redcap(driver, "esquema_completo_de_vacuna", row["esquema_vacunas_2"])
    seleccionar_radio_redcap(driver, "decide_vacunarse", row["decide_vacunarse_2"])
    seleccionar_dropdown_estatico(driver, "vacunas_que_se_apliacan_tr", row["triple_viral"])
    seleccionar_dropdown_estatico(driver, "vacunas_que_se_apliacan_tr_2", row["triple_bacteriana"])
    seleccionar_dropdown_estatico(driver, "vacunas_que_se_apliacan_tr_3", row["ipv"])
    seleccionar_dropdown_estatico(driver, "vacunas_que_se_apliacan_tr_4", row["varicela"])
    seleccionar_dropdown_estatico(driver, "vacunas_que_se_apliacan_tr_5", row["covid"])
    escribir_input_fijo(driver, "otras_vacunas", row["otras_vacunas"])
    
    # -------- OFTALMO --------
    seleccionar_radio_redcap(driver, "agudeza_visual_ojo_derecho", row["agudeza_od"])
    seleccionar_radio_redcap(driver, "agudeza_visual_ojo_derecho_2", row["agudeza_oi"])
    seleccionar_radio_redcap(driver, "alteraciones_del_ojo_exter", row["ojo_externo_2"])
    escribir_input_fijo(driver, "cuales2", row["ojo_externo_cuales"])
    seleccionar_checkbox_redcap(driver, "oftalmolog_a_indicar_todas", 1, row["fondo_ojo"])
    seleccionar_checkbox_redcap(driver, "oftalmolog_a_indicar_todas", 2, row["programa_vpa"])
    
    # -------- ODONTO --------
    seleccionar_radio_redcap(driver, "acciones_promopreventivas", row["acciones_promoprev_2"])
    seleccionar_radio_redcap(driver, "aplicacion_de_fl_or", row["fluor_2"])
    seleccionar_radio_redcap(driver, "alteraciones_del_crecimien", row["maxilares_2"])
    escribir_input_fijo(driver, "oltras_alteraciones_odont", row["otras_odonto"]) # <- [CORREGIDO NOMBRE]
    
    # -------- FINAL --------
    escribir_input_fijo(driver, "observaciones_indicar_toda", row["observaciones"])
    
    # [CRONÓMETRO] Tiempo de finalización de tipeo
    tiempo_final = time.time()
    tiempo_total = tiempo_final - tiempo_inicio
    
    # =========================================================================
    # MODO ASISTENTE: ESPERA DE ENVÍO MANUAL
    # =========================================================================
    url_actual = driver.current_url
    
    # Ejecutamos el scroll al principio antes de los avisos para máxima fluidez
    driver.execute_script("window.scrollTo(0, 300);") 
    
    print("\n[!] Formulario completado automáticamente.")
    print(f"[⏱️] El script tardó {tiempo_total:.2f} segundos en llenar los datos.")
    print("[>] Por favor, elegí la escuela a mano y hacé clic en ENVIAR en el navegador...")
    
    try:
        # Espera activa hasta que hagas clic en Enviar y la URL cambie
        WebDriverWait(driver, 300).until(
            EC.url_changes(url_actual)
        )
        print(f"[+] ¡Perfecto! Detecté el envío manual del registro {i+1}.")
        time.sleep(1) 
        
    except Exception:
        print("[-] Pasó demasiado tiempo sin detectar el envío.")
        input(">>> Si ya lo enviaste, presioná ENTER acá para continuar con la siguiente fila...")

# --- ANCLA DE SEGURIDAD ---
print("\n[+] El bucle de carga terminó exitosamente.")
input(">>> Presioná ENTER en esta consola para cerrar la ventana de Chrome definitivamente...")
driver.quit()