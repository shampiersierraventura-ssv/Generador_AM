import streamlit as st
import pandas as pd
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

from docxtpl import DocxTemplate


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Generador de Ayuda Memoria",
    page_icon="📄",
    layout="wide"
)

MODO_VISIBLE = False
TIMEOUT_PAGINA = 30
TIMEOUT_ELEMENTO = 15
TIMEOUT_MODAL = 10
MAX_REINTENTOS = 2

CARPETA_INPUT = Path("Input")
CARPETA_SALIDA = Path("Output")
CARPETA_SALIDA.mkdir(exist_ok=True)

RUTA_PLANTILLA = CARPETA_INPUT / "plantillas" / "AM_plantilla_1.docx"


# ============================================================
# FUNCIONES GENERALES
# ============================================================

def limpiar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def crear_driver():
    options = webdriver.ChromeOptions()

    options.binary_location = "/usr/bin/chromium"

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")

    service = Service("/usr/bin/chromedriver")

    driver = Chrome(service=service, options=options)
    driver.set_page_load_timeout(TIMEOUT_PAGINA)
    driver.set_script_timeout(15)

    return driver


# ============================================================
# SCRAPING 1: DATOS FINANCIEROS SSI
# ============================================================

def cerrar_ventanas_emergentes(driver):
    try:
        driver.execute_script("""
            document.querySelectorAll('.modal.show').forEach(m => {
                m.style.display='none';
                m.classList.remove('show');
            });
            document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
            document.body.classList.remove('modal-open');
            document.body.style.overflow='';
            document.body.style.paddingRight='';
        """)
    except:
        pass


def esperar_modal_visible(driver, timeout=10):
    try:
        Wait(driver, timeout).until(
            EC.visibility_of_element_located((By.ID, "divResumenCont"))
        )

        time.sleep(0.5)

        tiene_contenido = driver.execute_script("""
            return document.getElementById('td_mtototal2_r') !== null &&
                   document.getElementById('val_pim_r') !== null;
        """)

        return tiene_contenido

    except:
        return False


def extraer_datos_modal(driver):
    datos = driver.execute_script("""
        function getText(id) {
            try {
                var elem = document.getElementById(id);
                if (!elem) return 'NO DISPONIBLE';
                var texto = elem.textContent || elem.innerText || '';
                return texto.trim() || 'NO DISPONIBLE';
            } catch(e) {
                return 'NO DISPONIBLE';
            }
        }

        return {
            nombre_inversion:       getText('td_nominv_r'),
            uf:                     getText('td_uf_r'),
            uei:                    getText('td_uei_r'),
            costo_total:            getText('td_mtototal2_r'),
            costo_viable_aprobado:  getText('td_mtoviab_r'),
            fecha_viabilidad:       getText('td_fecviab_r'),
            pim_2026:               getText('val_pim_r'),
            devengado_acum:         getText('val_efin_r'),
            devengado_2026:         getText('val_avan_r'),
            avance_financiero_acum: getText('por_avanacum_r'),
            costo_total_viable:     getText('td_totviab_r')
        };
    """)

    return datos


def obtener_datos_financieros(cui):
    driver = crear_driver()

    try:
        driver.get("https://ofi5.mef.gob.pe/ssi/")

        input_box = Wait(driver, TIMEOUT_ELEMENTO).until(
            EC.element_to_be_clickable((By.ID, "txt_cu"))
        )
        input_box.clear()
        input_box.send_keys(cui)

        btn_buscar = Wait(driver, TIMEOUT_ELEMENTO).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn_bus"))
        )
        btn_buscar.click()

        Wait(driver, TIMEOUT_ELEMENTO).until(
            EC.presence_of_element_located((By.ID, "td_cu"))
        )

        cerrar_ventanas_emergentes(driver)

        try:
            btn_resumen = Wait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "img[src*='resumen.png']"))
            )
            driver.execute_script("arguments[0].click();", btn_resumen)
        except:
            driver.execute_script("""
                var imgs = document.querySelectorAll('img[src*="resumen.png"]');
                if (imgs.length > 0) imgs[0].click();
            """)

        if not esperar_modal_visible(driver, TIMEOUT_MODAL):
            raise Exception("No cargó el resumen financiero.")

        datos = extraer_datos_modal(driver)

        df_financieros = pd.DataFrame([{
            "CUI": cui,
            "Nombre de la Inversión": datos["nombre_inversion"],
            "Unidad Formuladora (UF)": datos["uf"],
            "Unidad Ejecutora de Inversiones (UEI)": datos["uei"],
            "Costo Inversión Total (a)": datos["costo_total"],
            "Costo Viable/Aprobado (S/) (a)": datos["costo_viable_aprobado"],
            "Fecha Viabilidad/Aprobación": datos["fecha_viabilidad"],
            "PIM 2026 (c)": datos["pim_2026"],
            "Devengado Acumulado al 2026 (b)": datos["devengado_acum"],
            "Devengado 2026 (d)": datos["devengado_2026"],
            "Avance Financiero Acumulado (b/a)": datos["avance_financiero_acum"],
            "Costo Total Inversión Viable/Aprobado (a+b)": datos["costo_total_viable"],
        }])

        return df_financieros

    finally:
        driver.quit()


# ============================================================
# SCRAPING 2: F12B
# ============================================================

def obtener_f12b(cui):
    driver = crear_driver()

    try:
        url = f"https://ofi5.mef.gob.pe/inviertews/Repseguim/ResumF12B?codigo={cui}"
        driver.get(url)

        Wait(driver, TIMEOUT_ELEMENTO).until(
            EC.presence_of_element_located((By.ID, "situ_act"))
        )

        time.sleep(4)

        datos = driver.execute_script("""
            function getText(id) {
                try {
                    var elem = document.getElementById(id);
                    if (!elem) return 'NO DISPONIBLE';
                    var texto = elem.textContent || elem.innerText || '';
                    return texto.trim() || 'NO DISPONIBLE';
                } catch(e) {
                    return 'NO DISPONIBLE';
                }
            }

            return {
                situ_act:    getText('situ_act'),
                probl_restr: getText('probl_restr'),
                prox_pasos:  getText('prox_pasos')
            };
        """)

        df_f12b = pd.DataFrame([{
            "CUI": cui,
            "SITUACIÓN ACTUAL": datos["situ_act"],
            "PROBLEMAS / RESTRICCIONES": datos["probl_restr"],
            "PRÓXIMOS PASOS": datos["prox_pasos"],
        }])

        return df_f12b

    finally:
        driver.quit()


# ============================================================
# SCRAPING 3: AVANCE FÍSICO
# ============================================================

def obtener_avance_fisico(cui):
    driver = crear_driver()

    try:
        url = f"https://ofi5.mef.gob.pe/invierteWS/Repseguim/RepEstimac?codigo={cui}"
        driver.get(url)
        time.sleep(4)

        avance = driver.execute_script("""
            var texto = document.body.innerText;

            var match = texto.match(/Avance\\s*F[ií]sico[^0-9]*([0-9]+(?:[.,][0-9]+)?\\s*%?)/i);

            if (match) {
                return match[1];
            }

            var elem = document.getElementById('por_afis');
            if (elem) {
                return (elem.textContent || elem.innerText || '').trim();
            }

            return '0';
        """)

        return pd.DataFrame([{
            "CUI": cui,
            "% Avance Físico": avance
        }])

    finally:
        driver.quit()

# ============================================================
# GENERAR WORD
# ============================================================

def generar_ayuda_memoria(df_final, cui):
    if not RUTA_PLANTILLA.exists():
        raise FileNotFoundError(f"No se encontró la plantilla: {RUTA_PLANTILLA}")

    row = df_final.iloc[0]

    doc = DocxTemplate(str(RUTA_PLANTILLA))
    anio = "2026"

    context = {
        "nombre_pi": limpiar(row.get("Nombre de la Inversión")),
        "cui": limpiar(row.get("CUI")),
        "uf": limpiar(row.get("Unidad Formuladora (UF)")),
        "uei": limpiar(row.get("Unidad Ejecutora de Inversiones (UEI)")),
        "fecha_viable": limpiar(row.get("Fecha Viabilidad/Aprobación")),
        "monto_viable": limpiar(row.get("Costo Viable/Aprobado (S/) (a)")),
        "pim_2026": limpiar(row.get("PIM 2026 (c)")),
        "dev_2026": limpiar(row.get("Devengado 2026 (d)")),
        "av_fin": limpiar(row.get("Avance Financiero Acumulado (b/a)")),
        "av_fis": limpiar(row.get("% Avance Físico")),
        "f12_1": limpiar(row.get("SITUACIÓN ACTUAL")),
        "f12_2": limpiar(row.get("PROBLEMAS / RESTRICCIONES")),
        "f12_3": limpiar(row.get("PRÓXIMOS PASOS")),
    }

    nombre_archivo = f"AYUDA MEMORIA - {cui} - {anio} - OPMI - MINSA.docx"
    ruta_salida = CARPETA_SALIDA / nombre_archivo

    doc.render(context)
    doc.save(str(ruta_salida))

    return ruta_salida


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.title("📄 Generador de Ayuda Memoria - OPMI MINSA")

st.write("Ingrese un CUI para extraer información del MEF y generar automáticamente la ayuda memoria.")

cui = st.text_input("Ingrese el CUI:", placeholder="Ejemplo: 1234567")

if st.button("Generar ayuda memoria"):
    if not cui.strip():
        st.error("Debe ingresar un CUI.")
    else:
        cui = cui.strip()

        try:
            with st.spinner("Consultando datos financieros..."):
                df_financieros = obtener_datos_financieros(cui)

            with st.spinner("Consultando comentarios F12B..."):
                df_f12b = obtener_f12b(cui)

            with st.spinner("Consultando avance físico..."):
                df_av_fis = obtener_avance_fisico(cui)

            df_final = (
                df_financieros
                .merge(df_f12b.drop(columns="CUI"), left_index=True, right_index=True)
                .merge(df_av_fis.drop(columns="CUI"), left_index=True, right_index=True)
            )

            st.success("Datos obtenidos correctamente.")

            st.subheader("Vista previa de datos")
            st.dataframe(df_final)

            with st.spinner("Generando documento Word..."):
                ruta_docx = generar_ayuda_memoria(df_final, cui)

            st.success("Ayuda memoria generada correctamente.")

            with open(ruta_docx, "rb") as file:
                st.download_button(
                    label="📥 Descargar ayuda memoria",
                    data=file,
                    file_name=ruta_docx.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")

