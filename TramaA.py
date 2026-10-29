import os
import io
import re
import zipfile
import pandas as pd
import streamlit as st

# =========================================================
# CONFIGURACIÓN DE LA PÁGINA STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Sistema SETI-PRES - Consola Única SUSALUD",
    page_icon="🏥",
    layout="wide"
)

# =========================================================
# VALORES POR DEFECTO PARA LA TABLA A (RECURSOS)
# =========================================================
if "recursos_base" not in st.session_state:
    st.session_state.recursos_base = {
        "3554": [3, 4, 0, 2, 1, 0, 3, 1, 0, 0, 0, 2, 0, 4, 0, 1],
        "3555": [2, 2, 0, 1, 1, 0, 2, 1, 0, 0, 0, 1, 0, 2, 0, 0],
        "3556": [2, 2, 0, 1, 0, 0, 2, 0, 0, 0, 0, 1, 0, 2, 0, 0],
        "3557": [1, 2, 0, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 1, 0, 0],
        "3558": [2, 3, 0, 1, 0, 0, 2, 1, 0, 0, 0, 1, 0, 3, 0, 0],
        "6907": [2, 2, 0, 1, 1, 0, 2, 0, 0, 0, 0, 1, 0, 2, 0, 0],
    }

CAMPOS_RECURSOS = [
    "Consultorios Físicos",
    "Consultorios Funcionales",
    "Camas Hospitalarias",
    "Total de Médicos",
    "Médicos SERUMS",
    "Médicos Residentes",
    "Enfermeras",
    "Odontólogos",
    "Psicólogos",
    "Nutricionistas",
    "Tecnólogos Médicos",
    "Obstetrices",
    "Farmacéuticos",
    "Auxiliares / Técnicos",
    "Otros Profesionales",
    "Ambulancias",
]

# =========================================================
# FUNCIONES BÁSICAS Y MAPEO DE DATOS
# =========================================================
def limpiar_codigo_renipress(valor):
    if pd.isna(valor):
        return ""
    val_str = str(valor).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.lstrip("0")


def obtener_grupo_etario(edad_y):
    try:
        y = int(float(edad_y))
    except (ValueError, TypeError):
        return None

    if y < 1:
        return 1
    elif y < 5:
        return 2
    elif y < 10:
        return 3
    elif y < 15:
        return 4
    elif y < 20:
        return 5
    elif y < 25:
        return 6
    elif y < 30:
        return 7
    elif y < 35:
        return 8
    elif y < 40:
        return 9
    elif y < 45:
        return 10
    elif y < 50:
        return 11
    elif y < 55:
        return 12
    elif y < 60:
        return 13
    elif y < 65:
        return 14
    else:
        return 15


def normalizar_sexo(sexo):
    if pd.isna(sexo):
        return None
    s = str(sexo).strip().upper()
    if s in ["M", "MASCULINO", "1", "HOMBRES", "HOMBRE"]:
        return "1"
    elif s in ["F", "FEMENINO", "2", "MUJERES", "MUJER"]:
        return "2"
    return None


def estandarizar_cie10(codigo):
    if pd.isna(codigo):
        return ""
    c = str(codigo).strip().upper().replace(".", "")
    if c.startswith("Z298"):
        return "Z29.8"
    if len(c) == 4 and re.match(r"^[A-Z][0-9A-Z]{3}$", c):
        return f"{c[:3]}.{c[3]}"
    elif len(c) == 5 and re.match(r"^[A-Z][0-9A-Z]{4}$", c):
        return f"{c[:3]}.{c[3]}"
    return c


def es_cie10_morbilidad(codigo):
    if pd.isna(codigo):
        return False
    c = estandarizar_cie10(codigo)
    c_limpio = c.replace(".", "")
    return bool(re.match(r"^[A-Z][0-9A-Z]{3}$", c_limpio))


def cargar_y_normalizar_excel(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
    except Exception:
        df_raw = pd.read_excel(uploaded_file, engine="xlrd", header=None)

    header_row = 0
    keywords_cabecera = [
        "RENIPRES", "NUMDOC", "DNI", "PACIENTE",
        "CIE10", "EESS", "TIPODX", "ID_UPS", "FECPARTO"
    ]

    for idx, row in df_raw.iloc[:30].iterrows():
        valores_fila = [str(val).upper().strip() for val in row.values if pd.notna(val)]
        coincidencias = sum(1 for kw in keywords_cabecera if any(kw in val for val in valores_fila))
        if coincidencias >= 2:
            header_row = idx
            break

    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row)
    except Exception:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, engine="xlrd", header=header_row)

    df.columns = [str(col).strip().upper() for col in df.columns]

    alias_map = {
        "RENIPRESS": ["RENIPRESS", "RENIPRE", "COD_IPRESS", "ESTABLECIMIENTO", "ID_ESTABLECIMIENTO", "RENAES"],
        "NUMDOC": ["RHNUMDOC", "NUMDOC", "NUM_DOC", "NRO_DOC", "DOCUMENTO", "DNI", "PACIENTE_DOC", "NUMERO_DOCUMENTO"],
        "SEXO": ["SEXO", "SEXO_PACIENTE", "GENERO"],
        "EDAD_Y": ["EDAD_Y", "EDAD", "EDAD_ANOS", "EDAD_A", "EDAD_PACIENTE"],
        "CODIGO": ["CODIGO", "CIE10", "CIE_10", "COD_CIE10", "DIAGNOSTICO", "COD_DIAGNOSTICO", "ITEM_CODIGO", "CODDIG1"],
        "TIPODX": ["TIPODX", "TIPO_DX", "TIPO_DIAG", "TIPO_DIAGNOSTICO", "TIP_DX", "TIPO"],
        "IDCITA": ["IDCITA", "ID_CITA", "IDCITA_HIS", "ID_ATENCION", "FECHA_ATENCION", "FECHA"],
        "UPS": ["ID_UPS", "UPS", "UPS_ATENCION", "COD_UPS"],
        "SERVICIO": ["SERVICIO", "DESC_UPS", "NOMBRE_SERVICIO"],
        "FECPARTO": ["FECPARTO", "FECHA_PARTO"],
    }

    renombrar = {}
    columnas_existentes = list(df.columns)

    for estandar, aliases in alias_map.items():
        if estandar not in columnas_existentes:
            for alias in aliases:
                coincidencias = [col for col in columnas_existentes if alias in col]
                if coincidencias:
                    renombrar[coincidencias[0]] = estandar
                    break

    if renombrar:
        df = df.rename(columns=renombrar)

    return df


# =========================================================
# LÓGICA DE GENERACIÓN DE TRAMAS
# =========================================================
def procesar_tabla_a(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_a = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)
    lista_ordenada = sorted(codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        recursos = st.session_state.recursos_base.get(reni, [0] * 16)
        recursos_str = "|".join(str(val) for val in recursos)
        linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{recursos_str}"
        lineas_a.append(linea)

    return lineas_a


def procesar_tabla_b1(df_base, ipress_matriz, periodo):
    df = df_base.copy()
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if "PROFESION" in df.columns:
        df["ES_MEDICO"] = df["PROFESION"].astype(str).str.upper().str.contains("MEDIC|CMP", regex=True)
    elif "PERSONAL" in df.columns:
        df["ES_MEDICO"] = df["PERSONAL"].astype(str).str.upper().str.contains("MEDIC|CMP|MED", regex=True)
    else:
        df["ES_MEDICO"] = False

    if "IDCITA" in df.columns and df["IDCITA"].notna().any():
        df["ID_ATENCION_UNICA"] = df["IDCITA"].astype(str).str.strip()
    elif "FECHA" in df.columns:
        df["ID_ATENCION_UNICA"] = df["NUMDOC"].astype(str).str.strip() + "_" + df["FECHA"].astype(str).str.strip()
    else:
        df["ID_ATENCION_UNICA"] = df.index.astype(str)

    lineas_b1 = []
    lista_renipress = sorted(df["RENIPRESS_STR"].unique(), key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df[df["RENIPRESS_STR"] == reni]

        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                sub_df = df_reni[(df_reni["SEXO_PROC"] == sexo) & (df_reni["GRUPO_EDAD"] == grupo)]
                if sub_df.empty:
                    continue

                df_med = sub_df[sub_df["ES_MEDICO"]]
                df_no_med = sub_df[~sub_df["ES_MEDICO"]]

                atenciones_med = df_med["ID_ATENCION_UNICA"].nunique() if not df_med.empty else 0
                atenciones_no_med = df_no_med["ID_ATENCION_UNICA"].nunique() if not df_no_med.empty else 0
                atendidos_unicos = sub_df["NUMDOC"].astype(str).str.strip().nunique()

                linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{sexo}|{grupo}|{atenciones_med}|{atenciones_no_med}|{atendidos_unicos}"
                lineas_b1.append(linea)

    return lineas_b1


def procesar_tabla_b2(df_base, ipress_matriz, periodo):
    df_b2 = df_base[df_base["CIE_PROC"].apply(es_cie10_morbilidad)].copy()
    df_b2 = df_b2[~df_b2["CIE_PROC"].str.startswith("C")]
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if df_b2.empty:
        return [], 0

    j00x_g1_h = df_b2[
        (df_b2["CIE_PROC"].str.replace(".", "", regex=False) == "J00X") &
        (df_b2["GRUPO_EDAD"] == 1) &
        (df_b2["SEXO_PROC"] == "1")
    ]
    total_j00x = len(j00x_g1_h)

    lineas_b2 = []
    lista_renipress = sorted(df_b2["RENIPRESS_STR"].unique(), key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df_b2[df_b2["RENIPRESS_STR"] == reni]

        agrupado = df_reni.groupby(["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"]).size().reset_index(name="TOTAL_ATENDIDOS")
        agrupado = agrupado.sort_values(by=["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"])

        for _, fila in agrupado.iterrows():
            lineas_b2.append(
                f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{fila['SEXO_PROC']}|{int(fila['GRUPO_EDAD'])}|{fila['CIE_PROC']}|{fila['TOTAL_ATENDIDOS']}"
            )

    return lineas_b2, total_j00x


def procesar_tabla_c1(df_base, ipress_matriz, periodo, solo_filtrar_ups=True, lista_todos_renipress=None):
    df = df_base.copy()
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if solo_filtrar_ups and "UPS" in df.columns:
        condicion = df["UPS"].astype(str).str.contains("300303|301601", regex=True)
        df_emg = df[condicion].copy()
    else:
        df_emg = df.copy()

    if not df_emg.empty:
        if "IDCITA" in df_emg.columns and df_emg["IDCITA"].notna().any():
            df_emg["ID_VISITA"] = df_emg["IDCITA"].astype(str).str.strip()
        elif "FECHA" in df_emg.columns:
            df_emg["ID_VISITA"] = df_emg["NUMDOC"].astype(str).str.strip() + "_" + df_emg["FECHA"].astype(str).str.strip()
        else:
            df_emg["ID_VISITA"] = df_emg.index.astype(str)

        df_emg["DNI_PACIENTE"] = df_emg["NUMDOC"].astype(str).str.strip().str.upper()

    lineas_c1 = []
    lista_renipress = sorted(lista_todos_renipress, key=lambda x: int(x) if x.isdigit() else x) if lista_todos_renipress else sorted(df_base["RENIPRESS_STR"].unique(), key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df_emg[df_emg["RENIPRESS_STR"] == reni] if not df_emg.empty and "RENIPRESS_STR" in df_emg.columns else pd.DataFrame()

        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                if df_reni.empty:
                    total_atenciones = "NE_0001"
                    total_atendidos = "NE_0001"
                else:
                    sub_df = df_reni[(df_reni["SEXO_PROC"] == sexo) & (df_reni["GRUPO_EDAD"] == grupo)]
                    if sub_df.empty:
                        total_atenciones = 0
                        total_atendidos = 0
                    else:
                        total_atenciones = sub_df["ID_VISITA"].nunique()
                        total_atendidos = sub_df["DNI_PACIENTE"].nunique()

                linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{sexo}|{grupo}|{total_atenciones}|{total_atendidos}"
                lineas_c1.append(linea)

    return lineas_c1


def procesar_tabla_c2(df_base, ipress_matriz, periodo, solo_filtrar_ups=True):
    df = df_base.copy()
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if solo_filtrar_ups and "UPS" in df.columns:
        condicion = df["UPS"].astype(str).str.contains("300303|301601", regex=True)
        df_emg = df[condicion].copy()
    else:
        df_emg = df.copy()

    df_emg = df_emg[df_emg["CIE_PROC"].apply(es_cie10_morbilidad)].copy()

    if df_emg.empty:
        return []

    df_emg["DNI_PACIENTE"] = df_emg["NUMDOC"].astype(str).str.strip().str.upper()

    lineas_c2 = []
    lista_renipress = sorted(df_emg["RENIPRESS_STR"].unique(), key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df_emg[df_emg["RENIPRESS_STR"] == reni]

        agrupado = df_reni.groupby(["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"])["DNI_PACIENTE"].nunique().reset_index(name="TOTAL_ATENDIDOS")
        agrupado = agrupado.sort_values(by=["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"])

        for _, fila in agrupado.iterrows():
            linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{fila['SEXO_PROC']}|{int(fila['GRUPO_EDAD'])}|{fila['CIE_PROC']}|{fila['TOTAL_ATENDIDOS']}"
            lineas_c2.append(linea)

    return lineas_c2


def procesar_tabla_d1(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_d1 = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)
    lista_ordenada = sorted(codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001"
        lineas_d1.append(linea)

    return lineas_d1


def procesar_tabla_d2(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_d2 = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)
    lista_ordenada = sorted(codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x)

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{sexo}|{grupo}|NE_0001|NE_0001"
                lineas_d2.append(linea)

    return lineas_d2


def procesar_tabla_e(codigos_renipress_lista, ipress_matriz, periodo, df_principal=None, uploaded_seem=None):
    lineas_e = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)
    lista_ordenada = sorted(codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x)

    df_fuente = None
    if uploaded_seem is not None:
        try:
            filename = uploaded_seem.name.lower()
            if filename.endswith((".xlsx", ".xls")):
                df_fuente = pd.read_excel(uploaded_seem)
            else:
                df_fuente = pd.read_csv(uploaded_seem, sep=None, engine="python", encoding="latin-1", on_bad_lines="skip")
            df_fuente.columns = [str(c).upper().strip() for c in df_fuente.columns]
        except Exception:
            df_fuente = None

    if df_fuente is None and df_principal is not None and "FECPARTO" in df_principal.columns:
        df_fuente = df_principal.copy()

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        reni_clean = reni.lstrip("0")

        df_reni_partos = pd.DataFrame()
        if df_fuente is not None and not df_fuente.empty:
            col_ren_fuente = next((c for c in df_fuente.columns if "RENIPRESS" in c or "RENIPRE" in c), None)
            if col_ren_fuente:
                df_fuente["REN_LIMPIO"] = df_fuente[col_ren_fuente].apply(limpiar_codigo_renipress)
                df_sub = df_fuente[df_fuente["REN_LIMPIO"] == reni_clean]
            else:
                df_sub = df_fuente.copy()

            col_fec_parto = next((c for c in df_sub.columns if "FECPARTO" in c or "FECHA_PARTO" in c), None)
            if col_fec_parto:
                df_reni_partos = df_sub[df_sub[col_fec_parto].notna()].copy()

        if not df_reni_partos.empty:
            col_tipo = next((c for c in df_reni_partos.columns if "TIPO" in c or "CESA" in c), None)
            col_diag = next((c for c in df_reni_partos.columns if "COD" in c or "DIAG" in c), None)

            def get_tipo_parto(row):
                val = str(row.get(col_tipo, "")) if col_tipo else ""
                val_diag = str(row.get(col_diag, "")) if col_diag else ""
                if "CESA" in val.upper() or val in ["2", "81", "82"] or val_diag.startswith("O82"):
                    return "2"
                return "1"

            df_reni_partos["TIPO_P"] = df_reni_partos.apply(get_tipo_parto, axis=1)
            df_reni_partos["COMP_P"] = "1"

            agrupado = df_reni_partos.groupby(["TIPO_P", "COMP_P"]).size().reset_index(name="CANT_PARTOS")

            for _, fila in agrupado.iterrows():
                tipo_p = fila["TIPO_P"]
                comp_p = fila["COMP_P"]
                cant_p = int(fila["CANT_PARTOS"])
                vivos = cant_p
                muertos = 0
                tot_nac = vivos + muertos

                linea = f"{periodo}|{ipress_fmt}|{ipress_matriz_fmt}|{tipo_p}|{comp_p}|{cant_p}|{tot_nac}|{vivos}|{muertos}"
                lineas_e.append(linea)
        else:
            linea = f"{periodo}|{ipress_fmt}|{ipress_matriz_fmt}|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001"
            lineas_e.append(linea)

    return lineas_e


# =========================================================
# MODAL DE CONFIGURACIÓN DE TABLA A
# =========================================================
@st.dialog("Configuración de Recursos - Tabla A", width="large")
def abrir_modal_tabla_a(codigos_validos):
    st.markdown("### RECURSOS EN SALUD (TABLA A)")
    ipress_sel = st.selectbox("Seleccione IPRESS a editar:", options=codigos_validos)

    if ipress_sel:
        valores_actuales = st.session_state.recursos_base.get(ipress_sel, [0] * 16)
        nuevos_valores = []

        col1, col2 = st.columns(2)
        for idx, campo in enumerate(CAMPOS_RECURSOS):
            target_col = col1 if idx < 8 else col2
            with target_col:
                val = st.number_input(
                    f"{campo}:",
                    min_value=0,
                    value=int(valores_actuales[idx]),
                    key=f"modal_rec_{ipress_sel}_{idx}"
                )
                nuevos_valores.append(val)

        if st.button("💾 Guardar Cambios IPRESS", type="primary", use_container_width=True):
            st.session_state.recursos_base[ipress_sel] = nuevos_valores
            st.success(f"Recursos actualizados correctamente para la IPRESS {ipress_sel}.")
            st.rerun()


# =========================================================
# INTERFAZ PRINCIPAL Y LAYOUT
# =========================================================
def main():
    # -----------------------------------------------------
    # BARRA LATERAL CON DISEÑO IDÉntICO DE ESCRITORIO
    # -----------------------------------------------------
    st.sidebar.markdown("### ⚙️ **SETI-PRES**")
    st.sidebar.markdown("**SUSALUD - Módulo de Tramas**")
    st.sidebar.markdown("---")

    codigos_input_raw = st.sidebar.text_input("Lista RENIPRESS rápida:", "3554, 3555, 3556, 3557, 3558, 6907")
    codigos_lista = [limpiar_codigo_renipress(x) for x in codigos_input_raw.split(",") if limpiar_codigo_renipress(x)]

    if st.sidebar.button("Configurar Tabla A", use_container_width=True):
        if codigos_lista:
            abrir_modal_tabla_a(codigos_lista)
        else:
            st.sidebar.error("Ingrese al menos un código RENIPRESS.")

    # CSS para forzar el pie de página exactamente en la base inferior con diseño exacto
    st.sidebar.markdown(
        """
        <style>
            section[data-testid="stSidebar"] > div:nth-child(1) {
                display: flex;
                flex-direction: column;
                height: 100vh;
                box-sizing: border-box;
            }
            .sidebar-footer {
                margin-top: auto;
                padding-bottom: 20px;
            }
            .custom-card {
                background-color: #1d4ed8;
                color: #ffffff;
                padding: 12px;
                border-radius: 8px;
                text-align: center;
                font-size: 13px;
                line-height: 1.5;
                font-family: sans-serif;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Tarjeta idéntica a la versión de escritorio
    st.sidebar.markdown(
        """
        <div class="sidebar-footer">
            <div class="custom-card">
                Elaborado por el Ingº Sistemas<br>
                <b>Carlos Sauñe Yupanqui</b><br>
                CIP 209420<br>
                Estado: En línea<br>
                Versión 2026.1
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # -----------------------------------------------------
    # PANEL PRINCIPAL
    # -----------------------------------------------------
    st.title("Consola Integrada de Generación de Tramas")
    st.markdown("---")

    # 1. Parámetros del Envío
    st.markdown("### 1. Parámetros del Envío")
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        periodo = st.text_input("Periodo (AAAAMM):", value="202607")
    with col2:
        ipress_decl = st.text_input("UGIPRESS Matriz:", value="00003554")
    with col3:
        txt_ipress_lista = st.text_input(
            "RENIPRESS a Incluir:",
            value="00003554, 00003555, 00003556, 00003557, 00003558, 00006907"
        )

    codigos_validos = [limpiar_codigo_renipress(x) for x in txt_ipress_lista.split(",") if limpiar_codigo_renipress(x)]

    # 2. Origen de Datos
    st.markdown("### 2. Origen de Datos (Excel HIS / Partos)")
    archivo_his = st.file_uploader("Seleccione el archivo Excel principal del HIS:", type=["xlsx", "xls"])
    archivo_seem = st.file_uploader("Opcional: Archivo SEEM / Partos independiente:", type=["xlsx", "xls", "csv", "txt"])

    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        var_deduplicar = st.checkbox("Descartar duplicados en B2 por Cita/Atención (IDCITA) o Paciente", value=True)
    with col_chk2:
        var_modo_emergencia = st.checkbox("Generar C1 y C2 con TODO el Excel (Marcar si el Excel es Exclusivo de Emergencias)", value=False)

    st.markdown("---")

    # Botón Principal de Procesamiento
    if st.button("🚀 GENERAR PAQUETE DE TRAMAS SETI (A, B1, B2, C1, C2, D1, D2, E)", type="primary", use_container_width=True):
        if not archivo_his:
            st.warning("⚠️ Seleccione un archivo de Excel HIS válido para continuar.")
        else:
            with st.spinner("Cargando y procesando paquete completo de tramas..."):
                try:
                    df = cargar_y_normalizar_excel(archivo_his)
                    cols_requeridas = ["RENIPRESS", "NUMDOC", "SEXO", "EDAD_Y", "CODIGO", "TIPODX"]
                    faltantes = [c for c in cols_requeridas if c not in df.columns]

                    if faltantes:
                        st.error(f"Estructura no coincide. Faltan las siguientes columnas: {', '.join(faltantes)}")
                        st.write("Columnas detectadas:", list(df.columns[:15]))
                        return

                    df_proc = df.copy()
                    set_codigos_validos = set(codigos_validos)

                    df_proc["RENIPRESS_STR"] = df_proc["RENIPRESS"].apply(limpiar_codigo_renipress)
                    df_proc = df_proc[df_proc["RENIPRESS_STR"].isin(set_codigos_validos)]

                    df_proc = df_proc[df_proc["TIPODX"].astype(str).str.strip().str.upper().str.startswith("D")]

                    df_proc["SEXO_PROC"] = df_proc["SEXO"].apply(normalizar_sexo)
                    df_proc = df_proc[df_proc["SEXO_PROC"].notna()]

                    if "IDCITA" in df_proc.columns:
                        df_proc["IDCITA"] = df_proc["IDCITA"].astype(str).str.strip()
                        df_proc.loc[df_proc["IDCITA"].isin(["", "NAN", "NONE", "NAT", "0"]), "IDCITA"] = pd.NA

                    df_proc["CIE_PROC"] = df_proc["CODIGO"].apply(estandarizar_cie10)
                    df_proc["GRUPO_EDAD"] = df_proc["EDAD_Y"].apply(obtener_grupo_etario)
                    df_proc = df_proc[df_proc["GRUPO_EDAD"].notna()]

                    anio = periodo[:4]
                    mes = periodo[4:]

                    lineas_a = procesar_tabla_a(codigos_validos, ipress_decl, periodo)
                    lineas_b1 = procesar_tabla_b1(df_proc, ipress_decl, periodo)
                    lineas_b2, total_j00x = procesar_tabla_b2(df_proc, ipress_decl, periodo)
                    
                    filtrar_ups = not var_modo_emergencia
                    lineas_c1 = procesar_tabla_c1(df_proc, ipress_decl, periodo, solo_filtrar_ups=filtrar_ups, lista_todos_renipress=codigos_validos)
                    lineas_c2 = procesar_tabla_c2(df_proc, ipress_decl, periodo, solo_filtrar_ups=filtrar_ups)
                    
                    lineas_d1 = procesar_tabla_d1(codigos_validos, ipress_decl, periodo)
                    lineas_d2 = procesar_tabla_d2(codigos_validos, ipress_decl, periodo)
                    lineas_e = procesar_tabla_e(codigos_validos, ipress_decl, periodo, df_principal=df, uploaded_seem=archivo_seem)

                    st.session_state.tramas_resultado = {
                        f"{ipress_decl}_{anio}_{mes}_TAA0.txt": "\n".join(lineas_a),
                        f"{ipress_decl}_{anio}_{mes}_TAB1.txt": "\n".join(lineas_b1),
                        f"{ipress_decl}_{anio}_{mes}_TAB2.txt": "\n".join(lineas_b2),
                        f"{ipress_decl}_{anio}_{mes}_TAC1.txt": "\n".join(lineas_c1),
                        f"{ipress_decl}_{anio}_{mes}_TAC2.txt": "\n".join(lineas_c2),
                        f"{ipress_decl}_{anio}_{mes}_TAD1.txt": "\n".join(lineas_d1),
                        f"{ipress_decl}_{anio}_{mes}_TAD2.txt": "\n".join(lineas_d2),
                        f"{ipress_decl}_{anio}_{mes}_TAE0.txt": "\n".join(lineas_e),
                    }
                    st.success("¡Paquete de Tramas SETI generado con éxito!")
                    st.balloons()

                except Exception as e:
                    st.error(f"No se pudo procesar el paquete: {str(e)}")

    # 3. TABVIEW Y DESCARGA DE ARCHIVOS
    if "tramas_resultado" in st.session_state:
        st.markdown("---")
        st.markdown("### 3. Previsualización y Descarga de Resultados")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for filename, content in st.session_state.tramas_resultado.items():
                zip_file.writestr(filename, content.encode("latin-1", errors="replace"))

        st.download_button(
            label="📦 DESCARGAR PAQUETE COMPLETO (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"SETI_TRAMAS_{periodo}.zip",
            mime="application/zip",
            type="primary"
        )

        tab_a, tab_b1, tab_b2, tab_c1, tab_c2, tab_d1, tab_d2, tab_e = st.tabs(
            ["Tabla A", "Tabla B1", "Tabla B2", "Tabla C1", "Tabla C2", "Tabla D1", "Tabla D2", "Tabla E"]
        )

        dict_tabs = {
            "TAA0": tab_a, "TAB1": tab_b1, "TAB2": tab_b2,
            "TAC1": tab_c1, "TAC2": tab_c2, "TAD1": tab_d1,
            "TAD2": tab_d2, "TAE0": tab_e
        }

        for filename, content in st.session_state.tramas_resultado.items():
            for key_sufijo, tab_obj in dict_tabs.items():
                if key_sufijo in filename:
                    with tab_obj:
                        st.code(content[:2000] if content else "Sin registros.", language="text")


if __name__ == "__main__":
    main()
