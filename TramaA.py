import os
import re
import pandas as pd

# Importación de la librería UI Moderna CustomTkinter
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# Configuración del Tema SETI-PRES
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# =========================================================
# VALORES POR DEFECTO PARA LA TABLA A (RECURSOS)
# =========================================================
RECURSOS_BASE = {
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


# ---------------------------------------------------------
# FUNCIONES BÁSICAS Y MAPEO DE DATOS
# ---------------------------------------------------------
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
    
    # Unificar variantes extendidas (Z298, Z2982, etc.) hacia el código oficial SUSALUD Z29.8
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


def cargar_y_normalizar_excel(ruta_excel):
    try:
        df_raw = pd.read_excel(ruta_excel, header=None)
    except Exception:
        df_raw = pd.read_excel(ruta_excel, engine="xlrd", header=None)

    header_row = 0
    keywords_cabecera = [
        "RENIPRES",
        "NUMDOC",
        "DNI",
        "PACIENTE",
        "CIE10",
        "EESS",
        "TIPODX",
        "ID_UPS",
        "FECPARTO",
    ]

    for idx, row in df_raw.iloc[:30].iterrows():
        valores_fila = [
            str(val).upper().strip() for val in row.values if pd.notna(val)
        ]
        coincidencias = sum(
            1
            for kw in keywords_cabecera
            if any(kw in val for val in valores_fila)
        )
        if coincidencias >= 2:
            header_row = idx
            break

    try:
        df = pd.read_excel(ruta_excel, header=header_row)
    except Exception:
        df = pd.read_excel(ruta_excel, engine="xlrd", header=header_row)

    df.columns = [str(col).strip().upper() for col in df.columns]

    alias_map = {
        "RENIPRESS": [
            "RENIPRESS",
            "RENIPRE",
            "COD_IPRESS",
            "ESTABLECIMIENTO",
            "ID_ESTABLECIMIENTO",
            "RENAES",
        ],
        "NUMDOC": [
            "RHNUMDOC",
            "NUMDOC",
            "NUM_DOC",
            "NRO_DOC",
            "DOCUMENTO",
            "DNI",
            "PACIENTE_DOC",
            "NUMERO_DOCUMENTO",
        ],
        "SEXO": ["SEXO", "SEXO_PACIENTE", "GENERO"],
        "EDAD_Y": ["EDAD_Y", "EDAD", "EDAD_ANOS", "EDAD_A", "EDAD_PACIENTE"],
        "CODIGO": [
            "CODIGO",
            "CIE10",
            "CIE_10",
            "COD_CIE10",
            "DIAGNOSTICO",
            "COD_DIAGNOSTICO",
            "ITEM_CODIGO",
            "CODDIG1",
        ],
        "TIPODX": [
            "TIPODX",
            "TIPO_DX",
            "TIPO_DIAG",
            "TIPO_DIAGNOSTICO",
            "TIP_DX",
            "TIPO",
        ],
        "IDCITA": [
            "IDCITA",
            "ID_CITA",
            "IDCITA_HIS",
            "ID_ATENCION",
            "FECHA_ATENCION",
            "FECHA",
        ],
        "UPS": [
            "ID_UPS",
            "UPS",
            "UPS_ATENCION",
            "COD_UPS",
        ],
        "SERVICIO": [
            "SERVICIO",
            "DESC_UPS",
            "NOMBRE_SERVICIO",
        ],
        "FECPARTO": [
            "FECPARTO",
            "FECHA_PARTO",
        ],
    }

    renombrar = {}
    columnas_existentes = list(df.columns)

    for estandar, aliases in alias_map.items():
        if estandar not in columnas_existentes:
            for alias in aliases:
                coincidencias = [
                    col for col in columnas_existentes if alias in col
                ]
                if coincidencias:
                    renombrar[coincidencias[0]] = estandar
                    break

    if renombrar:
        df = df.rename(columns=renombrar)

    return df


# ---------------------------------------------------------
# LÓGICA DE GENERACIÓN DE TRAMAS (FORMATO SUSALUD REGULARIZADO)
# ---------------------------------------------------------
def procesar_tabla_a(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_a = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    lista_ordenada = sorted(
        codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x
    )

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        recursos = RECURSOS_BASE.get(reni, [0] * 16)
        recursos_str = "|".join(str(val) for val in recursos)

        linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{recursos_str}"
        lineas_a.append(linea)

    return lineas_a


def procesar_tabla_b1(df_base, ipress_matriz, periodo):
    df = df_base.copy()
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if "PROFESION" in df.columns:
        df["ES_MEDICO"] = (
            df["PROFESION"]
            .astype(str)
            .str.upper()
            .str.contains("MEDIC|CMP", regex=True)
        )
    elif "PERSONAL" in df.columns:
        df["ES_MEDICO"] = (
            df["PERSONAL"]
            .astype(str)
            .str.upper()
            .str.contains("MEDIC|CMP|MED", regex=True)
        )
    else:
        df["ES_MEDICO"] = False

    if "IDCITA" in df.columns and df["IDCITA"].notna().any():
        df["ID_ATENCION_UNICA"] = df["IDCITA"].astype(str).str.strip()
    elif "FECHA" in df.columns:
        df["ID_ATENCION_UNICA"] = (
            df["NUMDOC"].astype(str).str.strip()
            + "_"
            + df["FECHA"].astype(str).str.strip()
        )
    else:
        df["ID_ATENCION_UNICA"] = df.index.astype(str)

    lineas_b1 = []
    lista_renipress = sorted(
        df["RENIPRESS_STR"].unique(), key=lambda x: int(x) if x.isdigit() else x
    )

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df[df["RENIPRESS_STR"] == reni]

        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                sub_df = df_reni[
                    (df_reni["SEXO_PROC"] == sexo)
                    & (df_reni["GRUPO_EDAD"] == grupo)
                ]

                if sub_df.empty:
                    continue

                df_med = sub_df[sub_df["ES_MEDICO"]]
                df_no_med = sub_df[~sub_df["ES_MEDICO"]]

                atenciones_med = (
                    df_med["ID_ATENCION_UNICA"].nunique()
                    if not df_med.empty
                    else 0
                )
                atenciones_no_med = (
                    df_no_med["ID_ATENCION_UNICA"].nunique()
                    if not df_no_med.empty
                    else 0
                )
                atendidos_unicos = sub_df["NUMDOC"].astype(str).str.strip().nunique()

                linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{sexo}|{grupo}|{atenciones_med}|{atenciones_no_med}|{atendidos_unicos}"
                lineas_b1.append(linea)

    return lineas_b1


def procesar_tabla_b2(df_base, ipress_matriz, periodo):
    df_b2 = df_base[df_base["CIE_PROC"].apply(es_cie10_morbilidad)].copy()
    
    # Excluir explícitamente los códigos CIE-10 que inician con la letra 'C'
    df_b2 = df_b2[~df_b2["CIE_PROC"].str.startswith("C")]
    
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if df_b2.empty:
        return [], 0

    j00x_g1_h = df_b2[
        (df_b2["CIE_PROC"].str.replace(".", "", regex=False) == "J00X")
        & (df_b2["GRUPO_EDAD"] == 1)
        & (df_b2["SEXO_PROC"] == "1")
    ]
    total_j00x = len(j00x_g1_h)

    lineas_b2 = []
    lista_renipress = sorted(
        df_b2["RENIPRESS_STR"].unique(),
        key=lambda x: int(x) if x.isdigit() else x,
    )

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df_b2[df_b2["RENIPRESS_STR"] == reni]

        agrupado = (
            df_reni.groupby(["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"])
            .size()
            .reset_index(name="TOTAL_ATENDIDOS")
        )
        agrupado = agrupado.sort_values(
            by=["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"]
        )

        for _, fila in agrupado.iterrows():
            lineas_b2.append(
                f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{fila['SEXO_PROC']}|{int(fila['GRUPO_EDAD'])}|{fila['CIE_PROC']}|{fila['TOTAL_ATENDIDOS']}"
            )

    return lineas_b2, total_j00x


def procesar_tabla_c1(df_base, ipress_matriz, periodo, solo_filtrar_ups=True, lista_todos_renipress=None):
    df = df_base.copy()
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    if solo_filtrar_ups and "UPS" in df.columns:
        condicion = df["UPS"].astype(str).str.contains(
            "300303|301601", regex=True
        )
        df_emg = df[condicion].copy()
    else:
        df_emg = df.copy()

    if not df_emg.empty:
        if "IDCITA" in df_emg.columns and df_emg["IDCITA"].notna().any():
            df_emg["ID_VISITA"] = df_emg["IDCITA"].astype(str).str.strip()
        elif "FECHA" in df_emg.columns:
            df_emg["ID_VISITA"] = (
                df_emg["NUMDOC"].astype(str).str.strip()
                + "_"
                + df_emg["FECHA"].astype(str).str.strip()
            )
        else:
            df_emg["ID_VISITA"] = df_emg.index.astype(str)

        df_emg["DNI_PACIENTE"] = (
            df_emg["NUMDOC"].astype(str).str.strip().str.upper()
        )

    lineas_c1 = []

    if lista_todos_renipress:
        lista_renipress = sorted(
            lista_todos_renipress,
            key=lambda x: int(x) if x.isdigit() else x,
        )
    else:
        lista_renipress = sorted(
            df_base["RENIPRESS_STR"].unique(),
            key=lambda x: int(x) if x.isdigit() else x,
        )

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = (
            df_emg[df_emg["RENIPRESS_STR"] == reni]
            if not df_emg.empty and "RENIPRESS_STR" in df_emg.columns
            else pd.DataFrame()
        )

        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                if df_reni.empty:
                    total_atenciones = "NE_0001"
                    total_atendidos = "NE_0001"
                else:
                    sub_df = df_reni[
                        (df_reni["SEXO_PROC"] == sexo)
                        & (df_reni["GRUPO_EDAD"] == grupo)
                    ]

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
        condicion = df["UPS"].astype(str).str.contains(
            " |301601", regex=True
        )
        df_emg = df[condicion].copy()
    else:
        df_emg = df.copy()

    df_emg = df_emg[df_emg["CIE_PROC"].apply(es_cie10_morbilidad)].copy()

    if df_emg.empty:
        return []

    df_emg["DNI_PACIENTE"] = (
        df_emg["NUMDOC"].astype(str).str.strip().str.upper()
    )

    lineas_c2 = []
    lista_renipress = sorted(
        df_emg["RENIPRESS_STR"].unique(),
        key=lambda x: int(x) if x.isdigit() else x,
    )

    for reni in lista_renipress:
        ipress_fmt = reni.zfill(8)
        df_reni = df_emg[df_emg["RENIPRESS_STR"] == reni]

        agrupado = (
            df_reni.groupby(["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"])[
                "DNI_PACIENTE"
            ]
            .nunique()
            .reset_index(name="TOTAL_ATENDIDOS")
        )

        agrupado = agrupado.sort_values(
            by=["SEXO_PROC", "GRUPO_EDAD", "CIE_PROC"]
        )

        for _, fila in agrupado.iterrows():
            linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{fila['SEXO_PROC']}|{int(fila['GRUPO_EDAD'])}|{fila['CIE_PROC']}|{fila['TOTAL_ATENDIDOS']}"
            lineas_c2.append(linea)

    return lineas_c2


def procesar_tabla_d1(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_d1 = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    lista_ordenada = sorted(
        codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x
    )

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001|NE_0001"
        lineas_d1.append(linea)

    return lineas_d1


def procesar_tabla_d2(codigos_renipress_lista, ipress_matriz, periodo):
    lineas_d2 = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    lista_ordenada = sorted(
        codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x
    )

    for reni in lista_ordenada:
        ipress_fmt = str(reni).strip().zfill(8)
        for sexo in ["1", "2"]:
            for grupo in range(1, 16):
                linea = f"{periodo}|{ipress_matriz_fmt}|{ipress_fmt}|{sexo}|{grupo}|NE_0001|NE_0001"
                lineas_d2.append(linea)

    return lineas_d2


def procesar_tabla_e(codigos_renipress_lista, ipress_matriz, periodo, df_principal=None, ruta_seem=None):
    lineas_e = []
    ipress_matriz_fmt = str(ipress_matriz).strip().zfill(8)

    lista_ordenada = sorted(
        codigos_renipress_lista, key=lambda x: int(x) if x.isdigit() else x
    )

    df_fuente = None
    if ruta_seem and os.path.exists(ruta_seem):
        try:
            if ruta_seem.endswith((".xlsx", ".xls")):
                df_fuente = pd.read_excel(ruta_seem)
            else:
                df_fuente = pd.read_csv(ruta_seem, sep=None, engine="python", encoding="latin-1", on_bad_lines="skip")
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


# ---------------------------------------------------------
# VENTANA SECUNDARIA EDITAR TABLA A
# ---------------------------------------------------------
class VentanaFormularioTablaA(ctk.CTkToplevel):

    def __init__(self, parent, lista_renipress):
        super().__init__(parent)
        self.title("Configuración de Recursos - Tabla A")
        self.geometry("580x680")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.lista_renipress = lista_renipress
        self.entries = {}

        self._crear_widgets()

    def _crear_widgets(self):
        lbl_titulo = ctk.CTkLabel(
            self,
            text="RECURSOS EN SALUD (TABLA A)",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1E3A8A",
        )
        lbl_titulo.pack(pady=(15, 5))

        frame_sel = ctk.CTkFrame(self, fg_color="transparent")
        frame_sel.pack(fill="x", padx=20, pady=5)

        ctk.CTkLabel(
            frame_sel, text="IPRESS Seleccionada:", font=ctk.CTkFont(weight="bold")
        ).pack(side="left", padx=(0, 10))

        self.combo_reni = ctk.CTkOptionMenu(
            frame_sel,
            values=self.lista_renipress,
            command=self._cargar_datos_ipress,
            fg_color="#1E3A8A",
            button_color="#1D4ED8",
        )
        self.combo_reni.pack(side="left")

        frame_scroll = ctk.CTkScrollableFrame(self, label_text="Campos de Recursos Físicos y Personal")
        frame_scroll.pack(fill="both", expand=True, padx=20, pady=10)

        for idx, campo in enumerate(CAMPOS_RECURSOS):
            row_frame = ctk.CTkFrame(frame_scroll, fg_color="transparent")
            row_frame.pack(fill="x", pady=4, padx=5)

            lbl = ctk.CTkLabel(row_frame, text=f"{campo}:", anchor="w", width=220)
            lbl.pack(side="left")

            ent = ctk.CTkEntry(row_frame, width=100, justify="center")
            ent.pack(side="right", padx=10)
            self.entries[idx] = ent

        btn_guardar_ipress = ctk.CTkButton(
            self,
            text="Guardar Cambios IPRESS",
            fg_color="#0D9488",
            hover_color="#0F766E",
            font=ctk.CTkFont(weight="bold"),
            command=self._guardar_ipress_actual,
        )
        btn_guardar_ipress.pack(pady=15)

        if self.lista_renipress:
            self.combo_reni.set(self.lista_renipress[0])
            self._cargar_datos_ipress(self.lista_renipress[0])

    def _cargar_datos_ipress(self, choice=None):
        reni = self.combo_reni.get().strip()
        valores = RECURSOS_BASE.get(reni, [0] * 16)

        for idx, ent in self.entries.items():
            ent.delete(0, "end")
            ent.insert(0, str(valores[idx]))

    def _guardar_ipress_actual(self):
        reni = self.combo_reni.get().strip()
        nuevos_valores = []

        try:
            for idx in range(16):
                val = int(self.entries[idx].get().strip())
                nuevos_valores.append(val)

            RECURSOS_BASE[reni] = nuevos_valores
            messagebox.showinfo(
                "Guardado",
                f"Recursos actualizados correctamente para la IPRESS {reni}.",
            )
        except ValueError:
            messagebox.showerror(
                "Error", "Todos los valores deben ser números enteros."
            )


# ---------------------------------------------------------
# INTERFAZ GRÁFICA PRINCIPAL
# ---------------------------------------------------------
class ConsolaSetiSusalud(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Sistema SETI-PRES - Consola Única SUSALUD")
        self.geometry("980x760")
        self.minsize(900, 700)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._crear_sidebar()
        self._crear_panel_principal()

    def _crear_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#1E3A8A")
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        lbl_logo = ctk.CTkLabel(
            self.sidebar,
            text="SETI-PRES",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#FFFFFF",
        )
        lbl_logo.pack(pady=(25, 5))

        lbl_sub = ctk.CTkLabel(
            self.sidebar,
            text="SUSALUD - Módulo de Tramas",
            font=ctk.CTkFont(size=11),
            text_color="#93C5FD",
        )
        lbl_sub.pack(pady=(0, 25))

        btn_config = ctk.CTkButton(
            self.sidebar,
            text=" Configurar Tabla A",
            fg_color="#3B82F6",
            hover_color="#2563EB",
            font=ctk.CTkFont(weight="bold"),
            command=self.abrir_formulario_tabla_a,
        )
        btn_config.pack(padx=15, pady=10, fill="x")

        info_card = ctk.CTkFrame(self.sidebar, fg_color="#1D4ED8", corner_radius=8)
        info_card.pack(padx=15, pady=20, fill="x", side="bottom")

        ctk.CTkLabel(
            info_card,
            text="Elaborado por el Ingº Sistemas\nCarlos Sauñe Yupanqui\nCIP 209420\nEstado: En linea\nVersión 2026.1",
            font=ctk.CTkFont(size=11),
            text_color="#FFFFFF",
        ).pack(pady=10)

    def _crear_panel_principal(self):
        self.main_panel = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=0)
        self.main_panel.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)

        self.main_panel.grid_rowconfigure(0, weight=0)
        self.main_panel.grid_rowconfigure(1, weight=0)
        self.main_panel.grid_rowconfigure(2, weight=0)
        self.main_panel.grid_rowconfigure(3, weight=1)
        self.main_panel.grid_rowconfigure(4, weight=0)
        self.main_panel.grid_columnconfigure(0, weight=1)

        lbl_header = ctk.CTkLabel(
            self.main_panel,
            text="Consola Integrada de Generación de Tramas",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#0F172A",
        )
        lbl_header.grid(row=0, column=0, sticky="w", padx=10, pady=(5, 10))

        # 1. Card Parámetros Generales
        card_param = ctk.CTkFrame(self.main_panel, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#E2E8F0")
        card_param.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        lbl_sec1 = ctk.CTkLabel(
            card_param, text="1. Parámetros del Envío", font=ctk.CTkFont(weight="bold"), text_color="#1E3A8A"
        )
        lbl_sec1.grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(8, 4))

        ctk.CTkLabel(card_param, text="Periodo (AAAAMM):").grid(row=1, column=0, sticky="w", padx=15, pady=3)
        self.txt_periodo = ctk.CTkEntry(card_param, width=150)
        self.txt_periodo.insert(0, "202607")
        self.txt_periodo.grid(row=1, column=1, sticky="w", padx=10, pady=3)

        ctk.CTkLabel(card_param, text="UGIPRESS Matriz:").grid(row=2, column=0, sticky="w", padx=15, pady=3)
        self.txt_ipress_decl = ctk.CTkEntry(card_param, width=150)
        self.txt_ipress_decl.insert(0, "00003554")
        self.txt_ipress_decl.grid(row=2, column=1, sticky="w", padx=10, pady=3)

        ctk.CTkLabel(card_param, text="RENIPRESS a Incluir:").grid(row=3, column=0, sticky="w", padx=15, pady=3)
        self.txt_ipress_lista = ctk.CTkEntry(card_param, width=420)
        self.txt_ipress_lista.insert(
            0, "00003554, 00003555, 00003556, 00003557, 00003558, 00006907"
        )
        self.txt_ipress_lista.grid(row=3, column=1, sticky="w", padx=10, pady=(3, 10))

        # 2. Card Origen de Datos
        card_file = ctk.CTkFrame(self.main_panel, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#E2E8F0")
        card_file.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        lbl_sec2 = ctk.CTkLabel(
            card_file, text="2. Origen de Datos (Excel HIS / Partos)", font=ctk.CTkFont(weight="bold"), text_color="#1E3A8A"
        )
        lbl_sec2.pack(anchor="w", padx=15, pady=(8, 4))

        f_input_file = ctk.CTkFrame(card_file, fg_color="transparent")
        f_input_file.pack(fill="x", padx=15, pady=3)

        self.txt_ruta = ctk.CTkEntry(f_input_file, placeholder_text="Seleccione la ruta del archivo Excel principal...")
        self.txt_ruta.pack(side="left", expand=True, fill="x", padx=(0, 10))

        btn_buscar = ctk.CTkButton(
            f_input_file, text="Buscar Excel", fg_color="#64748B", hover_color="#475569", command=self.buscar_excel
        )
        btn_buscar.pack(side="right")

        f_seem_file = ctk.CTkFrame(card_file, fg_color="transparent")
        f_seem_file.pack(fill="x", padx=15, pady=3)

        self.txt_ruta_seem = ctk.CTkEntry(f_seem_file, placeholder_text="Opcional: Archivo SEEM / Partos independiente...")
        self.txt_ruta_seem.pack(side="left", expand=True, fill="x", padx=(0, 10))

        btn_buscar_seem = ctk.CTkButton(
            f_seem_file, text="Buscar SEEM", fg_color="#64748B", hover_color="#475569", command=self.buscar_seem
        )
        btn_buscar_seem.pack(side="right")

        self.var_deduplicar = ctk.BooleanVar(value=True)
        chk_dedup = ctk.CTkCheckBox(
            card_file,
            text="Descartar duplicados en B2 por Cita/Atención (IDCITA) o Paciente",
            variable=self.var_deduplicar,
            font=ctk.CTkFont(size=12),
        )
        chk_dedup.pack(anchor="w", padx=15, pady=3)

        self.var_modo_emergencia = ctk.BooleanVar(value=False)
        chk_emg = ctk.CTkCheckBox(
            card_file,
            text="Generar C1 y C2 con TODO el Excel (Marcar si el Excel es Exclusivo de Emergencias)",
            variable=self.var_modo_emergencia,
            font=ctk.CTkFont(size=12),
            text_color="#B91C1C",
        )
        chk_emg.pack(anchor="w", padx=15, pady=(3, 10))

        # 3. Tabview Visualizador de Salidas (A a E - Se removió la pestaña F)
        self.tabview = ctk.CTkTabview(self.main_panel)
        self.tabview.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)

        self.tab_a = self.tabview.add("Tabla A")
        self.tab_b1 = self.tabview.add("Tabla B1")
        self.tab_b2 = self.tabview.add("Tabla B2")
        self.tab_c1 = self.tabview.add("Tabla C1")
        self.tab_c2 = self.tabview.add("Tabla C2")
        self.tab_d1 = self.tabview.add("Tabla D1")
        self.tab_d2 = self.tabview.add("Tabla D2")
        self.tab_e = self.tabview.add("Tabla E")

        self.txt_prev_a = self._crear_caja_previsualizacion(self.tab_a)
        self.txt_prev_b1 = self._crear_caja_previsualizacion(self.tab_b1)
        self.txt_prev_b2 = self._crear_caja_previsualizacion(self.tab_b2)
        self.txt_prev_c1 = self._crear_caja_previsualizacion(self.tab_c1)
        self.txt_prev_c2 = self._crear_caja_previsualizacion(self.tab_c2)
        self.txt_prev_d1 = self._crear_caja_previsualizacion(self.tab_d1)
        self.txt_prev_d2 = self._crear_caja_previsualizacion(self.tab_d2)
        self.txt_prev_e = self._crear_caja_previsualizacion(self.tab_e)

        # 4. Botón de Procesamiento Principal
        btn_procesar = ctk.CTkButton(
            self.main_panel,
            text="GENERAR PAQUETE DE TRAMAS SETI (A, B1, B2, C1, C2, D1, D2, E)",
            fg_color="#0D9488",
            hover_color="#0F766E",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=42,
            command=self.procesar_paquete_completo,
        )
        btn_procesar.grid(row=4, column=0, sticky="ew", padx=10, pady=(5, 10))

    def _crear_caja_previsualizacion(self, tab):
        txt = ctk.CTkTextbox(tab, font=ctk.CTkFont(family="Consolas", size=11))
        txt.pack(fill="both", expand=True)
        return txt

    def obtener_lista_renipress_limpia(self):
        lista_raw = self.txt_ipress_lista.get().strip()
        codigos = []
        for item in lista_raw.split(","):
            c = limpiar_codigo_renipress(item)
            if c and c not in codigos:
                codigos.append(c)
        return codigos

    def abrir_formulario_tabla_a(self):
        codigos = self.obtener_lista_renipress_limpia()
        if not codigos:
            messagebox.showwarning(
                "Atención", "Ingrese al menos un código RENIPRESS válido."
            )
            return
        VentanaFormularioTablaA(self, codigos)

    def buscar_excel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel principal",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
        )
        if archivo:
            self.txt_ruta.delete(0, "end")
            self.txt_ruta.insert(0, archivo)

    def buscar_seem(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo SEEM / Partos",
            filetypes=[("Todos los archivos", "*.*"), ("Archivos Excel / CSV", "*.xlsx *.xls *.csv *.txt")],
        )
        if archivo:
            self.txt_ruta_seem.delete(0, "end")
            self.txt_ruta_seem.insert(0, archivo)

    def procesar_paquete_completo(self):
        ruta_excel = self.txt_ruta.get().strip()
        ruta_seem = self.txt_ruta_seem.get().strip()
        periodo = self.txt_periodo.get().strip()
        ipress_decl = self.txt_ipress_decl.get().strip()
        codigos_validos = self.obtener_lista_renipress_limpia()
        set_codigos_validos = set(codigos_validos)

        if not ruta_excel or not os.path.exists(ruta_excel):
            messagebox.showwarning(
                "Archivo requerido", "Seleccione un archivo de Excel válido."
            )
            return

        try:
            df = cargar_y_normalizar_excel(ruta_excel)

            cols_requeridas = [
                "RENIPRESS",
                "NUMDOC",
                "SEXO",
                "EDAD_Y",
                "CODIGO",
                "TIPODX",
            ]
            faltantes = [c for c in cols_requeridas if c not in df.columns]

            if faltantes:
                columnas_detectadas = "\n".join(df.columns[:15])
                messagebox.showerror(
                    "Estructura no coincide",
                    f"Faltan las siguientes columnas requeridas:\n{', '.join(faltantes)}\n\n"
                    f"Columnas detectadas en tu Excel:\n{columnas_detectadas}",
                )
                return

            df_proc = df.copy()

            df_proc["RENIPRESS_STR"] = df_proc["RENIPRESS"].apply(
                limpiar_codigo_renipress
            )
            df_proc = df_proc[df_proc["RENIPRESS_STR"].isin(set_codigos_validos)]

            # FILTRO ESTRICTO: Solo diagnósticos Definitivos (que inician con 'D')
            df_proc = df_proc[
                df_proc["TIPODX"]
                .astype(str)
                .str.strip()
                .str.upper()
                .str.startswith("D")
            ]

            df_proc["SEXO_PROC"] = df_proc["SEXO"].apply(normalizar_sexo)
            df_proc = df_proc[df_proc["SEXO_PROC"].notna()]

            if "IDCITA" in df_proc.columns:
                df_proc["IDCITA"] = df_proc["IDCITA"].astype(str).str.strip()
                df_proc.loc[df_proc["IDCITA"].isin(["", "NAN", "NONE", "NAT", "0"]), "IDCITA"] = pd.NA

            # Estandarizamos y unificamos el código CIE-10 (incluyendo Z298, Z2982 -> Z29.8)
            df_proc["CIE_PROC"] = df_proc["CODIGO"].apply(estandarizar_cie10)

            df_proc["GRUPO_EDAD"] = df_proc["EDAD_Y"].apply(
                obtener_grupo_etario
            )
            df_proc = df_proc[df_proc["GRUPO_EDAD"].notna()]

            carpeta_salida = os.path.join(
                os.path.dirname(ruta_excel), f"SETI_TRAMAS_{periodo}"
            )
            os.makedirs(carpeta_salida, exist_ok=True)

            # ---------------------------------------------------------
            # GENERACIÓN Y GUARDADO DE TRAMAS (CON NOMENCLATURA OFICIAL)
            # ---------------------------------------------------------
            anio = periodo[:4]
            mes = periodo[4:]

            # 1. Tabla A (Sufijo TAA0)
            lineas_a = procesar_tabla_a(codigos_validos, ipress_decl, periodo)
            ruta_a = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAA0.txt"
            )
            with open(ruta_a, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_a))

            # 2. Tabla B1 (Sufijo TAB1)
            lineas_b1 = procesar_tabla_b1(df_proc, ipress_decl, periodo)
            ruta_b1 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAB1.txt"
            )
            with open(ruta_b1, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_b1))

            # 3. Tabla B2 (Sufijo TAB2)
            lineas_b2, total_j00x = procesar_tabla_b2(df_proc, ipress_decl, periodo)
            ruta_b2 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAB2.txt"
            )
            with open(ruta_b2, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_b2))

            # 4. Tabla C1 (Sufijo TAC1)
            filtrar_ups = not self.var_modo_emergencia.get()
            lineas_c1 = procesar_tabla_c1(
                df_proc,
                ipress_decl,
                periodo,
                solo_filtrar_ups=filtrar_ups,
                lista_todos_renipress=codigos_validos,
            )
            ruta_c1 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAC1.txt"
            )
            with open(ruta_c1, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_c1))

            # 5. Tabla C2 (Sufijo TAC2)
            lineas_c2 = procesar_tabla_c2(
                df_proc,
                ipress_decl,
                periodo,
                solo_filtrar_ups=filtrar_ups,
            )
            ruta_c2 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAC2.txt"
            )
            with open(ruta_c2, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_c2))

            # 6. Tabla D1 (Sufijo TAD1)
            lineas_d1 = procesar_tabla_d1(codigos_validos, ipress_decl, periodo)
            ruta_d1 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAD1.txt"
            )
            with open(ruta_d1, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_d1))

            # 7. Tabla D2 (Sufijo TAD2)
            lineas_d2 = procesar_tabla_d2(codigos_validos, ipress_decl, periodo)
            ruta_d2 = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAD2.txt"
            )
            with open(ruta_d2, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_d2))

            # 8. Tabla E (Sufijo TAE0)
            lineas_e = procesar_tabla_e(codigos_validos, ipress_decl, periodo, df_principal=df, ruta_seem=ruta_seem)
            ruta_e = os.path.join(
                carpeta_salida, f"{ipress_decl}_{anio}_{mes}_TAE0.txt"
            )
            with open(ruta_e, "w", encoding="ansi", errors="replace") as f:
                f.write("\n".join(lineas_e))

            self._actualizar_previsualizaciones(
                lineas_a, lineas_b1, lineas_b2, lineas_c1, lineas_c2, lineas_d1, lineas_d2, lineas_e
            )

            messagebox.showinfo(
                "¡Generación Exitosa!",
                f"Archivos generados en formato ANSI en:\n{carpeta_salida}\n\n"
                f"• Tabla A: {len(lineas_a)} registros\n"
                f"• Tabla B1: {len(lineas_b1)} líneas\n"
                f"• Tabla B2: {len(lineas_b2)} líneas\n"
                f"• Tabla C1: {len(lineas_c1)} líneas\n"
                f"• Tabla C2: {len(lineas_c2)} líneas\n"
                f"• Tabla D1: {len(lineas_d1)} líneas\n"
                f"• Tabla D2: {len(lineas_d2)} líneas\n"
                f"• Tabla E: {len(lineas_e)} líneas",
            )

        except Exception as e:
            messagebox.showerror(
                "Error técnico", f"No se pudo procesar el paquete:\n{str(e)}"
            )

    def _actualizar_previsualizaciones(self, a, b1, b2, c1, c2, d1, d2, e):
        for widget, lista, in zip(
            [
                self.txt_prev_a,
                self.txt_prev_b1,
                self.txt_prev_b2,
                self.txt_prev_c1,
                self.txt_prev_c2,
                self.txt_prev_d1,
                self.txt_prev_d2,
                self.txt_prev_e,
            ],
            [a, b1, b2, c1, c2, d1, d2, e],
        ):
            widget.delete("1.0", "end")
            widget.insert("1.0", "\n".join(lista[:100]))


if __name__ == "__main__":
    app = ConsolaSetiSusalud() 
    app.mainloop()