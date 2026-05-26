try:
    import plotly.express as px
    print("PLOTLY OK")
except Exception as e:
    print("ERROR PLOTLY:", e)
import streamlit as st
import pandas as pd
import numpy as np
import plotly
import plotly.express as px
import plotly.graph_objects as go
import warnings
import os

warnings.filterwarnings("ignore")

# ── Colores UCLA ──────────────────────────────────────────────────────────────
BLUE    = "#007B99"
ORANGE  = "#F39200"
GREY    = "#848585"
GREEN   = "#28a745"
RED     = "#dc3545"
PALETTE = [BLUE, ORANGE, GREY, "#A9D6E5", "#F9C784", "#5BA4CF", "#E8A838"]

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard IT — Funlam",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
  .block-container { padding-top: 1rem; }
  .header-bar {
    background: linear-gradient(135deg, #007B99, #005f78);
    padding: 1rem 1.5rem; border-radius: 10px; margin-bottom: 1.2rem;
  }
  .header-bar h1 { color: white; margin: 0; font-size: 1.5rem; font-weight: 700; }
  .header-bar p  { color: rgba(255,255,255,0.8); margin: 0; font-size: 0.85rem; }
  .kpi { background:white; border-radius:10px; padding:1rem 1.2rem;
         border-left:5px solid #007B99; box-shadow:0 2px 8px rgba(0,0,0,0.08); }
  .kpi.or { border-left-color:#F39200; }
  .kpi.gr { border-left-color:#28a745; }
  .kpi.gy { border-left-color:#848585; }
  .kpi.rd { border-left-color:#dc3545; }
  .kpi-label { font-size:0.72rem; color:#848585; text-transform:uppercase;
               letter-spacing:.05em; font-weight:600; }
  .kpi-value { font-size:1.9rem; font-weight:800; color:#1a1a2e; line-height:1.1; }
  .kpi-sub   { font-size:0.73rem; color:#848585; margin-top:2px; }
  .stTabs [data-baseweb="tab"] { font-weight:600; font-size:0.9rem; }
  .sidebar-section { background:#f0f4f8; border-radius:8px; padding:0.7rem 0.8rem; margin-bottom:0.8rem; }
  .sidebar-section h4 { color:#007B99; margin:0 0 0.4rem 0; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
  <h1>🖥️ Dashboard de Infraestructura Tecnológica</h1>
  <p>Universidad Católica Luis Amigó · Departamento de Infraestructura TI · ISE009 Big Data</p>
  <p><small>Sugerencia: El modo claro permite una mayor armonía visual</small></p>
</div>
""", unsafe_allow_html=True)

# ── KPI helper ────────────────────────────────────────────────────────────────
def kpi(label, value, sub="", variant=""):
    st.markdown(f"""
    <div class="kpi {variant}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

# ── Carga de datos ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def cargar_datos_procesados(uploaded_tickets, uploaded_inventario, uploaded_software, uploaded_actividades):
    def detectar_header(xl, hoja, max_filas=6):
        df_raw = xl.parse(hoja, header=None, nrows=max_filas)
        mejor_fila, mejor_score = 0, -1
        for i, row in df_raw.iterrows():
            score = sum(2 if isinstance(v, str) and v.strip() else
                        1 if pd.notna(v) else 0 for v in row)
            if score > mejor_score:
                mejor_score, mejor_fila = score, i
        return int(mejor_fila)

    def leer_hoja(xl, hoja, h=None):
        h = h if h is not None else detectar_header(xl, hoja)
        df = xl.parse(hoja, header=h).dropna(how="all").dropna(axis=1, how="all")
        cols = [str(c).strip() if not str(c).startswith("Unnamed") else f"_col_{i}"
                for i, c in enumerate(df.columns)]
        df.columns = cols
        df = df[[c for c in cols if not c.startswith("_col_")]]
        return df.reset_index(drop=True)

    def fix_mojibake(text):
        if isinstance(text, str):
            try:
                return text.encode("latin1").decode("utf-8")
            except (UnicodeDecodeError, UnicodeEncodeError):
                return text
        return text

    con = det = inv = eventos = ups = sw = act = pd.DataFrame()

    if uploaded_tickets:
        xl_t = pd.ExcelFile(uploaded_tickets)
        det  = leer_hoja(xl_t, "DETALLADO", h=0)
        for col in det.select_dtypes("object").columns:
            det[col] = det[col].apply(fix_mojibake)
        det["Fecha de Creación"] = pd.to_datetime(
            det.get("Fecha de Creación", det.get("Fecha de CreaciÃ³n", pd.NaT)), errors="coerce")
        det["mes"]        = det["Fecha de Creación"].dt.month
        det["año"]        = det["Fecha de Creación"].dt.year
        det["mes_label"]  = det["Fecha de Creación"].dt.strftime("%b %Y")
        det["mes_orden"]  = det["Fecha de Creación"].dt.to_period("M").astype(str)
        det["dia_semana"] = det["Fecha de Creación"].dt.day_name()
        det["semana"]     = det["Fecha de Creación"].dt.isocalendar().week.astype("Int64")
        con = leer_hoja(xl_t, "CONSOLIDADO")

    if uploaded_inventario:
        xl_i = pd.ExcelFile(uploaded_inventario)

        # Headers fijos por hoja (detectados del diagnóstico real del archivo)
        # Todas las hojas de inventario tienen header en fila 1 (fila 0 = título/logo)
        HEADERS_INV = {
            "Servidores":           1,
            "Switches de datos":    1,
            "Access Point":         1,
            "Seguridad Perimetral": 1,
            "UPS":                  1,
            "Registro y Gestión":   1,
            "Canales de internet":  1,
        }

        def inv_hoja(hoja, tipo):
            h = HEADERS_INV.get(hoja, 1)
            d = xl_i.parse(hoja, header=h)
            d = d.dropna(how="all")           # eliminar filas vacías
            # NO eliminar columnas vacías aquí — se hace después del concat
            # Limpiar nombres de columna
            cols = []
            for i, c in enumerate(d.columns):
                s = str(c).strip()
                cols.append(s if not s.startswith("Unnamed") and s != "nan" else f"_col_{i}")
            d.columns = cols
            d = d[[c for c in cols if not c.startswith("_col_")]]
            d["TIPO_EQUIPO"] = tipo
            return d.reset_index(drop=True)

        # Concat con sort=False para preservar todas las columnas de todas las hojas
        piezas = []
        for hoja, tipo in [
            ("Servidores",           "Servidor"),
            ("Switches de datos",    "Switch"),
            ("Access Point",         "Access Point"),
            ("Seguridad Perimetral", "Seguridad Perimetral"),
            ("UPS",                  "UPS"),
        ]:
            try:
                piezas.append(inv_hoja(hoja, tipo))
            except Exception:
                pass

        inv = pd.concat(piezas, ignore_index=True, sort=False)

        # UPS separado para la tabla de baterías
        ups = inv_hoja("UPS", "UPS")
        col_bat = "FECHA ÚLTIMO CAMBIO DE BATERÍAS"
        if col_bat in ups.columns:
            ups[col_bat] = pd.to_datetime(ups[col_bat], errors="coerce")
            ups["dias_desde_cambio"] = (pd.Timestamp.now() - ups[col_bat]).dt.days

        # Eventos
        try:
            eventos = xl_i.parse("Registro y Gestión", header=HEADERS_INV["Registro y Gestión"])
            eventos = eventos.dropna(how="all")
        except Exception:
            eventos = pd.DataFrame()

    if uploaded_software:
        xl_s = pd.ExcelFile(uploaded_software)
        # Ambas hojas de software tienen header en fila 0 (confirmado en diagnóstico)
        def prep_sw(hoja, cat):
            d = xl_s.parse(hoja, header=0)
            d = d.dropna(how="all")
            cols = []
            for i, c in enumerate(d.columns):
                s = str(c).strip()
                cols.append(s if not s.startswith("Unnamed") and s != "nan" else f"_col_{i}")
            d.columns = cols
            d = d[[c for c in cols if not c.startswith("_col_")]]
            d["CATEGORIA"] = cat
            if "FECHA DE VENCIMIENTO" in d.columns:
                d["FECHA DE VENCIMIENTO"] = pd.to_datetime(d["FECHA DE VENCIMIENTO"], errors="coerce")
                d["dias_para_vencer"] = (d["FECHA DE VENCIMIENTO"] - pd.Timestamp.now()).dt.days
                d["estado_licencia"] = d["dias_para_vencer"].apply(
                    lambda x: "Sin fecha" if pd.isna(x) else
                              "Vencida" if x < 0 else
                              "Por vencer (<60d)" if x <= 60 else "Vigente")
            return d.reset_index(drop=True)
        sw = pd.concat([prep_sw("SOTWARE EDUCACION",  "Educativo"),
                        prep_sw("SOTWARE OPERACIONAL", "Operacional")],
                       ignore_index=True, sort=False)

    if uploaded_actividades:
        xl_a = pd.ExcelFile(uploaded_actividades)
        # Actividades tiene header en fila 0 (confirmado en diagnóstico: 223 filas × 8 col)
        act = xl_a.parse(xl_a.sheet_names[0], header=0)
        act = act.dropna(how="all")
        cols = []
        for i, c in enumerate(act.columns):
            s = str(c).strip()
            cols.append(s if not s.startswith("Unnamed") and s != "nan" else f"_col_{i}")
        act.columns = cols
        act = act[[c for c in cols if not c.startswith("_col_")]].reset_index(drop=True)
        for col in ["Fecha de creación", "Fecha estimada de cierre", "Fecha real de cierre"]:
            if col in act.columns:
                act[col] = pd.to_datetime(act[col], errors="coerce")
        if {"Fecha real de cierre", "Fecha estimada de cierre"} <= set(act.columns):
            act["dias_retraso"] = (act["Fecha real de cierre"] - act["Fecha estimada de cierre"]).dt.days
        if "Asignado a" in act.columns:
            act["Desarrollador"] = act["Asignado a"].str.split("<").str[0].str.strip()

    return con, det, inv, eventos, ups, sw, act

# ════════════════════════════════════════════════════════════════════
# SIDEBAR — Carga de archivos
# ════════════════════════════════════════════════════════════════════
st.sidebar.header("📂 Cargar Archivos Excel")
uploaded_tickets     = st.sidebar.file_uploader("🎫 Tickets (tickets.xlsx)",           type=["xlsx"])
uploaded_inventario  = st.sidebar.file_uploader("🖥️ Inventario Infraestructura (.xlsx)", type=["xlsx"])
uploaded_software    = st.sidebar.file_uploader("💿 Software Institucional (.xlsx)",    type=["xlsx"])
uploaded_actividades = st.sidebar.file_uploader("🛠️ Reporte de Actividades (.xlsx)",    type=["xlsx"])

con, det, inv, eventos, ups, sw, act = cargar_datos_procesados(
    uploaded_tickets, uploaded_inventario, uploaded_software, uploaded_actividades)

st.sidebar.markdown("---")

# ════════════════════════════════════════════════════════════════════
# TABS PRINCIPALES
# ════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "🎫 Tickets de Soporte",
    "🖥️ Inventario de Infraestructura",
    "💿 Software & Licencias",
    "🛠️ Actividades de Desarrollo",
])

# ═══════════════════════════════════════════════════════════
# TAB 1 — TICKETS
# ═══════════════════════════════════════════════════════════
with tab1:
    if not uploaded_tickets or det.empty:
        st.info("👈 Sube el archivo de Tickets para ver esta sección.")
    else:
        # ── Filtros propios de esta pestaña ──────────────────────
        with st.sidebar:
            st.markdown('<div class="sidebar-section"><h4>🎫 Filtros — Tickets</h4>', unsafe_allow_html=True)

            # Filtro de rango de fechas
            fecha_min = det["Fecha de Creación"].min()
            fecha_max = det["Fecha de Creación"].max()
            if pd.notna(fecha_min) and pd.notna(fecha_max):
                rango_fecha = st.date_input(
                    "Rango de fechas:",
                    value=(fecha_min.date(), fecha_max.date()),
                    min_value=fecha_min.date(),
                    max_value=fecha_max.date(),
                    key="tick_fecha"
                )
            else:
                rango_fecha = None

            tipos_disp   = sorted(det["tipo"].dropna().unique())   if "tipo"    in det.columns else []
            estados_disp = sorted(det["Estado"].dropna().unique()) if "Estado"  in det.columns else []
            topicos_disp = sorted(det["Tópico"].dropna().unique()) if "Tópico"  in det.columns else []
            empleados_disp = sorted(det["Empleado"].dropna().unique()) if "Empleado" in det.columns else []

            tipos_sel    = st.multiselect("Tipo de ticket:",  tipos_disp,    default=tipos_disp,    key="tick_tipo")
            estados_sel  = st.multiselect("Estado:",          estados_disp,  default=estados_disp,  key="tick_estado")
            topicos_sel  = st.multiselect("Tópico:",          topicos_disp,  default=topicos_disp,  key="tick_topico")
            empleados_sel = st.multiselect("Empleado asignado:", empleados_disp, default=empleados_disp, key="tick_emp")
            st.markdown('</div>', unsafe_allow_html=True)

        # Aplicar filtros
        df_t = det.copy()
        if rango_fecha and len(rango_fecha) == 2:
            df_t = df_t[
                (df_t["Fecha de Creación"].dt.date >= rango_fecha[0]) &
                (df_t["Fecha de Creación"].dt.date <= rango_fecha[1])
            ]
        if tipos_sel    and "tipo"     in df_t.columns: df_t = df_t[df_t["tipo"].isin(tipos_sel)]
        if estados_sel  and "Estado"   in df_t.columns: df_t = df_t[df_t["Estado"].isin(estados_sel)]
        if topicos_sel  and "Tópico"   in df_t.columns: df_t = df_t[df_t["Tópico"].isin(topicos_sel)]
        if empleados_sel and "Empleado" in df_t.columns: df_t = df_t[df_t["Empleado"].isin(empleados_sel)]

        st.caption(f"Mostrando **{len(df_t):,}** tickets de {len(det):,} totales con los filtros aplicados.")

        # ── KPIs ─────────────────────────────────────────────────
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Total Tickets", f"{len(df_t):,}", f"de {len(det):,} totales")
        abiertos = len(df_t[df_t["Estado"] == "Abierto"]) if "Estado" in df_t.columns else 0
        with c2: kpi("Abiertos", f"{abiertos:,}", "pendientes de resolución", "or")
        cerrados = len(df_t[df_t["Estado"] == "Cerrar"]) if "Estado" in df_t.columns else 0
        with c3: kpi("Cerrados", f"{cerrados:,}", f"{cerrados/max(len(df_t),1)*100:.1f}% del total", "gr")
        en_curso = len(df_t[df_t["Estado"] == "En curso"]) if "Estado" in df_t.columns else 0
        with c4: kpi("En Curso", f"{en_curso:,}", "en proceso", "gy")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Fila 1: Tipo + Estado ─────────────────────────────────
        g1, g2 = st.columns(2)
        with g1:
            if "tipo" in df_t.columns:
                conteo = df_t["tipo"].value_counts().reset_index()
                conteo.columns = ["Tipo", "Cantidad"]
                fig = px.bar(conteo, x="Cantidad", y="Tipo", orientation="h",
                             color_discrete_sequence=[BLUE], title="Tickets por Tipo", text="Cantidad")
                fig.update_traces(textposition="outside")
                fig.update_layout(showlegend=False, plot_bgcolor="white", margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig, use_container_width=True)

        with g2:
            if "Estado" in df_t.columns:
                conteo_e = df_t["Estado"].value_counts().reset_index()
                conteo_e.columns = ["Estado", "Cantidad"]
                fig2 = px.pie(conteo_e, names="Estado", values="Cantidad",
                              color_discrete_sequence=PALETTE, title="Distribución por Estado", hole=0.45)
                fig2.update_layout(margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig2, use_container_width=True)

        # ── Tendencia mensual ─────────────────────────────────────
        if "mes_orden" in df_t.columns:
            tend = (df_t.groupby(["mes_orden", "mes_label"]).size()
                    .reset_index(name="tickets").sort_values("mes_orden"))
            fig3 = px.line(tend, x="mes_label", y="tickets", markers=True,
                           color_discrete_sequence=[BLUE], title="Volumen Mensual de Tickets")
            fig3.update_layout(plot_bgcolor="white", xaxis_title="Período",
                               yaxis_title="Tickets", margin=dict(t=40,b=10))
            fig3.update_traces(fill="tozeroy", fillcolor="rgba(0,123,153,0.08)")
            st.plotly_chart(fig3, use_container_width=True)

        # ── Fila 2: Día semana + Tópico ───────────────────────────
        g3, g4 = st.columns(2)
        with g3:
            if "dia_semana" in df_t.columns:
                orden_dias = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
                labels_es  = {"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                              "Thursday":"Jueves","Friday":"Viernes","Saturday":"Sábado","Sunday":"Domingo"}
                by_day = df_t["dia_semana"].value_counts().reindex(orden_dias).reset_index()
                by_day.columns = ["dia", "tickets"]
                by_day["dia"] = by_day["dia"].map(labels_es)
                fig4 = px.bar(by_day, x="dia", y="tickets", color_discrete_sequence=[ORANGE],
                              title="Tickets por Día de la Semana", text="tickets")
                fig4.update_traces(textposition="outside")
                fig4.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=320)
                st.plotly_chart(fig4, use_container_width=True)

        with g4:
            if "Tópico" in df_t.columns:
                top_top = df_t["Tópico"].value_counts().reset_index()
                top_top.columns = ["Tópico", "Cantidad"]
                fig5 = px.bar(top_top, x="Cantidad", y="Tópico", orientation="h",
                              color_discrete_sequence=[GREY], title="Tickets por Tópico", text="Cantidad")
                fig5.update_traces(textposition="outside")
                fig5.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=320)
                st.plotly_chart(fig5, use_container_width=True)

        # ── Empleados ─────────────────────────────────────────────
        if "Empleado" in df_t.columns:
            by_emp = df_t["Empleado"].value_counts().head(10).reset_index()
            by_emp.columns = ["Empleado", "Tickets"]
            fig6 = px.bar(by_emp, x="Tickets", y="Empleado", orientation="h",
                          color_discrete_sequence=[GREY],
                          title="Top 10 Empleados con más Tickets asignados", text="Tickets")
            fig6.update_traces(textposition="outside")
            fig6.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=360)
            st.plotly_chart(fig6, use_container_width=True)

        # ── NUEVO: Heatmap tickets por día×mes ────────────────────
        if "dia_semana" in df_t.columns and "mes_label" in df_t.columns:
            st.markdown("#### 🗓️ Mapa de calor — Actividad por Día y Mes")
            orden_dias_es = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
            labels_inv = {v: k for k, v in {"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                          "Thursday":"Jueves","Friday":"Viernes","Saturday":"Sábado","Sunday":"Domingo"}.items()}
            heat = df_t.copy()
            heat["dia_es"] = heat["dia_semana"].map(
                {"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                 "Thursday":"Jueves","Friday":"Viernes","Saturday":"Sábado","Sunday":"Domingo"})
            heat_pivot = (heat.groupby(["mes_orden","mes_label","dia_es"]).size()
                          .reset_index(name="tickets"))
            meses_orden = heat_pivot.sort_values("mes_orden")["mes_label"].unique()
            pivot = heat_pivot.pivot_table(index="dia_es", columns="mes_label",
                                           values="tickets", aggfunc="sum", fill_value=0)
            pivot = pivot.reindex([d for d in orden_dias_es if d in pivot.index])
            pivot = pivot[[m for m in meses_orden if m in pivot.columns]]
            fig_heat = px.imshow(pivot, color_continuous_scale=[[0,"#e8f4f8"],[1,BLUE]],
                                 title="Intensidad de Tickets por Día y Mes",
                                 labels=dict(color="Tickets"))
            fig_heat.update_layout(margin=dict(t=50,b=10), height=320)
            st.plotly_chart(fig_heat, use_container_width=True)

        # ── NUEVO: Acumulado de tickets en el tiempo ──────────────
        if "mes_orden" in df_t.columns:
            st.markdown("#### 📈 Tickets Acumulados en el Tiempo")
            acum = (df_t.groupby(["mes_orden","mes_label"]).size()
                    .reset_index(name="tickets").sort_values("mes_orden"))
            acum["acumulado"] = acum["tickets"].cumsum()
            fig_acum = go.Figure()
            fig_acum.add_trace(go.Bar(x=acum["mes_label"], y=acum["tickets"],
                                      name="Mensual", marker_color=BLUE, opacity=0.6))
            fig_acum.add_trace(go.Scatter(x=acum["mes_label"], y=acum["acumulado"],
                                          name="Acumulado", mode="lines+markers",
                                          line=dict(color=ORANGE, width=3)))
            fig_acum.update_layout(plot_bgcolor="white", barmode="overlay",
                                   xaxis_title="Mes", yaxis_title="Tickets",
                                   legend=dict(orientation="h", yanchor="bottom", y=1.02),
                                   margin=dict(t=30,b=10), height=340)
            st.plotly_chart(fig_acum, use_container_width=True)

# ═══════════════════════════════════════════════════════════
# TAB 2 — INVENTARIO
# ═══════════════════════════════════════════════════════════
with tab2:
    if not uploaded_inventario or inv.empty:
        st.info("👈 Sube el archivo de Inventario para ver esta sección.")
    else:
        # ── Filtros propios de esta pestaña ──────────────────────
        with st.sidebar:
            st.markdown('<div class="sidebar-section"><h4>🖥️ Filtros — Inventario</h4>', unsafe_allow_html=True)
            tipos_eq   = sorted(inv["TIPO_EQUIPO"].dropna().unique()) if "TIPO_EQUIPO" in inv.columns else []
            sedes_eq   = sorted(inv["CENTRO REGIONAL"].dropna().unique()) if "CENTRO REGIONAL" in inv.columns else []
            tipos_eq_sel = st.multiselect("Tipo de equipo:", tipos_eq, default=tipos_eq, key="inv_tipo")
            sedes_sel    = st.multiselect("Sede / Centro Regional:", sedes_eq, default=sedes_eq, key="inv_sede")
            if "TIPO DE INSTALACIÓN" in inv.columns:
                inst_disp = sorted(inv["TIPO DE INSTALACIÓN"].dropna().unique())
                inst_sel  = st.multiselect("Instalación:", inst_disp, default=inst_disp, key="inv_inst")
            else:
                inst_sel = []
            st.markdown('</div>', unsafe_allow_html=True)

        df_inv = inv.copy()
        if tipos_eq_sel and "TIPO_EQUIPO"       in df_inv.columns: df_inv = df_inv[df_inv["TIPO_EQUIPO"].isin(tipos_eq_sel)]
        if sedes_sel    and "CENTRO REGIONAL"   in df_inv.columns: df_inv = df_inv[df_inv["CENTRO REGIONAL"].isin(sedes_sel)]
        if inst_sel     and "TIPO DE INSTALACIÓN" in df_inv.columns: df_inv = df_inv[df_inv["TIPO DE INSTALACIÓN"].isin(inst_sel)]

        st.caption(f"Mostrando **{len(df_inv):,}** equipos de {len(inv):,} totales.")

        # ── KPIs ─────────────────────────────────────────────────
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Equipos Totales", f"{len(df_inv):,}", "en inventario filtrado")
        serv_n = len(df_inv[df_inv["TIPO_EQUIPO"] == "Servidor"]) if "TIPO_EQUIPO" in df_inv.columns else 0
        with c2: kpi("Servidores", f"{serv_n}", "físicos y virtuales", "or")
        ap_n = len(df_inv[df_inv["TIPO_EQUIPO"] == "Access Point"]) if "TIPO_EQUIPO" in df_inv.columns else 0
        with c3: kpi("Access Points", f"{ap_n}", "puntos de acceso WiFi", "gr")
        ups_n = len(df_inv[df_inv["TIPO_EQUIPO"] == "UPS"]) if "TIPO_EQUIPO" in df_inv.columns else 0
        with c4: kpi("UPS", f"{ups_n}", "respaldo eléctrico", "gy")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Fila 1: Tipo + Sede ───────────────────────────────────
        g1, g2 = st.columns(2)
        with g1:
            if "TIPO_EQUIPO" in df_inv.columns:
                ct = df_inv["TIPO_EQUIPO"].value_counts().reset_index()
                ct.columns = ["Tipo", "Cantidad"]
                fig = px.bar(ct, x="Tipo", y="Cantidad", color_discrete_sequence=[BLUE],
                             title="Equipos por Tipo", text="Cantidad")
                fig.update_traces(textposition="outside")
                fig.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=320)
                st.plotly_chart(fig, use_container_width=True)

        with g2:
            if "CENTRO REGIONAL" in df_inv.columns:
                cs = df_inv["CENTRO REGIONAL"].value_counts().reset_index()
                cs.columns = ["Sede", "Cantidad"]
                fig2 = px.pie(cs, names="Sede", values="Cantidad",
                              color_discrete_sequence=PALETTE, title="Equipos por Sede", hole=0.4)
                fig2.update_layout(margin=dict(t=40,b=10), height=320)
                st.plotly_chart(fig2, use_container_width=True)

        # ── Fila 2: Nube vs On-Premise + Misión Crítica ───────────
        if "TIPO DE INSTALACIÓN" in df_inv.columns:
            g3, g4 = st.columns(2)
            with g3:
                inst = df_inv["TIPO DE INSTALACIÓN"].value_counts().reset_index()
                inst.columns = ["Instalación", "Cantidad"]
                fig3 = px.pie(inst, names="Instalación", values="Cantidad",
                              color_discrete_sequence=[BLUE, ORANGE],
                              title="On Premise vs Nube", hole=0.4)
                fig3.update_layout(margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig3, use_container_width=True)

            with g4:
                if "MISIÓN CRÍTICA" in df_inv.columns:
                    mc = df_inv["MISIÓN CRÍTICA"].value_counts().reset_index()
                    mc.columns = ["Crítico", "Cantidad"]
                    fig4 = px.bar(mc, x="Crítico", y="Cantidad",
                                  color_discrete_sequence=[ORANGE, GREY],
                                  title="Equipos de Misión Crítica", text="Cantidad")
                    fig4.update_traces(textposition="outside")
                    fig4.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=300)
                    st.plotly_chart(fig4, use_container_width=True)

        # ── NUEVO: Marca por tipo de equipo ───────────────────────
        if "MARCA" in df_inv.columns and "TIPO_EQUIPO" in df_inv.columns:
            st.markdown("#### 🏷️ Marcas por Tipo de Equipo")
            marca_tipo = (df_inv.dropna(subset=["MARCA"])
                          .groupby(["TIPO_EQUIPO","MARCA"]).size()
                          .reset_index(name="Cantidad"))
            fig_marca = px.bar(marca_tipo, x="TIPO_EQUIPO", y="Cantidad", color="MARCA",
                               color_discrete_sequence=PALETTE,
                               title="Distribución de Marcas por Tipo de Equipo",
                               barmode="stack", text_auto=True)
            fig_marca.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=360)
            st.plotly_chart(fig_marca, use_container_width=True)

        # ── NUEVO: Sede × Tipo (heatmap) ──────────────────────────
        if "CENTRO REGIONAL" in df_inv.columns and "TIPO_EQUIPO" in df_inv.columns:
            st.markdown("#### 🗺️ Distribución de Equipos por Sede y Tipo")
            pivot_inv = df_inv.groupby(["CENTRO REGIONAL","TIPO_EQUIPO"]).size().unstack(fill_value=0)
            fig_hinv = px.imshow(pivot_inv, color_continuous_scale=[[0,"#e8f4f8"],[1,BLUE]],
                                 title="Cantidad de Equipos — Sede × Tipo",
                                 labels=dict(color="Equipos"))
            fig_hinv.update_layout(margin=dict(t=50,b=10), height=280)
            st.plotly_chart(fig_hinv, use_container_width=True)

        # ── UPS baterías + Eventos ────────────────────────────────
        if "dias_desde_cambio" in ups.columns:
            st.markdown("#### 🔋 Estado de Baterías UPS")
            ups_m = ups[[c for c in ["UBICACIÓN","MARCA","MODELO","CAPACIDAD",
                                      "FECHA ÚLTIMO CAMBIO DE BATERÍAS","dias_desde_cambio"]
                          if c in ups.columns]].copy().sort_values("dias_desde_cambio", ascending=False)
            st.dataframe(ups_m, use_container_width=True, height=260)

        if not eventos.empty:
            st.markdown("#### ⚠️ Registro de Eventos e Incidentes")
            st.dataframe(eventos, use_container_width=True, height=200)

# ═══════════════════════════════════════════════════════════
# TAB 3 — SOFTWARE
# ═══════════════════════════════════════════════════════════
with tab3:
    if not uploaded_software or sw.empty:
        st.info("👈 Sube el archivo de Software para ver esta sección.")
    else:
        # ── Filtros propios de esta pestaña ──────────────────────
        with st.sidebar:
            st.markdown('<div class="sidebar-section"><h4>💿 Filtros — Software</h4>', unsafe_allow_html=True)
            cat_disp  = sorted(sw["CATEGORIA"].dropna().unique())        if "CATEGORIA"       in sw.columns else []
            tipo_sw   = sorted(sw["SOFTWARE"].dropna().unique())         if "SOFTWARE"        in sw.columns else []
            est_lic   = sorted(sw["estado_licencia"].dropna().unique())  if "estado_licencia" in sw.columns else []
            prov_disp = sorted(sw["PROVEEDOR"].dropna().unique())        if "PROVEEDOR"       in sw.columns else []

            cat_sel  = st.multiselect("Categoría:",         cat_disp,  default=cat_disp,  key="sw_cat")
            tipo_sel = st.multiselect("Tipo licencia:",     tipo_sw,   default=tipo_sw,   key="sw_tipo")
            est_sel  = st.multiselect("Estado licencia:",   est_lic,   default=est_lic,   key="sw_est")
            prov_sel = st.multiselect("Proveedor:",         prov_disp, default=prov_disp, key="sw_prov")
            st.markdown('</div>', unsafe_allow_html=True)

        df_sw = sw.copy()
        if cat_sel  and "CATEGORIA"       in df_sw.columns: df_sw = df_sw[df_sw["CATEGORIA"].isin(cat_sel)]
        if tipo_sel and "SOFTWARE"        in df_sw.columns: df_sw = df_sw[df_sw["SOFTWARE"].isin(tipo_sel)]
        if est_sel  and "estado_licencia" in df_sw.columns: df_sw = df_sw[df_sw["estado_licencia"].isin(est_sel)]
        if prov_sel and "PROVEEDOR"       in df_sw.columns: df_sw = df_sw[df_sw["PROVEEDOR"].isin(prov_sel)]

        st.caption(f"Mostrando **{len(df_sw):,}** licencias de {len(sw):,} totales.")

        # ── KPIs ─────────────────────────────────────────────────
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Licencias Totales", f"{len(df_sw):,}", "edu + operacional")
        lic_n = len(df_sw[df_sw["SOFTWARE"] == "Licenciado"]) if "SOFTWARE" in df_sw.columns else len(df_sw)
        with c2: kpi("Licenciadas", f"{lic_n:,}", "", "or")
        if "estado_licencia" in df_sw.columns:
            venc = len(df_sw[df_sw["estado_licencia"] == "Vencida"])
            pv   = len(df_sw[df_sw["estado_licencia"] == "Por vencer (<60d)"])
            with c3: kpi("Vencidas",    f"{venc}", "requieren renovación", "rd")
            with c4: kpi("Por Vencer",  f"{pv}",  "en menos de 60 días",  "or")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Fila 1: Categoría + Estado ────────────────────────────
        g1, g2 = st.columns(2)
        with g1:
            if "CATEGORIA" in df_sw.columns:
                cat = df_sw["CATEGORIA"].value_counts().reset_index()
                cat.columns = ["Categoría", "Cantidad"]
                fig = px.pie(cat, names="Categoría", values="Cantidad",
                             color_discrete_sequence=[BLUE, ORANGE],
                             title="Software Educativo vs Operacional", hole=0.4)
                fig.update_layout(margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig, use_container_width=True)

        with g2:
            if "estado_licencia" in df_sw.columns:
                est = df_sw["estado_licencia"].value_counts().reset_index()
                est.columns = ["Estado", "Cantidad"]
                cmap = {"Vigente": GREEN, "Por vencer (<60d)": ORANGE,
                        "Vencida": RED, "Sin fecha": GREY}
                fig2 = px.bar(est, x="Estado", y="Cantidad", color="Estado",
                              color_discrete_map=cmap, title="Estado de Licencias", text="Cantidad")
                fig2.update_traces(textposition="outside")
                fig2.update_layout(showlegend=False, plot_bgcolor="white",
                                   margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig2, use_container_width=True)

        # ── Proveedores ───────────────────────────────────────────
        if "PROVEEDOR" in df_sw.columns:
            top_prov = df_sw["PROVEEDOR"].value_counts().head(10).reset_index()
            top_prov.columns = ["Proveedor", "Cantidad"]
            fig3 = px.bar(top_prov, x="Cantidad", y="Proveedor", orientation="h",
                          color_discrete_sequence=[GREY],
                          title="Top Proveedores de Software", text="Cantidad")
            fig3.update_traces(textposition="outside")
            fig3.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=350)
            st.plotly_chart(fig3, use_container_width=True)

        # ── NUEVO: Timeline de vencimientos ───────────────────────
        if "FECHA DE VENCIMIENTO" in df_sw.columns and "NOMBRE SOFTWARE" in df_sw.columns:
            st.markdown("#### 📅 Timeline de Vencimiento de Licencias")
            sw_fechas = df_sw.dropna(subset=["FECHA DE VENCIMIENTO"]).copy()
            sw_fechas = sw_fechas.sort_values("FECHA DE VENCIMIENTO")
            cmap2 = {"Vigente": GREEN, "Por vencer (<60d)": ORANGE,
                     "Vencida": RED, "Sin fecha": GREY}
            if "estado_licencia" in sw_fechas.columns:
                fig_tl = px.scatter(sw_fechas, x="FECHA DE VENCIMIENTO", y="NOMBRE SOFTWARE",
                                    color="estado_licencia", color_discrete_map=cmap2,
                                    title="Vencimiento por Licencia",
                                    hover_data=["CATEGORIA","PROVEEDOR"] if "PROVEEDOR" in sw_fechas.columns else ["CATEGORIA"])
                fig_tl.update_traces(marker=dict(size=10))
                fig_tl.update_layout(margin=dict(t=50,b=10), height=max(300, len(sw_fechas)*18))
                st.plotly_chart(fig_tl, use_container_width=True)

        # ── NUEVO: Categoría × Estado (stacked bar) ───────────────
        if "CATEGORIA" in df_sw.columns and "estado_licencia" in df_sw.columns:
            cat_est = (df_sw.groupby(["CATEGORIA","estado_licencia"]).size()
                       .reset_index(name="Cantidad"))
            fig_ce = px.bar(cat_est, x="CATEGORIA", y="Cantidad", color="estado_licencia",
                            color_discrete_map={"Vigente": GREEN, "Por vencer (<60d)": ORANGE,
                                                "Vencida": RED, "Sin fecha": GREY},
                            title="Estado de Licencias por Categoría", barmode="stack", text_auto=True)
            fig_ce.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=320)
            st.plotly_chart(fig_ce, use_container_width=True)

        # ── Tabla de urgentes ─────────────────────────────────────
        if "estado_licencia" in df_sw.columns:
            urgentes = df_sw[df_sw["estado_licencia"].isin(["Vencida","Por vencer (<60d)"])]
            if not urgentes.empty:
                st.markdown("#### 📋 Licencias que requieren atención inmediata")
                cols_sw = [c for c in ["NOMBRE SOFTWARE","CATEGORIA","PROVEEDOR",
                                        "FECHA DE VENCIMIENTO","dias_para_vencer","estado_licencia"]
                           if c in urgentes.columns]
                st.dataframe(urgentes[cols_sw].sort_values("dias_para_vencer"),
                             use_container_width=True, height=280)

# ═══════════════════════════════════════════════════════════
# TAB 4 — ACTIVIDADES DE DESARROLLO
# ═══════════════════════════════════════════════════════════
with tab4:
    if not uploaded_actividades or act.empty:
        st.info("👈 Sube el archivo de Actividades para ver esta sección.")
    else:
        # ── Filtros propios de esta pestaña ──────────────────────
        with st.sidebar:
            st.markdown('<div class="sidebar-section"><h4>🛠️ Filtros — Actividades</h4>', unsafe_allow_html=True)

            # Rango de fechas
            if "Fecha de creación" in act.columns:
                fmin_a = act["Fecha de creación"].min()
                fmax_a = act["Fecha de creación"].max()
                if pd.notna(fmin_a) and pd.notna(fmax_a):
                    rango_act = st.date_input(
                        "Rango de creación:",
                        value=(fmin_a.date(), fmax_a.date()),
                        min_value=fmin_a.date(), max_value=fmax_a.date(),
                        key="act_fecha"
                    )
                else:
                    rango_act = None
            else:
                rango_act = None

            tipos_act  = sorted(act["Tipo"].dropna().unique())         if "Tipo"         in act.columns else []
            estados_act = sorted(act["Estado"].dropna().unique())      if "Estado"       in act.columns else []
            devs_act   = sorted(act["Desarrollador"].dropna().unique()) if "Desarrollador" in act.columns else []

            tipos_a_sel  = st.multiselect("Tipo:",         tipos_act,   default=tipos_act,   key="act_tipo")
            estados_a_sel = st.multiselect("Estado:",       estados_act, default=estados_act, key="act_est")
            devs_sel      = st.multiselect("Desarrollador:", devs_act,  default=devs_act,    key="act_dev")
            st.markdown('</div>', unsafe_allow_html=True)

        df_a = act.copy()
        if rango_act and len(rango_act) == 2 and "Fecha de creación" in df_a.columns:
            df_a = df_a[
                (df_a["Fecha de creación"].dt.date >= rango_act[0]) &
                (df_a["Fecha de creación"].dt.date <= rango_act[1])
            ]
        if tipos_a_sel   and "Tipo"          in df_a.columns: df_a = df_a[df_a["Tipo"].isin(tipos_a_sel)]
        if estados_a_sel and "Estado"         in df_a.columns: df_a = df_a[df_a["Estado"].isin(estados_a_sel)]
        if devs_sel      and "Desarrollador"  in df_a.columns: df_a = df_a[df_a["Desarrollador"].isin(devs_sel)]

        st.caption(f"Mostrando **{len(df_a):,}** ítems de {len(act):,} totales.")

        # ── KPIs ─────────────────────────────────────────────────
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Total Ítems", f"{len(df_a):,}", "bugs + user stories")
        if "Tipo" in df_a.columns:
            bugs = len(df_a[df_a["Tipo"] == "Bug"])
            us   = len(df_a[df_a["Tipo"] == "User Story"])
            with c2: kpi("Bugs",         f"{bugs}", "defectos reportados", "rd")
            with c3: kpi("User Stories", f"{us}",   "funcionalidades",     "or")
        if "Estado" in df_a.columns:
            cerr = len(df_a[df_a["Estado"] == "Closed"])
            with c4: kpi("Cerrados", f"{cerr}", f"{cerr/max(len(df_a),1)*100:.1f}% completados", "gr")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Fila 1: Estado + Tipo ─────────────────────────────────
        g1, g2 = st.columns(2)
        with g1:
            if "Estado" in df_a.columns:
                est = df_a["Estado"].value_counts().reset_index()
                est.columns = ["Estado", "Cantidad"]
                cmap3 = {"Closed": GREEN, "Resolved": BLUE, "Active": ORANGE, "New": GREY}
                fig = px.bar(est, x="Cantidad", y="Estado", orientation="h",
                             color="Estado", color_discrete_map=cmap3,
                             title="Ítems por Estado", text="Cantidad")
                fig.update_traces(textposition="outside")
                fig.update_layout(showlegend=False, plot_bgcolor="white",
                                  margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig, use_container_width=True)

        with g2:
            if "Tipo" in df_a.columns:
                tipo = df_a["Tipo"].value_counts().reset_index()
                tipo.columns = ["Tipo", "Cantidad"]
                fig2 = px.pie(tipo, names="Tipo", values="Cantidad",
                              color_discrete_sequence=[ORANGE, BLUE],
                              title="Bugs vs User Stories", hole=0.4)
                fig2.update_layout(margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig2, use_container_width=True)

        # ── Carga por desarrollador ───────────────────────────────
        if "Desarrollador" in df_a.columns:
            g3, g4 = st.columns(2)
            with g3:
                by_dev = df_a["Desarrollador"].value_counts().reset_index()
                by_dev.columns = ["Desarrollador", "Ítems"]
                fig3 = px.bar(by_dev, x="Ítems", y="Desarrollador", orientation="h",
                              color_discrete_sequence=[BLUE],
                              title="Carga Total por Desarrollador", text="Ítems")
                fig3.update_traces(textposition="outside")
                fig3.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=300)
                st.plotly_chart(fig3, use_container_width=True)

            with g4:
                # NUEVO: bugs vs US por desarrollador
                if "Tipo" in df_a.columns:
                    dev_tipo = (df_a.groupby(["Desarrollador","Tipo"]).size()
                                .reset_index(name="Cantidad"))
                    fig4b = px.bar(dev_tipo, x="Cantidad", y="Desarrollador", color="Tipo",
                                   color_discrete_sequence=[ORANGE, BLUE],
                                   orientation="h", barmode="stack",
                                   title="Bugs vs User Stories por Desarrollador")
                    fig4b.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10), height=300)
                    st.plotly_chart(fig4b, use_container_width=True)

        # ── Retraso ───────────────────────────────────────────────
        if "dias_retraso" in df_a.columns:
            act_cerr = df_a[df_a["Estado"] == "Closed"].dropna(subset=["dias_retraso"])
            if not act_cerr.empty:
                g5, g6 = st.columns(2)
                with g5:
                    fig5 = px.histogram(act_cerr, x="dias_retraso", nbins=20,
                                        color_discrete_sequence=[ORANGE],
                                        title="Distribución de Días de Retraso (ítems cerrados)")
                    fig5.add_vline(x=0, line_dash="dash", line_color="red",
                                   annotation_text="Sin retraso")
                    fig5.update_layout(plot_bgcolor="white", margin=dict(t=40,b=10))
                    st.plotly_chart(fig5, use_container_width=True)

                with g6:
                    # NUEVO: Retraso por desarrollador (boxplot)
                    if "Desarrollador" in act_cerr.columns:
                        fig6 = px.box(act_cerr, x="Desarrollador", y="dias_retraso",
                                      color_discrete_sequence=[BLUE],
                                      title="Retraso por Desarrollador",
                                      points="outliers")
                        fig6.add_hline(y=0, line_dash="dash", line_color="red",
                                       annotation_text="Sin retraso")
                        fig6.update_layout(plot_bgcolor="white",
                                           xaxis_tickangle=-30,
                                           margin=dict(t=40,b=10))
                        st.plotly_chart(fig6, use_container_width=True)

        # ── NUEVO: Evolución de creación de ítems por semana ──────
        if "Fecha de creación" in df_a.columns:
            df_a["semana_act"] = df_a["Fecha de creación"].dt.to_period("W").astype(str)
            by_week = (df_a.groupby(["semana_act","Tipo"]).size()
                       .reset_index(name="Cantidad") if "Tipo" in df_a.columns
                       else df_a.groupby("semana_act").size().reset_index(name="Cantidad"))
            if "Tipo" in df_a.columns:
                fig_wk = px.line(by_week.sort_values("semana_act"),
                                 x="semana_act", y="Cantidad", color="Tipo",
                                 color_discrete_sequence=[ORANGE, BLUE],
                                 markers=True,
                                 title="Ítems Creados por Semana (Bugs vs User Stories)")
            else:
                fig_wk = px.line(by_week.sort_values("semana_act"),
                                 x="semana_act", y="Cantidad",
                                 color_discrete_sequence=[BLUE], markers=True,
                                 title="Ítems Creados por Semana")
            fig_wk.update_layout(plot_bgcolor="white", xaxis_title="Semana",
                                  xaxis_tickangle=-45, margin=dict(t=40,b=10))
            st.plotly_chart(fig_wk, use_container_width=True)
