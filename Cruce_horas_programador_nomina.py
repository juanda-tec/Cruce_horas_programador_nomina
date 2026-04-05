import pandas as pd
import os
from datetime import datetime

# --- 1. CONFIGURACIÓN DE RUTAS ---
mi_usuario = "juanhes" 
ruta_base = rf'C:\Users\{mi_usuario}\Proyectos Desarrollo Soft\TDS\Cruce_horas_programador_nomina\ArchivosCruce'

archivo_prog = os.path.join(ruta_base, 'Programador.xlsm')
archivo_nom = os.path.join(ruta_base, 'Nomina.csv')

def redondear_y_normalizar_hora(hora_str):
    if pd.isna(hora_str) or str(hora_str).strip() == "": return ""
    try:
        partes = str(hora_str).strip().split(':')
        h, m = int(partes[0]), int(partes[1])
        if m in [29, 59, 14, 44]:
            m += 1
            if m == 60: m = 0; h += 1
        return f"{h:02d}:{m:02d}"
    except: return str(hora_str).strip()[:5]

def ejecutar_conciliacion_v9():
    print("🚀 Generando Informe v9 con ajustes de limpieza y valores absolutos...")

    # --- 2. CARGA ---
    df_p = pd.read_excel(archivo_prog, dtype=str)
    df_n = pd.read_csv(archivo_nom, sep=';', skiprows=[1], dtype=str)

    df_p.columns = df_p.columns.str.strip()
    df_n.columns = df_n.columns.str.strip()

    # Filtro y Normalización de Concepto 450
    df_p = df_p[df_p['concept'].astype(str).str.contains('450')].copy()
    df_n = df_n[df_n['timeType'].astype(str).str.contains('450')].copy()
    
    df_p['concept_norm'] = df_p['concept'].astype(str).str.lstrip('0')
    df_n['concept_norm'] = df_n['timeType'].astype(str).str.lstrip('0')

    # --- 3. NORMALIZACIÓN ---
    df_p['h_ini_norm'] = df_p['date_from'].apply(redondear_y_normalizar_hora)
    df_p['h_fin_norm'] = df_p['date_until'].apply(redondear_y_normalizar_hora)
    df_n['h_ini_norm'] = df_n['startTime'].apply(redondear_y_normalizar_hora)
    df_n['h_fin_norm'] = df_n['endTime'].apply(redondear_y_normalizar_hora)

    df_p['fecha_obj'] = pd.to_datetime(df_p['start_date'], errors='coerce')
    df_n['fecha_obj'] = pd.to_datetime(df_n['startDate'], dayfirst=True, errors='coerce')

    def calc_h(h1, h2):
        try:
            t1 = pd.to_datetime(h1, format='%H:%M')
            t2 = pd.to_datetime(h2, format='%H:%M')
            return round((t2 - t1).total_seconds() / 3600, 2)
        except: return 0.0

    df_p['horas_val'] = df_p.apply(lambda x: calc_h(x['h_ini_norm'], x['h_fin_norm']), axis=1)
    df_n['horas_val'] = df_n.apply(lambda x: calc_h(x['h_ini_norm'], x['h_fin_norm']), axis=1)

    # --- 4. CRUCE ---
    detalle = pd.merge(
        df_p[['bp', 'fecha_obj', 'h_ini_norm', 'h_fin_norm', 'horas_val', 'concept_norm']],
        df_n[['userId', 'fecha_obj', 'h_ini_norm', 'h_fin_norm', 'horas_val', 'concept_norm']],
        left_on=['bp', 'fecha_obj', 'h_ini_norm'],
        right_on=['userId', 'fecha_obj', 'h_ini_norm'],
        how='outer', suffixes=('_P', '_N')
    )

    # Lógica de Auditoría (Siempre Positivos)
    detalle['Falta_Prog'] = detalle.apply(lambda r: r['horas_val_N'] if pd.isna(r['bp']) else 0, axis=1)
    detalle['Falta_Nom'] = detalle.apply(lambda r: r['horas_val_P'] if pd.isna(r['userId']) else 0, axis=1)

    # --- 5. PESTAÑA: BALANCE (Solo registros con diferencias y en positivo) ---
    meses_es = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

    detalle['Año'] = detalle['fecha_obj'].dt.year
    detalle['Mes'] = detalle['fecha_obj'].dt.month.map(meses_es)
    detalle['Quincena'] = detalle['fecha_obj'].apply(lambda x: 'Q1' if pd.notnull(x) and x.day <= 15 else 'Q2')
    detalle['ID_Docente'] = detalle['bp'].fillna(detalle['userId'])

    balance = detalle.groupby(['Año', 'Mes', 'Quincena', 'ID_Docente']).agg({
        'Falta_Prog': 'sum',
        'Falta_Nom': 'sum'
    }).reset_index()
    
    balance['Total Diferencia'] = balance['Falta_Prog'] + balance['Falta_Nom']
    # Filtrar: Solo mostrar si hay alguna diferencia
    balance = balance[balance['Total Diferencia'] > 0].copy()
    balance = balance.rename(columns={'Falta_Prog': 'Falta Horas Programador (A)', 'Falta_Nom': 'Falta Horas Nómina (B)'})

    # --- 6. PESTAÑA: DETALLES (Llenado de campos vacíos) ---
    detalle['E_Prog'] = detalle['bp'].apply(lambda x: 'SI' if pd.notna(x) else 'NO')
    detalle['E_Nom'] = detalle['userId'].apply(lambda x: 'SI' if pd.notna(x) else 'NO')
    
    def marcar_dif(r):
        if r['E_Prog'] == 'SI' and r['E_Nom'] == 'SI': return 'Ninguna'
        return 'Falta horas nomina' if r['E_Prog'] == 'SI' else 'Falta horas programador'

    detalle['Diferencia_Status'] = detalle.apply(marcar_dif, axis=1)
    
    # Llenado crítico de vacíos
    detalle['date_until_final'] = detalle['h_fin_norm_P'].fillna(detalle['h_fin_norm_N'])
    detalle['Concepto_final'] = detalle['concept_norm_P'].fillna(detalle['concept_norm_N'])
    detalle['horas_final'] = detalle['horas_val_P'].fillna(detalle['horas_val_N'])

    res_detalle = detalle[[
        'ID_Docente', 'fecha_obj', 'h_ini_norm', 'date_until_final', 'Concepto_final', 
        'horas_final', 'E_Prog', 'E_Nom', 'Diferencia_Status'
    ]].copy()
    res_detalle['fecha_obj'] = res_detalle['fecha_obj'].dt.strftime('%Y/%m/%d')
    res_detalle.columns = ['bp', 'fecha_dt', 'date_from', 'date_until', 'Concepto', 'horas_Prog', 'Esta_en_programador(Si/NO)', 'Esta_en_nomina(Si/NO)', 'Diferencia']

    # --- 7. PESTAÑA: RESUMEN EJECUTIVO (Métricas solicitadas) ---
    resumen_ejecutivo = pd.DataFrame({
        'Métrica': [
            'Fecha de ejecución',
            'Total Horas Programadas',
            'Total Horas Nómina',
            'Total horas diferencias'
        ],
        'Valor': [
            datetime.now().strftime('%Y-%m-%d %H:%M'),
            df_p['horas_val'].sum(),
            df_n['horas_val'].sum(),
            balance['Total Diferencia'].sum()
        ]
    })

    # --- 8. EXPORTACIÓN ---
    ahora = datetime.now().strftime('%Y%m%d_%H%M')
    ruta_salida = os.path.join(ruta_base, f'Informe_Final_Ajustado_{ahora}.xlsx')

    with pd.ExcelWriter(ruta_salida) as writer:
        resumen_ejecutivo.to_excel(writer, sheet_name='Resumen Ejecutivo', index=False)
        balance.to_excel(writer, sheet_name='Balance de Horas por Docente', index=False)
        res_detalle.to_excel(writer, sheet_name='Diferencias Detalladas', index=False)

    print(f"✅ Informe v9 generado con éxito: {os.path.basename(ruta_salida)}")

if __name__ == "__main__":
    ejecutar_conciliacion_v9()