import os
import datetime
import pandas as pd

# =============================================================================
# 1. CONFIGURACIÓN DE RUTAS (Ecosistema del Bot)
# =============================================================================
RUTA_BASE = r"C:\RPA_Validacion_Ventas"
PATH_VENTAS = os.path.join(RUTA_BASE, "02_Temporal", "Ventas_Extraidas.xlsx")
PATH_COBERTURA = os.path.join(RUTA_BASE, "01_Insumos", "Matriz_Cobertura.xlsx")
PATH_DEUDORES = os.path.join(RUTA_BASE, "01_Insumos", "Base_Clientes_Deudores.xlsx")
PATH_TARIFARIO = os.path.join(RUTA_BASE, "01_Insumos", "Tarifario_Oficial.xlsx")
PATH_EVIDENCIAS = os.path.join(RUTA_BASE, "03_Evidencias")
PATH_REPORTES = os.path.join(RUTA_BASE, "04_Reportes")

print("🤖 BOT: Iniciando proceso de validación multidimensional...")

# =============================================================================
# 2. CARGA DE ARCHIVOS MAESTROS EN MEMORIA (Simulación de Sistemas)
# =============================================================================
try:
    df_ventas = pd.read_excel(PATH_VENTAS)
    df_distritos = pd.read_excel(PATH_COBERTURA, sheet_name="Distritos_Fibra")
    df_deudores = pd.read_excel(PATH_DEUDORES)
    df_tarifario = pd.read_excel(PATH_TARIFARIO)
except Exception as e:
    print(f"❌ ERROR CRÍTICO: No se pudieron cargar los insumos. Detalle: {e}")
    exit()

# Listas para almacenar los resultados del procesamiento
estados_finales = []
motivos_observacion = []

# =============================================================================
# 3. BUCLE TRANSACCIONAL (Procesamiento Fila por Fila)
# =============================================================================
for index, fila in df_ventas.iterrows():
    id_venta = str(fila['ID_Venta']).strip()
    dni = int(fila['Cliente_Documento'])
    distrito_cliente = str(fila['Distrito']).strip()
    plan_solicitado = str(fila['Plan_Solicitado']).strip()
    precio_pactado = float(fila['Precio_Pactado'])
    
    print(f"\n────────────────────────────────────────────────────────")
    print(f"📦 Procesando Transacción: {id_venta} | Cliente DNI: {dni}")
    
    # Inicializamos variables de control por cada fila
    venta_aprobada = True
    motivo = "Cumple con todas las condiciones."

    # 🛑 REGLA 1: VALIDACIÓN DE COBERTURA DE FIBRA (Filtro Técnico)
    # Buscamos el distrito en el maestro de cobertura
    cobertura_distrito = df_distritos[df_distritos['Distrito'].str.lower() == distrito_cliente.lower()]
    
    if cobertura_distrito.empty:
        venta_aprobada = False
        motivo = f"RECHAZADO: El distrito '{distrito_cliente}' no cuenta con despliegue de red."
    elif cobertura_distrito.iloc[0]['Estado_Red'].upper() == "SATURADO":
        venta_aprobada = False
        motivo = f"RECHAZADO: Red saturada (sin bornes disponibles) en {distrito_cliente}."
        
    if not venta_aprobada:
        estados_finales.append("RECHAZADO")
        motivos_observacion.append(motivo)
        print(f" ✘ {motivo}")
        continue  # Salta a la siguiente fila inmediatamente (Regla de negocio prioritaria)

    # 🛑 REGLA 2: VALIDACIÓN DE RIESGO / DEUDAS (Filtro Financiero)
    # Buscamos si el DNI figura en la lista de deudores morosos o bloqueados
    registro_deuda = df_deudores[df_deudores['Documento_Identidad'] == dni]
    
    if not registro_deuda.empty:
        estado_financiero = str(registro_deuda.iloc[0]['Estado_Financiero']).upper()
        monto_deuda = float(registro_deuda.iloc[0]['Monto_Deuda'])
        
        if estado_financiero in ["MOROSO", "BLOQUEADO"] or monto_deuda > 0:
            venta_aprobada = False
            motivo = f"RECHAZADO: Cliente figura como {estado_financiero}. Deuda pendiente: S/. {monto_deuda:.2f}"
            estados_finales.append("RECHAZADO")
            motivos_observacion.append(motivo)
            print(f" ✘ {motivo}")
            continue

    # 🛑 REGLA 3: AUDITORÍA DE TARIFAS Y PRECIOS (Filtro Comercial)
    # Verificamos que el plan exista y el precio pactado por el asesor sea el correcto
    registro_plan = df_tarifario[df_tarifario['Plan_Nombre'].str.lower() == plan_solicitado.lower()]
    
    if registro_plan.empty:
        venta_aprobada = False
        motivo = f"OBSERVADO: El plan '{plan_solicitado}' no existe en el tarifario vigente."
    else:
        precio_oficial = float(registro_plan.iloc[0]['Precio_Lista'])
        if precio_pactado != precio_oficial:
            venta_aprobada = False
            motivo = f"OBSERVADO: Inconsistencia de precio. Pactado: S/. {precio_pactado:.2f} | Oficial: S/. {precio_oficial:.2f}"

    # 🛑 REGLA 4: AUDITORÍA DE EVIDENCIAS DIGITALES (Filtro de Cumplimiento Legal)
    # Si sigue aprobada o solo observada por precio, auditamos que existan sus archivos
    ruta_carpeta_evidencia = os.path.join(PATH_EVIDENCIAS, id_venta)
    archivos_obligatorios = ["DNI.pdf", "Contrato.pdf", "Audio.mp3"]
    archivos_faltantes = []

    if os.path.exists(ruta_carpeta_evidencia):
        archivos_en_carpeta = os.listdir(ruta_carpeta_evidencia)
        for archivo in archivos_obligatorios:
            if archivo not in archivos_en_carpeta:
                archivos_faltantes = []
                archivos_faltantes.append(archivo)
    else:
        archivos_faltantes = archivos_obligatorios

    if archivos_faltantes:
        venta_aprobada = False
        if "OBSERVADO" in motivo:
            motivo += f" | Faltan documentos: {', '.join(archivos_faltantes)}"
        else:
            motivo = f"OBSERVADO: Falta documentación obligatoria en repositorio: {', '.join(archivos_faltantes)}"

    # =============================================================================
    # 4. ASIGNACIÓN DEL VEREDICTO FINAL DE LA FILA
    # =============================================================================
    if venta_aprobada:
        estados_finales.append("APROBADO")
        motivos_observacion.append(motivo)
        print(f" ✔ APROBADO: Factibilidad e integridad confirmadas de forma exitosa.")
    else:
        estados_finales.append("OBSERVADO")
        motivos_observacion.append(motivo)
        print(f" ⚠ {motivo}")

# =============================================================================
# 5. CIERRE DE CICLO: GENERACIÓN DEL REPORTE DE TRAZABILIDAD (LOG)
# =============================================================================
df_ventas['Resultado_RPA'] = estados_finales
df_ventas['Detalle_Validacion'] = motivos_observacion
df_ventas['Fecha_Procesamiento'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Guardar informe de auditoría final
fecha_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
nombre_reporte = f"Log_Ejecucion_{fecha_str}.xlsx"
PATH_FINAL_REPORTE = os.path.join(PATH_REPORTES, nombre_reporte)

df_ventas.to_excel(PATH_FINAL_REPORTE, index=False)

print(f"\n🏁 PROCESO FINALIZADO. Reporte de trazabilidad generado con éxito en:")
print(f"📂 {PATH_FINAL_REPORTE}")