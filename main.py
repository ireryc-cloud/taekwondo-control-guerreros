from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from datetime import date
import psycopg2
from psycopg2.extras import RealDictCursor
import io
import openpyxl

app = FastAPI(
    title="Taekwondo Control Guerreros API",
    description="Backend unificado para la gestión integral del Doyang.",
    version="3.0.0"
)

# Configuración de tu base de datos PostgreSQL
DB_CONFIG = {
    "dbname": "taekwondo_db",
    "user": "postgres",
    "password": "tu_password",
    "host": "localhost",
    "port": "5432"
}

def get_db_connection():
    return psycopg2.connect(**DB_CONFIG)

# --- MODELOS DE DATOS (PYDANTIC) ---
class AlumnoCreate(BaseModel):
    nombre: str
    apellido: str
    fecha_nacimiento: date
    grado_actual_id: int
    foto_url: str = None

class CobroCreate(BaseModel):
    id_alumno: int
    concepto: str
    monto: float
    fecha_vencimiento: date
    estado: str = "Pendiente"

class EvaluacionCreate(BaseModel):
    id_alumno: int
    grado_evaluado_id: int
    fecha_evaluacion: date
    calificacion_tecnica: float
    calificacion_combate: float
    resultado: str = "Aprobado"

class AsistenciaCreate(BaseModel):
    id_alumno: int
    tipo_clase: str = "Clase Regular"

class ProductoCreate(BaseModel):
    nombre_producto: str
    categoria: str
    talla: str
    precio_venta: float
    stock_actual: int

class VentaCreate(BaseModel):
    id_alumno: int = None
    id_producto: int
    cantidad: int

class TorneoCreate(BaseModel):
    nombre_torneo: str
    fecha_evento: date
    sede: str

class EgresoCreate(BaseModel):
    concepto: str
    monto: float
    fecha_egreso: date

# --- ENDPOINTS: ALUMNOS ---
@app.get("/api/alumnos/activos", summary="Listar alumnos activos con cinta y foto")
def get_alumnos_activos():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cursor.execute("""
            SELECT a.*, g.nombre_cinta AS cinta_actual 
            FROM alumnos a 
            JOIN grados g ON a.grado_actual_id = g.id_grado 
            WHERE a.activo = TRUE 
            ORDER BY g.orden ASC;
        """)
        return {"status": "success", "data": cursor.fetchall()}
    finally:
        cursor.close(); conn.close()

@app.post("/api/alumnos", summary="Dar de alta alumno")
def crear_alumno(alumno: AlumnoCreate):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO alumnos (nombre, apellido, fecha_nacimiento, grado_actual_id, foto_url, activo) 
            VALUES (%s, %s, %s, %s, %s, TRUE);
        """, (alumno.nombre, alumno.apellido, alumno.fecha_nacimiento, alumno.grado_actual_id, alumno.foto_url))
        conn.commit()
        return {"status": "success", "message": "Alumno registrado con éxito"}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        cursor.close(); conn.close()

# --- ENDPOINTS: FINANZAS Y COBROS ---
@app.get("/api/cobros/morosos", summary="Reporte de alumnos morosos")
def get_morosos():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cursor.execute("""
            SELECT a.id_alumno, a.nombre, a.apellido, c.concepto, c.monto, c.fecha_vencimiento, 
                   (CURRENT_DATE - c.fecha_vencimiento) AS dias_de_retraso 
            FROM cobros c 
            JOIN alumnos a ON c.id_alumno = a.id_alumno 
            WHERE c.estado = 'Pendiente' AND c.fecha_vencimiento < CURRENT_DATE;
        """)
        return {"status": "success", "data": cursor.fetchall()}
    finally:
        cursor.close(); conn.close()

@app.post("/api/cobros", summary="Registrar cobro")
def crear_cobro(cobro: CobroCreate):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO cobros (id_alumno, concepto, monto, fecha_vencimiento, estado) 
            VALUES (%s, %s, %s, %s, %s);
        """, (cobro.id_alumno, cobro.concepto, cobro.monto, cobro.fecha_vencimiento, cobro.estado))
        conn.commit()
        return {"status": "success", "message": "Cobro registrado"}
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        cursor.close(); conn.close()

@app.get("/api/finanzas/balance-mensual", summary="Balance financiero mensual (Ingresos vs Egresos)")
def get_balance_mensual():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        query = """
            SELECT 
                COALESCE(i.mes, e.mes) AS mes,
                COALESCE(i.total_ingresos, 0) AS total_ingresos,
                COALESCE(e.total_egresos, 0) AS total_egresos,
                (COALESCE(i.total_ingresos, 0) - COALESCE(e.total_egresos, 0)) AS balance_neto
            FROM (
                SELECT TO_CHAR(COALESCE(fecha_pago, fecha_vencimiento), 'YYYY-MM') AS mes, SUM(monto) AS total_ingresos
                FROM cobros WHERE estado = 'Pagado'
                GROUP BY TO_CHAR(COALESCE(fecha_pago, fecha_vencimiento), 'YYYY-MM')
            ) i
            FULL OUTER JOIN (
                SELECT TO_CHAR(fecha_egreso, 'YYYY-MM') AS mes, SUM(monto) AS total_egresos
                FROM egresos
                GROUP BY TO_CHAR(fecha_egreso, 'YYYY-MM')
            ) e ON i.mes = e.mes
            ORDER BY mes DESC;
        """
        cursor.execute(query)
        return {"status": "success", "data": cursor.fetchall()}
    finally:
        cursor.close(); conn.close()

@app.post("/api/egresos", summary="Registrar egreso o gasto operativo")
def registrar_egreso(e: EgresoCreate):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO egresos (concepto, monto, fecha_egreso) VALUES (%s, %s, %s);", 
                       (e.concepto, e.monto, e.fecha_egreso))
        conn.commit()
        return {"status": "success", "message": "Egreso registrado"}
    finally:
        cursor.close(); conn.close()

# --- ENDPOINTS: ASISTENCIAS ---
@app.post("/api/asistencias", summary="Pase de asistencia rápido por ID")
def registrar_asistencia(asist: AsistenciaCreate):
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cursor.execute("SELECT nombre, apellido FROM alumnos WHERE id_alumno = %s AND activo = TRUE;", (asist.id_alumno,))
        alumno = cursor.fetchone()
        if not alumno:
            raise HTTPException(status_code=404, detail="Alumno no encontrado o inactivo")
        
        cursor.execute("INSERT INTO asistencias (id_alumno, tipo_clase) VALUES (%s, %s);", (asist.id_alumno, asist.tipo_clase))
        conn.commit()
        return {"status": "success", "message": f"¡Asistencia registrada para {alumno['nombre']} {alumno['apellido']}!"}
    finally:
        cursor.close(); conn.close()

# --- ENDPOINTS: INVENTARIO Y TIENDA ---
@app.get("/api/productos", summary="Listar inventario de tienda")
def get_productos():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cursor.execute("SELECT * FROM productos ORDER BY categoria;")
        return {"status": "success", "data": cursor.fetchall()}
    finally:
        cursor.close(); conn.close()

# --- ENDPOINTS: AVISOS Y EXCEL ---
@app.get("/api/avisos", summary="Obtener mural de avisos")
def get_avisos():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cursor.execute("SELECT * FROM avisos ORDER BY fecha_evento ASC NULLS LAST;")
        return {"status": "success", "data": cursor.fetchall()}
    finally:
        cursor.close(); conn.close()

@app.get("/api/reportes/exportar-excel", summary="Exportar reporte general a Excel")
def exportar_excel():
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alumnos Activos"
    ws.append(["ID", "Nombre", "Apellido", "Cinta", "Inscripción"])
    
    cursor.execute("""
        SELECT a.id_alumno, a.nombre, a.apellido, g.nombre_cinta, a.fecha_inscripcion 
        FROM alumnos a JOIN grados g ON a.grado_actual_id = g.id_grado WHERE a.activo = TRUE;
    """)
    for r in cursor.fetchall():
        ws.append([r['id_alumno'], r['nombre'], r['apellido'], r['nombre_cinta'], str(r['fecha_inscripcion'])])
        
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    cursor.close(); conn.close()
    
    return StreamingResponse(
        excel_io, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Taekwondo_Control_Guerreros.xlsx"}
    )
