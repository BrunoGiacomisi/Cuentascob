"""
Lógica para generar el Excel semanal "CUENTAS A COBRAR".

Lee los 3 Excel exportados del ERP (Agencia Marítima, Intercompany, FFWW),
los parsea y arma un workbook consolidado con 4 hojas.
Este módulo no tiene nada de interfaz: lo usa app.py (Streamlit) y los tests.
"""

from __future__ import annotations

import calendar
import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

# Etiquetas de cada archivo de entrada (las mismas que ve el usuario en la UI).
ETIQUETA_AGENCIA = "1. Agencia Marítima"
ETIQUETA_INTERCOMPANY = "2. Intercompany"
ETIQUETA_FFWW = "3. FFWW"

# Nombres de las hojas de salida, en el orden en que deben aparecer.
HOJA_AGENCIA_USD = "AGENCIA MARITIMA USD"
HOJA_INTERCOMPANY = "INTERCOMPANY"
HOJA_AGENCIA_UYU = "AGENCIA MARITIMA UYU"
HOJA_FFWW = "FFWW"
# Hojas adicionales, al final del libro.
HOJA_RECLAMAR = "Posibles cobranzas a reclamar"
HOJA_ANALISIS = "Análisis y gráficas"

# Un documento se considera "a reclamar" con más de estos días de vencido.
DIAS_PARA_RECLAMAR = 30
# Tramos de antigüedad (según días de atraso a la fecha del reporte).
TRAMOS = ["Por vencer", "0 - 30", "31 - 60", "61 - 90", "91 - 120", "121+"]

# Columnas que necesitamos de los archivos de entrada (se buscan por nombre).
COL_CODIGO = "Código de cliente"
COL_NOMBRE = "Nombre del cliente"
COL_ACUERDO = "Acuerdo global"
COL_TIPO = "Tipo"
COL_FECHA_CONTAB = "Fecha de contabilización"
COL_FECHA_VENC = "Fecha de vencimiento"
COL_REFERENCIA = "Número de referencia del SN"
COL_SALDO = "Saldo vencido"
COLUMNAS_REQUERIDAS = [
    COL_CODIGO, COL_NOMBRE, COL_ACUERDO, COL_TIPO,
    COL_FECHA_CONTAB, COL_FECHA_VENC, COL_REFERENCIA, COL_SALDO,
]

MONEDAS_VALIDAS = ("USD", "UYU")
# Único cliente que hoy tiene documentos en pesos (no genera warning).
CLIENTE_UYU_ESPERADO = "FERSPO"

# Encabezados de la salida.
ENCABEZADOS = [
    "Nombre del cliente", "Fecha de contabilización", "Fecha de vencimiento",
    "Días vencido", "Número de referencia del SN", "Saldo vencido",
]
# Posición de cada dato en las hojas de salida (1 = columna A).
COLUMNA_CONTAB, COLUMNA_VENC, COLUMNA_DIAS, COLUMNA_REFERENCIA, COLUMNA_MONTO = 2, 3, 4, 5, 6
LETRA_MONTO = "F"
ENCABEZADOS_EXTRA_FFWW = ["COMENTARIOS COMERCIAL", "COMENTARIO ADMINISTRACIÓN"]

# Estilos de la salida.
FUENTE = "Aptos Narrow"
TAMANO_FUENTE = 11
COLOR_ENCABEZADO = "F6C6AD"
COLOR_VENCIDO = "C0504D"  # rojo suave: días vencido > DIAS_PARA_RECLAMAR con saldo a cobrar
FORMATO_FECHA = "dd/mm/yyyy"
FORMATO_MONTO = {
    "USD": '_-[$USD]\\ * #,##0.00_-;\\-[$USD]\\ * #,##0.00_-;_-[$USD]\\ * "-"??_-;_-@_-',
    "UYU": "[$UYU]\\ #,##0.00",
}
FORMATO_TOTAL = {
    "USD": "[$USD]\\ #,##0.00",
    "UYU": "[$UYU]\\ #,##0.00",
}
ANCHOS = {"A": 50, "B": 24, "C": 22, "D": 13, "E": 28, "F": 19, "G": 27, "H": 31}
RELLENO_ENCABEZADO = PatternFill("solid", start_color=COLOR_ENCABEZADO, end_color=COLOR_ENCABEZADO)
BORDE_DOBLE = Border(bottom=Side(style="double"))


class ErrorArchivo(Exception):
    """Error con un mensaje pensado para mostrarle directamente al usuario."""


# ---------------------------------------------------------------------------
# Modelo de datos
# ---------------------------------------------------------------------------

@dataclass
class Documento:
    tipo: str
    fecha_contab: datetime | None
    fecha_venc: datetime | None
    referencia: str | None
    moneda: str
    monto: Decimal
    fila: int  # fila en el Excel fuente (para mensajes)


@dataclass
class Cliente:
    codigo: str | None
    nombre: str
    total_cabecera: object  # valor crudo de "Saldo vencido" en la cabecera
    fila: int
    documentos: list[Documento] = field(default_factory=list)

    @property
    def subtotal(self) -> Decimal:
        return sum((d.monto for d in self.documentos), Decimal("0"))


@dataclass
class Hoja:
    """Una hoja del Excel de salida, ya filtrada."""
    nombre: str
    moneda: str
    clientes: list[Cliente]
    con_comentarios: bool = False  # True solo en FFWW (columnas F y G)

    @property
    def cantidad_clientes(self) -> int:
        return len(self.clientes)

    @property
    def cantidad_documentos(self) -> int:
        return sum(len(c.documentos) for c in self.clientes)

    @property
    def total(self) -> Decimal:
        return sum((c.subtotal for c in self.clientes), Decimal("0"))


@dataclass
class Indicadores:
    """Números resumen de una hoja (o de un cliente) a la fecha del reporte."""
    nombre: str
    moneda: str
    clientes: int
    documentos: int
    total: Decimal
    vencido: Decimal  # saldo a cobrar con más de DIAS_PARA_RECLAMAR días de vencido
    atraso_promedio: int | None  # días, ponderado por monto (solo saldos a cobrar)
    vencimiento_mas_antiguo: datetime | None

    @property
    def porcentaje_vencido(self) -> float | None:
        return float(self.vencido / self.total) if self.total > 0 else None


@dataclass
class Resultado:
    contenido: bytes  # el .xlsx completo, listo para descargar
    nombre_archivo: str
    hojas: list[Hoja]  # las 4 hojas de datos
    avisos: list[str]
    fecha: date
    contenido_ffww: bytes  # versión para Comercial: solo FFWW
    nombre_archivo_ffww: str


# ---------------------------------------------------------------------------
# Parseo de valores sueltos
# ---------------------------------------------------------------------------

def parsear_monto(valor) -> tuple[str, Decimal]:
    """
    Convierte el texto de "Saldo vencido" en (moneda, monto).

    Ejemplos: "USD  1.472,20" -> ("USD", 1472.20)
              "USD (968,99)"  -> ("USD", -968.99)
              "UYU  200.488,28" -> ("UYU", 200488.28)
    Si viene como número, se asume USD (quien llama debe avisar al usuario).
    Lanza ValueError si el texto no se puede interpretar.
    """
    if isinstance(valor, bool):
        raise ValueError(f"valor inválido: {valor!r}")
    if isinstance(valor, (int, float, Decimal)):
        return "USD", Decimal(str(valor)).quantize(Decimal("0.01"))

    texto = str(valor or "").strip()
    moneda = texto[:3].upper()
    if moneda not in MONEDAS_VALIDAS:
        raise ValueError(f"moneda desconocida en {texto!r} (se esperaba USD o UYU)")

    numero = texto[3:].strip()
    negativo = numero.startswith("(") and numero.endswith(")")
    if negativo:
        numero = numero[1:-1].strip()
    if numero.startswith("-"):
        negativo = True
        numero = numero[1:].strip()

    # Formato local: punto = miles, coma = decimal.
    numero = numero.replace(".", "").replace(",", ".")
    if not re.fullmatch(r"\d+(\.\d+)?", numero):
        raise ValueError(f"no se pudo leer el monto {texto!r}")
    try:
        monto = Decimal(numero)
    except InvalidOperation:
        raise ValueError(f"no se pudo leer el monto {texto!r}") from None
    return moneda, -monto if negativo else monto


def parsear_fecha(valor) -> datetime | None:
    """Acepta datetime/date, o texto dd/mm/yyyy. Vacío -> None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day)
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            pass
    raise ValueError(f"fecha inválida {texto!r} (se esperaba dd/mm/yyyy)")


def _texto(valor) -> str | None:
    """Normaliza una celda de texto: None o vacío -> None."""
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def _normalizar(nombre) -> str:
    """Para comparar encabezados sin importar mayúsculas, tildes ni espacios."""
    texto = unicodedata.normalize("NFKD", str(nombre or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


# ---------------------------------------------------------------------------
# Lectura de un archivo del ERP
# ---------------------------------------------------------------------------

def leer_archivo(origen, etiqueta: str) -> tuple[list[Cliente], list[str]]:
    """
    Lee un Excel exportado del ERP y devuelve (clientes, avisos).

    `origen` puede ser una ruta o un objeto tipo archivo (lo que da Streamlit).
    `etiqueta` es el nombre del archivo para los mensajes (ej. "2. Intercompany").
    """
    try:
        wb = load_workbook(origen, data_only=True, read_only=True)
    except Exception:
        raise ErrorArchivo(
            f"No se pudo abrir el archivo cargado en «{etiqueta}». "
            "Verificá que sea un Excel (.xlsx) exportado del ERP."
        ) from None

    try:
        ws = wb.worksheets[0]
        filas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if not filas:
        raise ErrorArchivo(f"El archivo cargado en «{etiqueta}» está vacío.")

    # Ubicar cada columna por el nombre del encabezado (fila 1).
    posiciones = {_normalizar(v): i for i, v in enumerate(filas[0]) if v is not None}
    col = {}
    faltantes = []
    for nombre in COLUMNAS_REQUERIDAS:
        indice = posiciones.get(_normalizar(nombre))
        if indice is None:
            faltantes.append(nombre)
        col[nombre] = indice
    if faltantes:
        raise ErrorArchivo(
            f"Al archivo cargado en «{etiqueta}» le falta(n) la(s) columna(s): "
            + ", ".join(f"«{f}»" for f in faltantes)
            + ". ¿Es el export correcto del ERP? Revisá que no hayas cargado otro archivo en ese lugar."
        )

    def celda(fila, nombre):
        i = col[nombre]
        return fila[i] if i < len(fila) else None

    clientes: list[Cliente] = []
    avisos: list[str] = []
    actual: Cliente | None = None

    for nro, fila in enumerate(filas[1:], start=2):
        codigo = _texto(celda(fila, COL_CODIGO))
        nombre = _texto(celda(fila, COL_NOMBRE))
        tipo = _texto(celda(fila, COL_TIPO))

        # 1) Cabecera de cliente: tiene código y/o nombre.
        if codigo or nombre:
            actual = Cliente(
                codigo=codigo,
                nombre=nombre or codigo,
                total_cabecera=celda(fila, COL_SALDO),
                fila=nro,
            )
            clientes.append(actual)
            continue

        # 2) Filas sin tipo (las "N/A" o filas vacías): se ignoran.
        if not tipo:
            continue

        # 3) Documento: pertenece al último cliente leído.
        if actual is None:
            raise ErrorArchivo(
                f"«{etiqueta}», fila {nro}: hay un documento antes de cualquier cliente."
            )
        saldo = celda(fila, COL_SALDO)
        try:
            moneda, monto = parsear_monto(saldo)
            fecha_contab = parsear_fecha(celda(fila, COL_FECHA_CONTAB))
            fecha_venc = parsear_fecha(celda(fila, COL_FECHA_VENC))
        except ValueError as e:
            raise ErrorArchivo(f"«{etiqueta}», fila {nro} ({actual.nombre}): {e}.") from None

        if isinstance(saldo, (int, float)):
            avisos.append(
                f"«{etiqueta}», fila {nro} ({actual.nombre}): el saldo vino como número "
                f"sin moneda ({saldo}); se asumió USD."
            )

        actual.documentos.append(Documento(
            tipo=tipo,
            fecha_contab=fecha_contab,
            fecha_venc=fecha_venc,
            referencia=_texto(celda(fila, COL_REFERENCIA)),
            moneda=moneda,
            monto=monto,
            fila=nro,
        ))

    if not any(c.documentos for c in clientes):
        avisos.append(f"«{etiqueta}» no tiene ningún documento.")

    return clientes, avisos


def validar_cabeceras(clientes: list[Cliente], etiqueta: str) -> list[str]:
    """
    Compara el subtotal recalculado de cada cliente con el total que trae su
    cabecera en el fuente. Se saltean las cabeceras "****" (monedas mezcladas).
    """
    avisos = []
    for c in clientes:
        texto = _texto(c.total_cabecera)
        if texto is None or texto.startswith("*"):
            continue
        try:
            moneda, total_fuente = parsear_monto(c.total_cabecera)
        except ValueError:
            avisos.append(
                f"«{etiqueta}», {c.nombre}: no se pudo leer el total de cabecera ({texto!r})."
            )
            continue
        recalculado = sum(
            (d.monto for d in c.documentos if d.moneda == moneda), Decimal("0")
        )
        if recalculado != total_fuente:
            avisos.append(
                f"«{etiqueta}», {c.nombre}: la suma de sus documentos "
                f"({formatear_monto(recalculado, moneda)}) no coincide con el total "
                f"de la cabecera del ERP ({formatear_monto(total_fuente, moneda)})."
            )
    return avisos


def sumar_un_mes(fecha: date) -> date:
    """Misma fecha del mes siguiente (si no existe, el último día: 31/01 -> 28/02)."""
    anio, mes = (fecha.year + 1, 1) if fecha.month == 12 else (fecha.year, fecha.month + 1)
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return date(anio, mes, dia)


def validar_vencimientos(clientes: list[Cliente], etiqueta: str, fecha_reporte: date) -> list[str]:
    """
    En SAP el export se saca con "fecha de vencimiento hasta hoy + 1 mes".
    Si algún documento vence después de (fecha del reporte + 1 mes), el export
    se hizo con el parámetro incorrecto. Solo se avisa: el documento se incluye igual.
    """
    limite = sumar_un_mes(fecha_reporte)
    fuera = [
        (c, d) for c in clientes for d in c.documentos
        if d.fecha_venc is not None and d.fecha_venc.date() > limite
    ]
    if not fuera:
        return []
    detalle = "".join(
        f"\n- {c.nombre} – {d.referencia or 'sin referencia'} (vence {d.fecha_venc:%d/%m/%Y})"
        for c, d in fuera
    )
    return [
        f"«{etiqueta}»: {len(fuera)} documento(s) vencen después del {limite:%d/%m/%Y} "
        f"(fecha del reporte + 1 mes). Probablemente se exportó de SAP con la fecha de "
        f"vencimiento incorrecta. Se incluyeron igual en el reporte:{detalle}"
    ]


def formatear_monto(monto: Decimal, moneda: str) -> str:
    """Formato para mostrar: 'USD 1.472,20' (miles con punto, decimales con coma)."""
    texto = f"{abs(monto):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    signo = "-" if monto < 0 else ""
    return f"{moneda} {signo}{texto}"


# ---------------------------------------------------------------------------
# Reglas de negocio: armado de las 4 hojas
# ---------------------------------------------------------------------------

def _filtrar_moneda(clientes: list[Cliente], moneda: str) -> list[Cliente]:
    """Copia de los clientes con solo los documentos de esa moneda.
    Los clientes que quedan sin documentos no se incluyen."""
    resultado = []
    for c in clientes:
        docs = [d for d in c.documentos if d.moneda == moneda]
        if docs:
            resultado.append(Cliente(c.codigo, c.nombre, c.total_cabecera, c.fila, docs))
    return resultado


def _con_documentos(clientes: list[Cliente]) -> list[Cliente]:
    return [c for c in clientes if c.documentos]


def armar_hojas(
    agencia: list[Cliente], intercompany: list[Cliente], ffww: list[Cliente]
) -> tuple[list[Hoja], list[str]]:
    """Aplica las reglas de negocio y devuelve (hojas en orden, avisos)."""
    avisos = []

    # UYU fuera de lugar: en Intercompany/FFWW, o en un cliente que no sea FERSPO.
    for c in agencia:
        if CLIENTE_UYU_ESPERADO not in c.nombre.upper():
            for d in c.documentos:
                if d.moneda == "UYU":
                    avisos.append(
                        f"Agencia Marítima: {c.nombre} tiene un documento en UYU "
                        f"({d.referencia or 'sin referencia'}, fila {d.fila}). "
                        "Se incluyó en la hoja UYU; revisalo."
                    )
    for etiqueta, clientes in (("Intercompany", intercompany), ("FFWW", ffww)):
        for c in clientes:
            for d in c.documentos:
                if d.moneda != "USD":
                    avisos.append(
                        f"{etiqueta}: {c.nombre} tiene un documento en {d.moneda} "
                        f"({d.referencia or 'sin referencia'}, fila {d.fila}). "
                        "Quedó sumado junto con los USD en la hoja; revisalo."
                    )

    hojas = [
        Hoja(HOJA_AGENCIA_USD, "USD", _filtrar_moneda(agencia, "USD")),
        Hoja(HOJA_INTERCOMPANY, "USD", _con_documentos(intercompany)),
        Hoja(HOJA_AGENCIA_UYU, "UYU", _filtrar_moneda(agencia, "UYU")),
        Hoja(HOJA_FFWW, "USD", _con_documentos(ffww), con_comentarios=True),
    ]
    return hojas, avisos


# ---------------------------------------------------------------------------
# Cobranzas a reclamar y análisis
# ---------------------------------------------------------------------------

def dias_atraso(doc: Documento, fecha_reporte: date) -> int:
    """Días desde el vencimiento hasta la fecha del reporte (negativo = todavía no venció).
    Si el documento no tiene vencimiento se usa la fecha de contabilización."""
    referencia = doc.fecha_venc or doc.fecha_contab
    if referencia is None:
        return 0
    return (fecha_reporte - referencia.date()).days


def tramo_antiguedad(dias: int) -> str:
    if dias < 0:
        return "Por vencer"
    for tramo, hasta in (("0 - 30", 30), ("31 - 60", 60), ("61 - 90", 90), ("91 - 120", 120)):
        if dias <= hasta:
            return tramo
    return "121+"


def a_reclamar(hoja: Hoja, fecha_reporte: date) -> list[Cliente]:
    """
    Clientes con documentos a cobrar (saldo positivo) con más de
    DIAS_PARA_RECLAMAR días de vencidos. Pagos y notas de crédito no se incluyen.
    Clientes de mayor a menor monto a reclamar; documentos del más viejo al más nuevo.
    """
    resultado = []
    for c in hoja.clientes:
        docs = [
            d for d in c.documentos
            if d.monto > 0 and dias_atraso(d, fecha_reporte) > DIAS_PARA_RECLAMAR
        ]
        if docs:
            docs.sort(key=lambda d: dias_atraso(d, fecha_reporte), reverse=True)
            resultado.append(Cliente(c.codigo, c.nombre, c.total_cabecera, c.fila, docs))
    resultado.sort(key=lambda c: c.subtotal, reverse=True)
    return resultado


def indicadores(hoja: Hoja, fecha_reporte: date) -> Indicadores:
    docs = [d for c in hoja.clientes for d in c.documentos]
    a_cobrar = [d for d in docs if d.monto > 0]
    monto_a_cobrar = sum((d.monto for d in a_cobrar), Decimal("0"))
    atraso = None
    if monto_a_cobrar > 0:
        ponderado = sum(max(dias_atraso(d, fecha_reporte), 0) * d.monto for d in a_cobrar)
        atraso = round(ponderado / monto_a_cobrar)
    vencimientos = [d.fecha_venc for d in docs if d.fecha_venc]
    return Indicadores(
        nombre=hoja.nombre,
        moneda=hoja.moneda,
        clientes=hoja.cantidad_clientes,
        documentos=len(docs),
        total=hoja.total,
        vencido=sum((c.subtotal for c in a_reclamar(hoja, fecha_reporte)), Decimal("0")),
        atraso_promedio=atraso,
        vencimiento_mas_antiguo=min(vencimientos) if vencimientos else None,
    )


def antiguedad(hoja: Hoja, fecha_reporte: date) -> dict[str, Decimal]:
    """Saldo neto de la hoja repartido por tramo de antigüedad (suma = total de la hoja)."""
    tramos = {t: Decimal("0") for t in TRAMOS}
    for c in hoja.clientes:
        for d in c.documentos:
            tramos[tramo_antiguedad(dias_atraso(d, fecha_reporte))] += d.monto
    return tramos


def top_clientes(hojas: list[Hoja], cantidad: int = 10) -> list[tuple[Cliente, str]]:
    """Los clientes con mayor saldo entre las hojas dadas: [(cliente, nombre_hoja)]."""
    todos = [(c, h.nombre) for h in hojas for c in h.clientes]
    todos.sort(key=lambda par: par[0].subtotal, reverse=True)
    return todos[:cantidad]


# ---------------------------------------------------------------------------
# Escritura del Excel de salida
# ---------------------------------------------------------------------------

def _escribir(ws, fila, columna, valor, formato=None, negrita=False, relleno=False,
              cursiva=False, tamano=TAMANO_FUENTE, color=None):
    """Escribe una celda con la fuente del reporte y el formato indicado."""
    c = ws.cell(row=fila, column=columna, value=valor)
    c.font = Font(name=FUENTE, size=tamano, bold=negrita, italic=cursiva, color=color)
    if formato:
        c.number_format = formato
    if relleno:
        c.fill = RELLENO_ENCABEZADO
    return c


def _encabezados(ws, fila, titulos) -> None:
    for i, titulo in enumerate(titulos, start=1):
        c = _escribir(ws, fila, i, titulo, negrita=True, relleno=True)
        c.alignment = Alignment(horizontal="left", wrap_text=True)


def _aplicar_fuente(ws, max_col) -> None:
    """Fuente del reporte en todas las celdas usadas (las ya estilizadas no se tocan)."""
    fuente = Font(name=FUENTE, size=TAMANO_FUENTE)
    for fila_celdas in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max_col):
        for c in fila_celdas:
            if c.font.name != FUENTE:
                c.font = fuente


def _escribir_bloques(ws, fila, clientes, moneda, ultima_col, fecha: date) -> list[int]:
    """
    Escribe un bloque por cliente desde `fila`: nombre + subtotal, fila vacía,
    un documento por fila y borde doble al final.
    Los días vencido se calculan a la fecha del reporte (vacío si todavía no venció).
    Devuelve las filas de los subtotales.
    """
    formato_monto = FORMATO_MONTO[moneda]
    filas_subtotal = []
    for cliente in clientes:
        # Fila del cliente: nombre + subtotal (fórmula) en negrita.
        primera = fila + 2
        ultima = primera + len(cliente.documentos) - 1
        _escribir(ws, fila, 1, cliente.nombre)
        _escribir(ws, fila, COLUMNA_MONTO, f"=SUM({LETRA_MONTO}{primera}:{LETRA_MONTO}{ultima})",
                  formato_monto, negrita=True)
        filas_subtotal.append(fila)

        # (fila + 1 queda vacía) y luego un documento por fila.
        for n, doc in enumerate(cliente.documentos):
            f = primera + n
            dias = dias_atraso(doc, fecha)
            _escribir(ws, f, COLUMNA_CONTAB, doc.fecha_contab, FORMATO_FECHA)
            _escribir(ws, f, COLUMNA_VENC, doc.fecha_venc, FORMATO_FECHA)
            # Rojo solo si hay que reclamarlo: saldos a favor del cliente no se marcan.
            a_reclamar_doc = dias > DIAS_PARA_RECLAMAR and doc.monto > 0
            _escribir(ws, f, COLUMNA_DIAS, dias if dias > 0 else None, "0",
                      color=COLOR_VENCIDO if a_reclamar_doc else None)
            _escribir(ws, f, COLUMNA_REFERENCIA, doc.referencia)
            _escribir(ws, f, COLUMNA_MONTO, float(doc.monto), formato_monto)

        # Borde inferior doble en el último documento del cliente.
        for columna in range(1, ultima_col + 1):
            ws.cell(row=ultima, column=columna).border = BORDE_DOBLE

        fila = ultima + 1
    return filas_subtotal


def _formula_suma(filas) -> str | int:
    return "=" + "+".join(f"{LETRA_MONTO}{f}" for f in filas) if filas else 0


def _escribir_hoja(ws, hoja: Hoja, fecha: date) -> None:
    encabezados = ENCABEZADOS + (ENCABEZADOS_EXTRA_FFWW if hoja.con_comentarios else [])
    ultima_col = len(encabezados)

    for i, titulo in enumerate(encabezados, start=1):
        c = _escribir(ws, 1, i, titulo, negrita=True, relleno=True)
        c.alignment = Alignment(horizontal="left")

    filas_subtotal = _escribir_bloques(ws, 2, hoja.clientes, hoja.moneda, ultima_col, fecha)

    # Fila de TOTAL (justo después del último documento): suma de los subtotales.
    fila_total = ws.max_row + 1
    _escribir(ws, fila_total, COLUMNA_MONTO, _formula_suma(filas_subtotal),
              FORMATO_TOTAL[hoja.moneda], negrita=True, relleno=True)

    _aplicar_fuente(ws, ultima_col)
    for letra, ancho in ANCHOS.items():
        if ord(letra) - ord("A") < ultima_col:
            ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "A2"  # los títulos quedan visibles al bajar


def _escribir_hoja_reclamar(ws, hojas: list[Hoja], fecha: date) -> None:
    """Documentos con más de DIAS_PARA_RECLAMAR días de vencidos, una sección por hoja."""
    ultima_col = len(ENCABEZADOS)
    _encabezados(ws, 1, ENCABEZADOS)
    _escribir(
        ws, 2, 1,
        f"Documentos con más de {DIAS_PARA_RECLAMAR} días de vencidos al {fecha:%d/%m/%Y}. "
        "Solo saldos a cobrar (no incluye pagos ni notas de crédito). "
        "Clientes ordenados de mayor a menor monto a reclamar.",
        cursiva=True,
    )

    fila = 4
    for hoja in hojas:
        # Título de la sección (nombre de la hoja de origen).
        for columna in range(1, ultima_col + 1):
            _escribir(ws, fila, columna, hoja.nombre if columna == 1 else None,
                      negrita=True, relleno=True, tamano=12)
        fila += 1

        clientes = a_reclamar(hoja, fecha)
        if not clientes:
            _escribir(ws, fila, 1, "Sin documentos para reclamar.", cursiva=True)
            fila += 3
            continue

        filas_subtotal = _escribir_bloques(ws, fila, clientes, hoja.moneda, ultima_col, fecha)
        fila = ws.max_row + 1
        _escribir(ws, fila, 1, f"Total a reclamar {hoja.nombre}", negrita=True, relleno=True)
        _escribir(ws, fila, COLUMNA_MONTO, _formula_suma(filas_subtotal),
                  FORMATO_TOTAL[hoja.moneda], negrita=True, relleno=True)
        fila += 3

    _aplicar_fuente(ws, ultima_col)
    for letra, ancho in ANCHOS.items():
        if ord(letra) - ord("A") < ultima_col:
            ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "A2"


def _grafico_barras(titulo, tipo="col", apilado=False) -> BarChart:
    g = BarChart()
    g.type = tipo
    g.title = titulo
    g.title.overlay = False
    g.style = 10
    g.width, g.height = 20, 9
    if apilado:
        g.grouping = "stacked"
        g.overlap = 100
    # Sin esto, las versiones nuevas de Excel ocultan los ejes.
    g.x_axis.delete = False
    g.y_axis.delete = False
    g.y_axis.numFmt = "#,##0"
    g.y_axis.majorGridlines = None
    return g


def _escribir_hoja_analisis(ws, hojas: list[Hoja], fecha: date) -> None:
    """Indicadores, antigüedad de saldos y principales clientes, con gráficos."""
    hojas_usd = [h for h in hojas if h.moneda == "USD"]
    hojas_uyu = [h for h in hojas if h.moneda == "UYU"]
    total_usd = Hoja("Total USD", "USD", [c for h in hojas_usd for c in h.clientes])
    varias_usd = len(hojas_usd) > 1  # con una sola hoja USD, total y torta no aportan
    filas_tabla = hojas_usd + ([total_usd] if varias_usd else []) + hojas_uyu
    vencido = f"Vencido +{DIAS_PARA_RECLAMAR} días"

    _escribir(ws, 1, 1, f"Análisis de cuentas a cobrar al {fecha:%d/%m/%Y}", negrita=True, tamano=14)
    _escribir(
        ws, 2, 1,
        "Días vencido = días entre la fecha de vencimiento y la fecha del reporte. "
        + ("Los totales USD no incluyen la hoja UYU. " if hojas_uyu else "")
        + "Valores calculados al generar el reporte.",
        cursiva=True,
    )

    # 1) Indicadores por hoja.
    fila = 4
    _escribir(ws, fila, 1, "Indicadores por hoja", negrita=True, tamano=12)
    fila += 1
    _encabezados(ws, fila, [
        "Hoja", "Moneda", "Clientes", "Documentos", "Saldo total", vencido,
        f"% {vencido.lower()}", "Atraso promedio ponderado (días)", "Vencimiento más antiguo",
    ])
    fila += 1
    primera_ind = fila
    for h in filas_tabla:
        ind = indicadores(h, fecha)
        formato = FORMATO_TOTAL[h.moneda]
        valores = [
            (ind.nombre, None), (ind.moneda, None), (ind.clientes, "0"), (ind.documentos, "0"),
            (float(ind.total), formato), (float(ind.vencido), formato),
            (ind.porcentaje_vencido, "0.0%"), (ind.atraso_promedio, "0"),
            (ind.vencimiento_mas_antiguo, FORMATO_FECHA),
        ]
        for columna, (valor, fmt) in enumerate(valores, start=1):
            _escribir(ws, fila, columna, valor, fmt, negrita=h is total_usd)
        fila += 1

    # 2) Antigüedad de saldos.
    fila += 1
    _escribir(ws, fila, 1, "Antigüedad de saldos (según días de atraso)", negrita=True, tamano=12)
    fila += 1
    fila_enc_ant = fila
    _encabezados(ws, fila, ["Hoja"] + TRAMOS + ["Total"])
    fila += 1
    primera_ant = fila
    for h in filas_tabla:
        formato = FORMATO_TOTAL[h.moneda]
        tramos = antiguedad(h, fecha)
        _escribir(ws, fila, 1, h.nombre, negrita=h is total_usd)
        for columna, tramo in enumerate(TRAMOS, start=2):
            _escribir(ws, fila, columna, float(tramos[tramo]), formato, negrita=h is total_usd)
        _escribir(ws, fila, len(TRAMOS) + 2, float(h.total), formato, negrita=True)
        fila += 1

    # 3) Principales clientes USD.
    fila += 1
    top = top_clientes(hojas_usd, 10)
    _escribir(ws, fila, 1, f"Top {len(top)} clientes por saldo (USD)", negrita=True, tamano=12)
    fila += 1
    fila_enc_top = fila
    _encabezados(ws, fila, [
        "Cliente", "Hoja", "Saldo total", vencido, "% del total USD",
        "Atraso promedio ponderado (días)",
    ])
    fila += 1
    for cliente, nombre_hoja in top:
        ind = indicadores(Hoja(cliente.nombre, "USD", [cliente]), fecha)
        participacion = float(cliente.subtotal / total_usd.total) if total_usd.total else None
        _escribir(ws, fila, 1, cliente.nombre)
        _escribir(ws, fila, 2, nombre_hoja)
        _escribir(ws, fila, 3, float(cliente.subtotal), FORMATO_TOTAL["USD"])
        _escribir(ws, fila, 4, float(ind.vencido), FORMATO_TOTAL["USD"])
        _escribir(ws, fila, 5, participacion, "0.0%")
        _escribir(ws, fila, 6, ind.atraso_promedio, "0")
        fila += 1
    ultima_top = fila - 1
    if total_usd.total > 0:
        top5 = sum((c.subtotal for c, _ in top[:5]), Decimal("0"))
        _escribir(ws, fila + 1, 1, "Concentración: los 5 clientes con mayor saldo representan",
                  cursiva=True)
        _escribir(ws, fila + 1, 5, float(top5 / total_usd.total), "0.0%", negrita=True)

    # Gráficos (a la derecha de las tablas). Solo USD, para no mezclar monedas.
    if hojas_usd:
        g = _grafico_barras(
            "Antigüedad de saldos USD" + (" por hoja" if varias_usd else ""), apilado=True
        )
        g.add_data(Reference(ws, min_col=1, max_col=len(TRAMOS) + 1, min_row=primera_ant,
                             max_row=primera_ant + len(hojas_usd) - 1),
                   from_rows=True, titles_from_data=True)
        g.set_categories(Reference(ws, min_col=2, max_col=len(TRAMOS) + 1,
                                   min_row=fila_enc_ant, max_row=fila_enc_ant))
        if varias_usd:
            g.legend.position = "b"
            g.legend.overlay = False
        else:
            g.varyColors = False
            g.legend = None
        ws.add_chart(g, "K4")

    if top:
        g = _grafico_barras(f"Top {len(top)} clientes por saldo (USD)", tipo="bar")
        g.add_data(Reference(ws, min_col=3, min_row=fila_enc_top, max_row=ultima_top),
                   titles_from_data=True)
        g.set_categories(Reference(ws, min_col=1, min_row=fila_enc_top + 1, max_row=ultima_top))
        g.x_axis.scaling.orientation = "maxMin"  # el de mayor saldo arriba
        g.y_axis.crosses = "max"  # con el orden invertido, deja los montos abajo
        g.varyColors = False
        g.legend = None
        ws.add_chart(g, "K23")

    # Torta de composición del saldo USD por hoja (solo si todos los saldos son positivos).
    if varias_usd and all(h.total > 0 for h in hojas_usd):
        torta = PieChart()
        torta.title = "Composición del saldo USD por hoja"
        torta.width, torta.height = 14, 9
        torta.add_data(Reference(ws, min_col=5, min_row=primera_ind - 1,
                                 max_row=primera_ind + len(hojas_usd) - 1), titles_from_data=True)
        torta.set_categories(Reference(ws, min_col=1, min_row=primera_ind,
                                       max_row=primera_ind + len(hojas_usd) - 1))
        torta.dataLabels = DataLabelList(
            showPercent=True, showVal=False, showCatName=False, showSerName=False,
            showLegendKey=False,
        )
        torta.legend.position = "b"
        torta.legend.overlay = False
        ws.add_chart(torta, "K42")

    _aplicar_fuente(ws, 9)
    ws.column_dimensions["A"].width = 46
    for letra in "BCDEFGHI":
        ws.column_dimensions[letra].width = 17
    ws.column_dimensions["B"].width = 23
    for f in (primera_ind - 1, fila_enc_ant, fila_enc_top):
        ws.row_dimensions[f].height = 30


def construir_workbook(hojas: list[Hoja], fecha: date, con_reclamar: bool = True) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    for hoja in hojas:
        _escribir_hoja(wb.create_sheet(hoja.nombre), hoja, fecha)
    if con_reclamar:
        _escribir_hoja_reclamar(wb.create_sheet(HOJA_RECLAMAR), hojas, fecha)
    _escribir_hoja_analisis(wb.create_sheet(HOJA_ANALISIS), hojas, fecha)
    return wb


def nombre_archivo(fecha: date) -> str:
    """Ej: 24_09_2026_CUENTAS_A_COBRAR.xlsx"""
    return f"{fecha:%d_%m_%Y}_CUENTAS_A_COBRAR.xlsx"


def nombre_archivo_ffww(fecha: date) -> str:
    """Ej: 24_09_2026_CUENTAS_A_COBRAR_FFWW.xlsx"""
    return f"{fecha:%d_%m_%Y}_CUENTAS_A_COBRAR_FFWW.xlsx"


def _a_bytes(wb: Workbook) -> bytes:
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


# ---------------------------------------------------------------------------
# Punto de entrada principal
# ---------------------------------------------------------------------------

def _leer_bytes(origen) -> bytes:
    """Lee el contenido completo de una ruta o de un archivo subido."""
    if hasattr(origen, "getvalue"):
        return origen.getvalue()
    if hasattr(origen, "read"):
        origen.seek(0)
        return origen.read()
    with open(origen, "rb") as f:
        return f.read()


def generar_reporte(agencia, intercompany, ffww, fecha: date) -> Resultado:
    """
    Proceso completo: lee los 3 archivos, arma las hojas y genera el .xlsx.
    Cada parámetro de archivo puede ser una ruta o un archivo subido.
    Lanza ErrorArchivo si algo impide generar el reporte.
    """
    contenidos = {
        ETIQUETA_AGENCIA: _leer_bytes(agencia),
        ETIQUETA_INTERCOMPANY: _leer_bytes(intercompany),
        ETIQUETA_FFWW: _leer_bytes(ffww),
    }

    # El mismo archivo cargado en dos lugares distintos.
    etiquetas = list(contenidos)
    for i, a in enumerate(etiquetas):
        for b in etiquetas[i + 1:]:
            if contenidos[a] == contenidos[b]:
                raise ErrorArchivo(
                    f"Cargaste el mismo archivo en «{a}» y en «{b}». "
                    "Cada lugar lleva un archivo distinto."
                )

    avisos = []
    leidos = {}
    for etiqueta, contenido in contenidos.items():
        clientes, avisos_lectura = leer_archivo(io.BytesIO(contenido), etiqueta)
        leidos[etiqueta] = clientes
        avisos += avisos_lectura
        avisos += validar_cabeceras(clientes, etiqueta)
        avisos += validar_vencimientos(clientes, etiqueta, fecha)

    hojas, avisos_reglas = armar_hojas(
        leidos[ETIQUETA_AGENCIA], leidos[ETIQUETA_INTERCOMPANY], leidos[ETIQUETA_FFWW]
    )
    avisos += avisos_reglas

    # Versión para Comercial: la hoja FFWW y el análisis solo de FFWW (sin reclamos).
    hojas_ffww = [h for h in hojas if h.nombre == HOJA_FFWW]
    return Resultado(
        contenido=_a_bytes(construir_workbook(hojas, fecha)),
        nombre_archivo=nombre_archivo(fecha),
        hojas=hojas,
        avisos=avisos,
        fecha=fecha,
        contenido_ffww=_a_bytes(construir_workbook(hojas_ffww, fecha, con_reclamar=False)),
        nombre_archivo_ffww=nombre_archivo_ffww(fecha),
    )


def revisar_nombres_archivos(nombres: dict[str, str]) -> list[str]:
    """
    Aviso (no error) si el nombre del archivo sugiere que está en el lugar
    equivocado, ej. un "FFWW.xlsx" cargado en «2. Intercompany».
    `nombres` es {etiqueta: nombre_del_archivo}.
    """
    palabras = {
        ETIQUETA_AGENCIA: ("agencia", "maritima"),
        ETIQUETA_INTERCOMPANY: ("intercompany",),
        ETIQUETA_FFWW: ("ffww",),
    }
    avisos = []
    for lugar, archivo in nombres.items():
        archivo_norm = _normalizar(archivo)
        for otro, claves in palabras.items():
            if otro != lugar and any(k in archivo_norm for k in claves):
                avisos.append(
                    f"El archivo «{archivo}» está cargado en «{lugar}», pero por el nombre "
                    f"parece ser de «{otro}». Verificá que no estén cruzados."
                )
                break
    return avisos
