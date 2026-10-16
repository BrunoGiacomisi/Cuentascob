# Cuentas a Cobrar – Generador semanal

Genera el Excel `DD_MM_YYYY_CUENTAS_A_COBRAR.xlsx` a partir de los 3 exports del ERP
(Agencia Marítima, Intercompany y FFWW).

## Uso

1. Doble clic en **`iniciar.bat`**. La primera vez instala lo necesario (unos minutos).
2. Se abre el navegador. Cargá cada archivo en su lugar:
   1. Agencia Marítima
   2. Intercompany
   3. FFWW
3. Elegí la fecha del reporte y apretá **Generar**.
4. Revisá el resumen y los avisos, y descargá las 2 planillas:
   - **Completa**: `DD_MM_YYYY_CUENTAS_A_COBRAR.xlsx`
   - **Solo FFWW (para Comercial)**: `DD_MM_YYYY_CUENTAS_A_COBRAR_FFWW.xlsx`, con la hoja
     FFWW y la hoja «Análisis y gráficas» calculada solo con FFWW.

Para cerrar la aplicación, cerrá la ventana negra.

## Requisitos

- Windows con Python 3.11 o superior (al instalarlo, marcar *Add Python to PATH*).

## Qué genera

Un Excel con 4 hojas: AGENCIA MARITIMA USD, INTERCOMPANY, AGENCIA MARITIMA UYU y FFWW.
Por cliente: subtotal (fórmula), una línea por documento y total general al final.
La columna **Días vencido** muestra los días desde el vencimiento hasta la fecha del
reporte (vacía si el documento todavía no venció).

Al final se agregan 2 hojas de apoyo:
- **Posibles cobranzas a reclamar**: documentos con más de 30 días de vencidos a la fecha
  del reporte, separados por hoja de origen. Solo saldos a cobrar (sin pagos ni notas de
  crédito); clientes de mayor a menor monto, con los días de atraso de cada documento.
- **Análisis y gráficas**: indicadores por hoja (saldo, % vencido +30 días, atraso promedio),
  antigüedad de saldos por tramos, top 10 clientes USD y gráficos.

Avisos que puede mostrar (no frenan el proceso):
- Documentos en UYU fuera de lo esperado (en Intercompany/FFWW, o de un cliente que no sea FERSPO).
- Clientes cuyo subtotal recalculado no coincide con el total de cabecera del ERP.
- Documentos que vencen después de la fecha del reporte + 1 mes: indica que en SAP se
  exportó con la fecha de vencimiento incorrecta (los documentos se incluyen igual).
- Saldos que vinieron como número sin moneda (se asume USD).
- Nombres de archivo que sugieren que se cargaron en el lugar equivocado.

## Archivos

- `procesador.py`: lógica (lectura, reglas y armado del Excel).
- `app.py`: pantalla (Streamlit).
- `iniciar.bat`: arranque con doble clic.
