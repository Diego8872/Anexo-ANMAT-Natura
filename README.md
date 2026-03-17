# Anexo ANMAT Natura

Generador automático del Anexo de Productos ANMAT para operaciones de importación Natura y Avon.

## Archivos necesarios

| Archivo | Formato | Descripción |
|---|---|---|
| Packing List | .xlsx | PL de la operación |
| Próximas Importaciones | .xlsx | Una sola factura por operación |
| Registro ANMAT Histórico | .xlsb / .xlsx | Base de datos ANMAT |
| Registros Avon | .xlsx | Base de datos Avon |
| Fabricantes | .xls / .xlsx | Tabla de fabricantes por origen |
| Catálogo NCM | .xlsx | Posiciones arancelarias por material |

## Lógicas principales

- Una línea del PL = una línea del Anexo
- Cruce por MATERIAL CODE con ANMAT Histórico → si no se encuentra, busca en Registros Avon
- REFIL en descripción → agrega (REPUESTO) al nombre
- DIFUSOR y 3x1 generan anexos separados
- Fabricante por match de ORIGEN en planilla Fabricantes (normalización de acentos/mayúsculas)
- NCM por MATERIAL CODE en Catálogo NCM
- Salida: Excel completo + Excel sin primeras 2 cols + PDF apaisado

## Outputs por operación

Por cada grupo (PRINCIPAL, DIFUSOR, 3x1):
- `ANEXO_{GRUPO}_{INVOICE}.xlsx` — todas las columnas
- `ANEXO_{GRUPO}_{INVOICE}_SIN_MAT.xlsx` — sin MATERIAL y descripcion_factura  
- `ANEXO_{GRUPO}_{INVOICE}_SIN_MAT.pdf` — PDF apaisado

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```
