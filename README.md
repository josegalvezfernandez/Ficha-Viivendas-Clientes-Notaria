# Ficha Viviendas - IFEBA

Este repositorio contiene:
- Script de generación de fichas Excel.
- Plantilla Excel de ficha.
- Archivo origen con todas las viviendas.
- Salida generada (Ifeba 1.xlsx).

Resumen del mapeo (alto nivel):
- Fuente principal: pestaña "SITUACION COMERCIAL".
- Datos de mejoras: pestaña "MEJORAS" (por Nº Orden en columna B).
  - F15 = Q
  - G15 = O + H
  - H15 = R
  - H20 = R

Ejecución local:
- Edita rutas si cambian.
- Ejecuta: python ifeba_generate.py
