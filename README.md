# 📈 ALYNK - Monitor de Mercado en Vivo

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![xlwings](https://img.shields.io/badge/xlwings-Excel_Live-green?style=for-the-badge)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

Un tablero de control (Dashboard) de escritorio diseñado para mesas de dinero. Visualiza cotizaciones financieras (Dólar MEP, CCL y Arbitrajes) en tiempo real conectándose directamente a la memoria de una instancia activa de Microsoft Excel.

## 🚀 Características Principales

* **Conexión en Memoria:** Utiliza `xlwings` para leer datos de un Excel abierto sin necesidad de guardar el archivo en disco. Ideal para feeds de datos DDE/RTD que actualizan constantemente.
* **Interfaz Dark Mode:** Diseño visual optimizado para entornos de trading con alto contraste (Negro/Gris).
* **Formato Inteligente:** Distingue automáticamente entre:
    * 🟢 **Precios:** Formato moneda (`$ 1,200.00`) en color verde.
    * 🔵 **Spreads/Canjes:** Formato porcentaje (`0.54%`) en color celeste.
* **Parametrizable:** Permite pasar la ruta del archivo Excel como argumento de sistema para mayor flexibilidad.
* **Branding:** Soporte para logotipo corporativo dinámico.

## 📋 Requisitos Previos

* **Sistema Operativo:** Windows 10/11 (Requerido para la interacción COM con Excel).
* **Microsoft Excel:** Instalado y activado.
* **Python:** Versión 3.x instalada.

## 🛠️ Instalación

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/Julian-Rando/Monitor-Alynk.git
    cd Monitor-Alynk
    ```

2.  **Instalar dependencias:**
    Ejecuta el siguiente comando para instalar las librerías necesarias (`xlwings` para conectar con Excel y `Pillow` para procesar el logo).
    ```bash
    pip install xlwings pillow
    ```

## ⚙️ Configuración del Mapa de Datos

El archivo `ejecutor.py` contiene un diccionario llamado `MAPA_DATOS` que vincula las celdas del Excel con el Dashboard. Puedes editar este diccionario para cambiar qué celdas lee el programa.

**Estructura del `MAPA_DATOS`:**
* **Clave:** Nombre del instrumento (ej: "AL30").
* **Valor:** Diccionario con coordenadas y tipo de dato.
    * `"P"`: **Precio** (Se muestra como `$ 1,200.00`).
    * `"S"`: **Spread** (Se multiplica por 100 y muestra `%`).

```python
# Ejemplo de configuración en ejecutor.py
MAPA_DATOS = {
    "AL30": {
        "C":   ("B3", "P"), # Compra (Celda B3, Tipo Precio)
        "V":   ("E3", "P"), # Venta
        "CCL": ("H3", "P")  # Cable
    },
    "CANJE COMPRA": {
        "C":   ("B7", "S"), # Spread (Celda B7, Tipo Spread/Porcentaje)
        "V":   ("E7", "S"),
        "CCL": ("H7", "P")
    }
}
```

## 📂 Estructura de Archivos

Asegúrate de mantener los siguientes archivos en la misma carpeta:

```text
/alynk-monitor
│
├── ejecutor.py        # Script principal de Python
├── alynk logo.png     # Logo de la empresa (Requerido para el pie de página)
└── INICIAR.bat        # Script lanzador para usuarios finales

```
## ▶️ Uso y Ejecución

⚠️ **IMPORTANTE:** Antes de ejecutar cualquier comando, asegúrate de que **el archivo Excel esté ABIERTO** en tu escritorio.

Para facilitar el uso diario, ejecuta el archivo `INICIAR.bat`. Asegura de que las rutas de `SCRIPT` y `EXCEL` sean correctas.



## 🔍 Solución de Problemas Frecuentes

| Error / Síntoma | Causa Probable | Solución |
| :--- | :--- | :--- |
| **❌ NO SE DETECTA EXCEL...** | El archivo Excel está cerrado o el nombre no coincide. | Abre el Excel antes de iniciar el monitor. Verifica que el nombre pasado por parámetro sea correcto. |
| **Error: ModuleNotFoundError** | Faltan librerías. | Ejecuta `pip install xlwings pillow`. |
| **Ventana cierra al instante** | Error de sintaxis o librerías. | Ejecuta el script desde la terminal (`cmd`) para ver el mensaje de error real. |
| **Precios aparecen como %** | Configuración incorrecta en `MAPA_DATOS`. | Cambia el tipo de dato de `"S"` a `"P"` en la celda correspondiente dentro del script. |

## 📄 Licencia

Este proyecto es de uso interno y propietario de **ALYNK**.
