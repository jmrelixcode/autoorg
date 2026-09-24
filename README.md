# Calculadora de Finanzas

Es una ventana en Python para resolver los ejercicios del curso de Finanzas. Tú ingresas los datos y ella hace todos los cálculos. Además muestra cada paso: la fórmula en palabras, la misma fórmula con los números y el resultado.

También puede rellenar la plantilla Excel de los Ejercicios I–V. La copia queda con las fórmulas funcionando.

## Requisitos

- Windows con **Python 3.10 o superior** (se probó con 3.12). Al instalar Python, marca *"Add python.exe to PATH"*.
- La librería `openpyxl`, que está en `requirements.txt`. La ventana usa `tkinter`, que ya viene con Python.

## Instalación

Abre una terminal en esta carpeta y ejecuta:

```
pip install -r requirements.txt
```

## Cómo abrirla

- Haz doble clic en **`Calculadora finanzas.bat`**, o
- ejecuta `python calculadora.py` en una terminal dentro de esta carpeta.

La carpeta puede estar en cualquier lugar: el programa busca sus archivos junto a él.

## Qué hace cada pestaña

Cada pestaña tiene un menú **Ejercicio de clase** que carga los datos de los ejercicios vistos en clase. Puedes cambiar cualquier dato y presionar **Enter** (o **F5**) para recalcular.

| Pestaña | Resuelve |
|---|---|
| **1. Presupuesto operativo (I–V)** | Casos Panadería, Proyecto informático, Obra de construcción, Taller de electromovilidad y Mantención industrial. Calcula los presupuestos de ventas, producción, costo de producción y gastos fijos. Responde las 5 preguntas: precio para la utilidad deseada, precio −10 %, materia prima +10 %, nueva máquina y punto de equilibrio inicial y final. |
| **2. Punto de equilibrio** | Unidades (Qe) y venta en pesos ($e) de equilibrio. Compara con una situación nueva, calcula el resultado con distintos volúmenes de venta, entrega el punto de indiferencia y dibuja el gráfico. |
| **3. Estado de resultados** | Estado de resultados proyectado a varios años, con costeo directo: depreciación, gastos no operacionales, intereses e impuesto. |
| **4. Análisis vertical y horizontal** | Porcentaje de cada partida sobre la base (ventas o total de activos) y su variación entre dos años, con una frase que explica cada partida. |

El desarrollo que aparece a la derecha se puede **copiar** o **guardar como .txt**.

### Cómo escribir los números

- Puedes usar el formato chileno: `1.500.000` o `0,45`. También sirve el punto decimal: `0.45`.
- Cuando un campo pide varios años, separa los valores con `;`, por ejemplo `25.000; 30.000`. Si escribes menos valores que años, se repite el último (en intereses y en otros no operacionales se completa con 0).
- Las pérdidas se escriben en negativo, por ejemplo `-180.000.000`.

## Rellenar una plantilla Excel

Esta opción está en la pestaña 1, en el recuadro **Rellenar plantilla Excel**.

1. **Plantilla:** muestra los `.xlsx` de esta carpeta. Con **Examinar…** puedes elegir otro archivo.
2. **Hoja:** se propone sola según el caso que cargaste (caso III → hoja "Ejercicio III").
3. **Rellenar y guardar…:** eliges dónde guardar la copia y, si quieres, se abre al final.

Qué hace con el Excel:

- **Nunca modifica la plantilla original.** Siempre guarda una copia con otro nombre.
- Escribe tus datos en las celdas de entrada y las marca en **amarillo**.
- Deja los datos de las preguntas (baja de precio, alza de M.P., nueva máquina, utilidad deseada) en `K28:L34`.
- Vuelve a escribir las fórmulas de cálculo, así que si cambias un dato en la hoja, todo se recalcula.
- Corrige errores de la plantilla original: precios escritos a mano en `G56`/`F56` de las hojas II a V y referencias rotas en el cuadro `K51:O78`.

Solo funciona con hojas que tienen el formato de los Ejercicios I–V ("Paso 1…" en `B2` e "Ingresos por venta" en `C52`). Si eliges otra hoja, la ventana te avisa y no escribe nada.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `calculadora.py` | La ventana. Solo pide los datos y muestra los resultados. |
| `calculos.py` | Todas las fórmulas y los datos de los ejercicios de clase. Sirve sin la ventana. |
| `plantilla_excel.py` | Rellena la plantilla Excel. |
| `Calculadora finanzas.bat` | Abre la ventana con doble clic. |
| `requirements.txt` | Dependencias de Python. |

## Supuestos de cálculo

- **Costeo directo:** el costo de venta es unidades vendidas × (M.P. + M.O. + CIF). Los gastos de venta y administración, incluida la depreciación, son gastos fijos.
- **Producción** = ventas + inventario final deseado − inventario inicial. El inventario inicial de cada mes es el inventario final del mes anterior.
- **Precio para la utilidad deseada** = (costo variable total + gastos fijos + utilidad) ÷ unidades.
- **Nueva máquina:** los gastos fijos suben en (depreciación nueva − depreciación actual) × 3 meses.
- **Punto de equilibrio:** Qe = CF ÷ (P − CV) y $e = CF ÷ (1 − CV/P).
- **Impuesto:** solo se calcula si la utilidad antes de impuestos es positiva. Con pérdida no se paga impuesto.

## Problemas comunes

- **"No module named openpyxl":** falta instalar las dependencias (`pip install -r requirements.txt`).
- **"No se pudo guardar: el archivo está abierto en Excel":** cierra el archivo en Excel e intenta otra vez.
- **El `.bat` no abre nada:** Python no está en el PATH. Prueba con `py calculadora.py` o reinstala Python con la opción *Add to PATH* marcada.
