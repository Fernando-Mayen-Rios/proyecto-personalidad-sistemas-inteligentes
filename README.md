# Test de personalidad: ¿se puede predecir el perfil con formularios?

Proyecto de **Sistemas Inteligentes** · UNAM · Matemáticas Aplicadas y Computación
Equipo: 
-Chávez Ramírez Iker David
-Mayen Rios Fernando
-Rangel Sosa Oscar Yardel
-Sanchez Lopez Miguel Angel
-Torres Montes Sharon Danae
 · Entrega: 18 de noviembre de 2026

## ¿De qué trata?

Recolectamos las respuestas de tres formularios (**personalidad**, **música** y **relaciones amorosas**), las guardamos en una base de datos, las limpiamos con reglas explícitas, las analizamos y entrenamos modelos para intentar predecir el perfil de personalidad de una persona (**Introvertido** o **No introvertido**).

**Resultado principal:** con estos datos (21 personas con perfil) **no hay evidencia de que el perfil se pueda predecir mejor que el azar**. El mejor modelo obtiene 0.68 de aciertos balanceados (el azar es 0.50) con un valor p de 0.23. El valor del proyecto es el proceso completo y reproducible, y haber medido con evidencia hasta dónde se puede llegar.

## Contenido del repositorio

| Archivo | Qué es |
|---|---|
| `01_Datos_y_Etiqueta.ipynb` | Carga los datos y define los perfiles con reglas |
| `02_EDA.ipynb` | Análisis exploratorio (univariado, bivariado, multivariado, faltantes) |
| `03_Variables_Texto.ipynb` | Crea variables a partir de respuestas de texto con expresiones regulares |
| `04_Modelado.ipynb` | Entrena y compara modelos, prueba de azar y agrupamiento |
| `05_Base_de_Datos.ipynb` | Crea la base de datos SQLite y las consultas |
| `Datos_Unidos.xlsx` | Los tres formularios unidos por persona (anónimo) |
| `Datos_Limpios.xlsx` | Cada formulario limpio por separado |
| `Datos_Con_Perfil.xlsx`, `Datos_Modelado.xlsx`, `Resultados_Modelos.xlsx` | Salidas de los cuadernos 01, 03 y 04 |
| `Base_Datos_Personalidad.db`, `consultas_ejemplo.sql`, `exportes_csv.zip` | Base de datos, consultas de ejemplo y tablas en CSV para la nube |
| `Bitacora_Limpieza.xlsx`, `Reporte_Auditoria.xlsx` | Registro de cada paso de limpieza y auditoría de los datos |
| `codigo_limpieza.py`, `codigo_union.py` | Scripts de limpieza y unión (**evidencia**: parten de los datos originales, que no se incluyen) |
| `requirements.txt` | Librerías necesarias |

## Cómo reproducir los resultados

Cada cuaderno indica qué archivo necesita y, si no lo encuentra, abre una ventana para subirlo. Como el repositorio ya incluye todos los archivos intermedios, **cualquier cuaderno se puede ejecutar por separado**.

| Cuaderno | Archivo(s) de entrada | Archivo(s) que genera |
|---|---|---|
| 01 | `Datos_Unidos.xlsx` | `Datos_Con_Perfil.xlsx` |
| 02 | `Datos_Con_Perfil.xlsx` | (solo análisis y gráficas) |
| 03 | `Datos_Con_Perfil.xlsx` | `Datos_Modelado.xlsx` |
| 04 | `Datos_Modelado.xlsx` | `Resultados_Modelos.xlsx` |
| 05 | `Datos_Con_Perfil.xlsx`, `Datos_Modelado.xlsx`, `Datos_Unidos.xlsx`, `Bitacora_Limpieza.xlsx`, `Resultados_Modelos.xlsx` | `Base_Datos_Personalidad.db`, `consultas_ejemplo.sql`, `exportes_csv.zip` |

### Opción A: Google Colab (la más sencilla, no instala nada)

1. Descarga el repositorio (botón verde **Code → Download ZIP**) y descomprímelo.
2. Entra a <https://colab.research.google.com> con una cuenta de Google.
3. **Archivo → Subir cuaderno** y elige `01_Datos_y_Etiqueta.ipynb`.
4. **Entorno de ejecución → Ejecutar todo**. Cuando se abra la ventana para subir archivos, elige `Datos_Unidos.xlsx`.
5. Repite con los cuadernos 02, 03, 04 y 05, en ese orden. El 05 pide cinco archivos a la vez (puedes elegirlos juntos).

Colab ya trae las librerías necesarias. Los archivos subidos se borran al cerrar la sesión, por lo que hay que subirlos de nuevo en cada sesión.

### Opción B: en tu computadora

Se probó con Python 3.12.3; se recomienda Python 3.10 o superior.

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac o Linux:
source venv/bin/activate

pip install -r requirements.txt
pip install notebook
jupyter notebook
```

Coloca todos los archivos del repositorio en la misma carpeta, abre cada cuaderno en Jupyter y ejecútalos en orden (**Kernel → Restart & Run All**). Los cuadernos buscan los archivos en esa misma carpeta.

### Scripts de limpieza y unión (solo con los datos originales)

Los datos originales contienen correos y números de cuenta, por lo que **no están en el repositorio**. Estos scripts se incluyen para mostrar cómo se hizo la limpieza. Quien tenga el libro original (`RespuestasFormsOriginales.xlsx`) puede ejecutarlos así:

```bash
python codigo_limpieza.py --entrada RespuestasFormsOriginales.xlsx --salida .
python codigo_union.py --entrada RespuestasFormsOriginales.xlsx --salida .
```

## Tablero analítico

- Tablero web en la nube (no Excel): https://claude.ai/artifact/32nDuLq2Q2yYYz4YH49g7K 
- Versión alternativa en Looker Studio: [pegar aquí el link, si se creó]

## Datos y privacidad

- Los datos analíticos **no contienen** correos, números de cuenta ni marcas de tiempo; cada persona tiene un código anónimo (U001, U002...).
- Se excluyó la respuesta de una persona que no dio consentimiento.
- La unión de los formularios por persona fue autorizada por la profesora.

## Decisiones y limitaciones

- La limpieza usa únicamente **reglas explícitas y expresiones regulares** (sin IA ni NLP), y cada paso queda en la bitácora.
- No se imputaron datos fuera del modelado; los valores sospechosos (edad 0, 15 horas de escucha) se conservaron y se marcaron.
- Son muy pocas personas: los resultados son **exploratorios**.
- El perfil se definió con dos preguntas del formulario; no es un test psicométrico validado.
- Los resultados del bosque aleatorio pueden variar un poco según la versión de scikit-learn.

## Versiones con las que se probó

Python 3.12.3 · pandas 3.0.6 · numpy 2.5.4 · openpyxl 3.1.5 · matplotlib 3.11.2 · scikit-learn 1.9.1
