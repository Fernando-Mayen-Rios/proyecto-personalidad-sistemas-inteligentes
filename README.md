# Test de personalidad: ¿se puede predecir el perfil con formularios?

Proyecto de **Sistemas Inteligentes** · UNAM · Matemáticas Aplicadas y Computación
Equipo: [Nombres del equipo] · Entrega: 18 de noviembre de 2026

## ¿De qué trata?

Recolectamos las respuestas de tres formularios (**personalidad**, **música** y **relaciones amorosas**), las guardamos en una base de datos, las limpiamos con reglas explícitas, las analizamos y entrenamos modelos para intentar predecir el perfil de personalidad de una persona (**Introvertido** o **No introvertido**).

**Resultado principal:** con estos datos (21 personas con perfil) **no hay evidencia de que el perfil se pueda predecir mejor que el azar**. El mejor modelo obtiene 0.68 de aciertos balanceados (el azar es 0.50), pero una prueba de permutaciones da un valor p de 0.23: el azar iguala ese resultado casi una de cada cuatro veces. El valor del proyecto es el proceso completo y reproducible, y haber medido con evidencia hasta dónde se puede llegar.

## Resumen del trabajo

1. **Formularios.** Diseñamos el de personalidad y respondimos los de otros equipos. Elegimos dos temas complementarios: música y relaciones amorosas.
2. **Base de datos.** SQLite local (9 tablas, 6 vistas, llaves primaria y foránea), sin datos personales. Las tablas se exportan a CSV para llevarlas a la nube.
3. **Limpieza.** Solo reglas explícitas y expresiones regulares (sin IA ni NLP). Se conserva la respuesta original y cada paso queda en la bitácora.
4. **Perfil.** El formulario no trae una pregunta que indique si alguien es introvertido, así que lo definimos con reglas sobre dos preguntas (la 6, solo o acompañado, y la 7, comportamiento con desconocidos).
5. **EDA.** Análisis univariado, bivariado y multivariado, y revisión de faltantes.
6. **Variables nuevas.** Algunas respuestas de texto se convirtieron a variables 0/1 con palabras clave.
7. **Modelado.** Cuatro modelos y un modelo base, validación "dejar uno fuera", prueba de azar y agrupamiento (K-means).
8. **Tablero.** Dos versiones, ninguna en Excel (ver más abajo).

## Datos

| Formulario | Respuestas recibidas | Usadas | Qué se excluyó |
|---|---|---|---|
| Personalidad | 24 | 23 | 1 respuesta repetida de un integrante del equipo (se conservó la más reciente) |
| Música | 23 | 23 | |
| Relaciones amorosas | 20 | 19 | 1 persona que no dio su consentimiento |

Los formularios se unieron por persona (23 personas; 19 con los tres formularios) con autorización de la profesora. Los datos finales no tienen correos, números de cuenta ni marcas de tiempo.

## Perfil (variable a predecir)

| Perfil | Regla | Personas |
|---|---|---|
| Introvertido | Prefiere estar solo y se describe reservado con desconocidos | 14 |
| No introvertido | Prefiere estar acompañado o depende, o se describe sociable | 7 |
| Por revisar | Prefiere estar solo, pero la pregunta 7 no se puede clasificar (no se fuerza una etiqueta) | 2 |

Con 3 perfiles (introvertido, ambivertido, extrovertido) quedaban 17, 5 y solo 1 persona, así que se usaron 2. Las preguntas 6 y 7 **no** se usan para predecir.

## Variables usadas en el modelo (9) y de qué formulario salen

| Formulario | Variable | Pregunta de origen |
|---|---|---|
| Personalidad | `genero_femenino` | Género |
| Personalidad | `problema_con_otros` | 9. Qué hace ante un problema que no sabe resolver |
| Personalidad | `rol_lider` | 10. Papel que toma en un equipo |
| Personalidad | `reaccion_incomodidad` | 11. Reacción ante un cambio de planes |
| Personalidad | `act_conversar` | 13. Actividades que disfruta con otras personas |
| Música | `mus_tiempo_horas` | 2. Tiempo al día escuchando música (mínimo de horas) |
| Música | `usa_spotify`, `usa_youtube` | 4. Plataformas que usa |
| Relaciones amorosas | `tiene_pareja` | ¿Tiene pareja actualmente? |

**Variables que no se usaron, y por qué**

- **Personalidad:** pasatiempo, deporte, platillo, género de películas, qué hace cuando alguien le pide ayuda y actividad de fin de semana (texto libre sin categorías claras con reglas sencillas); cómo decide y cómo organiza su tiempo (casi todas las personas respondieron lo mismo, sin variación).
- **Música:** 14 de las 16 preguntas (tipo de música, momentos, estado de ánimo, géneros, concentración, sueño, artistas...): texto libre sin categorías claras con reglas sencillas.
- **Relaciones amorosas:** edad de la primera relación (33 % de vacíos y valores de 0) y diferencia de edad máxima (57 % de vacíos); área, carrera y semestre (iguales para todas las personas); y las demás preguntas abiertas (texto libre, y solo 19 personas respondieron ese formulario).

## Resultados de los modelos (21 personas, "dejar uno fuera")

| Modelo | Aciertos | Aciertos balanceados | F1 | AUC |
|---|---|---|---|---|
| Árbol de decisión (prof. 2) | 0.67 | 0.68 | 0.59 | 0.54 |
| Base (siempre "Introvertido") | 0.67 | 0.50 | 0.00 | 0.50 |
| Bosque aleatorio | 0.52 | 0.43 | 0.17 | 0.57 |
| k vecinos (k = 3) | 0.48 | 0.43 | 0.27 | 0.45 |
| Regresión logística | 0.29 | 0.21 | 0.00 | 0.15 |

El azar en aciertos balanceados es 0.50. Valor p del mejor modelo: **0.23** (100 repeticiones). K-means no encontró los perfiles (índice de Rand ajustado −0.06). Los resultados del bosque aleatorio y del agrupamiento pueden variar un poco según la versión de scikit-learn; los de este repositorio son los obtenidos en Google Colab.

## Tablero analítico

- **Tablero principal en Looker Studio** (en la nube, sin Excel): [pegar aquí el link]
- **Tablero web complementario** (en la nube): [pegar aquí el link de Netlify]. También está el archivo `tablero_personalidad.html` en este repositorio, que se abre con doble clic en el navegador.

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
| `tablero_personalidad.html` | Tablero web complementario |
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

## Datos y privacidad

- Los datos analíticos **no contienen** correos, números de cuenta ni marcas de tiempo; cada persona tiene un código anónimo (U001, U002...).
- Se excluyó la respuesta de una persona que no dio consentimiento.
- La unión de los formularios por persona fue autorizada por la profesora.

## Decisiones y limitaciones

- La limpieza usa únicamente **reglas explícitas y expresiones regulares** (sin IA ni NLP), y cada paso queda en la bitácora.
- No se imputaron datos fuera del modelado; los valores sospechosos (edad 0, 15 horas de escucha) se conservaron y se marcaron.
- Una persona (U022) repite textualmente respuestas de otras en dos formularios; se conservó y se dejó señalada (excluirla no cambia las conclusiones).
- Son muy pocas personas: los resultados son **exploratorios**.
- El perfil se definió con dos preguntas del formulario; no es un test psicométrico validado.

## Uso de herramientas de apoyo

Se utilizó Claude (Anthropic) como herramienta de apoyo para escribir y revisar código, explicar conceptos y redactar borradores de documentos y diapositivas. Las decisiones del proyecto (tema, preguntas, reglas del perfil, exclusiones y tratamiento de casos dudosos) fueron revisadas y tomadas por el equipo, y el código se ejecutó y probó en Google Colab. La limpieza de datos no utiliza IA ni NLP: solo reglas explícitas y expresiones regulares.

## Versiones con las que se probó

Python 3.12.3 · pandas 3.0.6 · numpy 2.5.4 · openpyxl 3.1.5 · matplotlib 3.11.2 · scikit-learn 1.9.1
