📘 Proyecto de Consolidación y Carga de Notas en MongoDB

Este proyecto procesa archivos de Excel con las sábanas de notas descriptivas de distintos grados, limpia y transforma los datos, y finalmente los inserta en una base de datos MongoDB Atlas para su consulta centralizada.

🚀 Tecnologías utilizadas

Python 3.12

Pandas
 → procesamiento de datos

PyMongo
 → conexión a MongoDB

MongoDB Atlas
 → base de datos en la nube

Archivos Excel (.xlsx) → fuente de notas académicas


.
├── periodos/
│   ├── Sabana_de_notas_descriptiva_11.xlsx
│   ├── Sabana_de_notas_descriptiva_10.xlsx
│   └── ... (archivos de los demás grados)
├── procesarNotas_11_10.py
├── procesarNotas_1_9.py
├── main.py
├── consolidado.json
└── README.md


git clone <repo-url>
cd proyecto-notas

python -m venv venv
source venv/bin/activate   # en Linux/Mac
venv\Scripts\activate      # en Windows

pip install pandas pymongo openpyxl

python main.py


Qué hace main.py:

Procesa cada archivo de notas según el grado:

Grados 11 y 10 → función procesarNotas

Grados 9 a 1 → función procesarNotas_1_9

Limpia los datos:

Elimina filas y columnas innecesarias.

Renombra las columnas.

Convierte notas a valores numéricos.

Crea columnas faltantes con valor 0 (entero).

Añade columnas adicionales (observaciones, metas, rep_eva).

Filtra los datos por período (PERIODO 1, PERIODO 2, etc.).

Concatena todos los grados en un único DataFrame.

Exporta los resultados a un archivo consolidado.json.

Inserta los datos en MongoDB Atlas en la base test, colección studentnotes.


--------------------------------------- 

Ejemplo de Actualización desde mongodb

Filtrar:

{
  grupo: { $in: ["7. A", "7. B"] },
  periodo: "TERCER PERIODO",
  tecnologia: { $gte: 3, $lte: 4 }
}

despues cambiar texto:

{
  $set: {
    observaciones_tecnologia: "Debe mejorar sus hábitos de estudio, organización y responsabilidad frente a los compromisos académicos.\n\n- A pesar de las orientaciones y oportunidades de recuperación, no logra demostrar el nivel de desempeño esperado en el área.\n- Presenta dificultades para relacionar los contenidos teóricos con situaciones prácticas y proyectos tecnológicos.\n- Necesita fortalecer la capacidad para trabajar de manera autónoma, seguir procesos y buscar soluciones a problemas tecnológicos.",

    metas_tecnologia: "Fortalecer su participación, iniciativa y disposición frente a las actividades de aprendizaje y mejorar sus hábitos de estudio, organización y cumplimiento de los compromisos académicos.\n\n- Aprovechar las oportunidades de recuperación para superar dificultades y alcanzar los desempeños esperados.\n- Relacionar los conocimientos teóricos con su aplicación en actividades prácticas y proyectos tecnológicos.\n- Fortalecer la autonomía, el seguimiento de procesos y la búsqueda de soluciones a situaciones tecnológicas."
  }
}