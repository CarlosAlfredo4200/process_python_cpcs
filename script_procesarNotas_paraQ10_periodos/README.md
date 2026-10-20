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
  tecnologia: { $gte: 3, $lte: 4 }         // cambiar el área
}

{
  grupo: { $in: ["6. A", "6. B"] },
  periodo: "SEGUNDO PERIODO",
  $expr: {
    $and: [
      {
        $gte: [
          { $convert: { input: "$idioma_extranjero", to: "double", onError: null, onNull: null } },
          3
        ]
      },
      {
        $lte: [
          { $convert: { input: "$idioma_extranjero", to: "double", onError: null, onNull: null } },
          3.95
        ]
      }
    ]
  }
}

despues cambiar texto:

{
  $set: {
    observaciones_idioma_extranjero: "El estudiante logró reconocer y trabajar algunas de las estructuras gramaticales abordadas durante el período, especialmente en las actividades relacionadas con Present Perfect y Past Simple. Sin embargo, la falta recurrente de presentación de trabajos y actividades limitó las oportunidades para evidenciar de manera completa y constante sus aprendizajes.\nEl estudiante tuvo la oportunidad de fortalecer sus habilidades de inglés mediante el trabajo autónomo en Rosetta, herramienta destinada al desarrollo progresivo de sus competencias lingüísticas. No obstante, el incumplimiento de las unidades y actividades requeridas en la plataforma afectó la continuidad de su proceso y el cumplimiento de los objetivos establecidos para el período.\nEl estudiante contó con espacios para participar en actividades de lectura, escritura, expresión oral y en el proyecto de Misiones, en los cuales podía demostrar responsabilidad, respeto y compromiso. Sin embargo, algunas actitudes y comportamientos durante las clases dificultaron el desarrollo adecuado de las actividades y el aprovechamiento de los espacios de aprendizaje, por lo que requiere fortalecer su disposición y compromiso frente al trabajo académico.",
    metas_idioma_extranjero: "Fortalecer la responsabilidad académica, presentando oportunamente los trabajos, actividades y proyectos asignados.\nCumplir de manera constante con las unidades establecidas en Rosetta, siguiendo el cronograma y organizando espacios de trabajo autónomo.\nMejorar su actitud y comportamiento durante las clases, demostrando respeto por docentes y compañeros, disposición para aprender y cumplimiento de las normas establecidas.\nParticipar activamente en las actividades de inglés, especialmente en aquellas que involucran lectura, escritura, producción oral y proyectos, aprovechando estos espacios para demostrar sus aprendizajes.",
    rep_eva_idioma_extranjero: "Durante el período se realizó seguimiento al proceso académico y actitudinal del estudiante mediante el acompañamiento en las actividades desarrolladas en clase, recordatorios de los compromisos y del cronograma de trabajo en Rosetta, así como espacios destinados al desarrollo y entrega de las actividades propuestas. Se brindaron orientaciones para favorecer la comprensión de los temas y el cumplimiento de los trabajos asignados.\nA pesar del acompañamiento realizado, persistieron algunas dificultades relacionadas con la presentación oportuna de trabajos, el cumplimiento de las actividades requeridas en Rosetta y la actitud asumida durante determinados espacios de clase. Para favorecer su avance, se recomienda continuar fortaleciendo la responsabilidad, el trabajo autónomo, el seguimiento de instrucciones y una actitud respetuosa y comprometida frente al proceso de aprendizaje."
  }
}

{
  $set: {
    observaciones_idioma_extranjero: "...",

    metas_idioma_extranjero: "...",
    
    rep_eva_idioma_extranjero:"...",
}
}