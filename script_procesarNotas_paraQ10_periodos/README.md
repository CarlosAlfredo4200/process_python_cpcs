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
  grupo: { $in: ["10. A", "10. B"] },
  periodo: "TERCER PERIODO",
  tecnologia: { $gte: 3, $lte: 4 }         // cambiar el área
}

{
  grupo: { $in: ["11. A", "11. B"] },
  periodo: "TERCER PERIODO",
  $expr: {
    $and: [
      {
        $gte: [
          { $convert: { input: "$quimica", to: "double", onError: null, onNull: null } },
          3
        ]
      },
      {
        $lte: [
          { $convert: { input: "$quimica", to: "double", onError: null, onNull: null } },
          3.95
        ]
      }
    ]
  }
}

despues cambiar texto:

{
  $set: {
    observaciones_quimica: "El estudiante demuestra una comprensión básica de los compuestos oxigenados de la química orgánica, reconociendo algunas de sus características y grupos funcionales. Sin embargo, necesita fortalecer la identificación, diferenciación y aplicación de las reglas de nomenclatura de estos compuestos.",
    
    metas_quimica: "Fortalecer la identificación y clasificación de los principales compuestos oxigenados (alcoholes, éteres, aldehídos, cetonas, ácidos carboxílicos y ésteres), aplicando correctamente las reglas de nomenclatura IUPAC y relacionando sus estructuras con sus respectivos grupos funcionales.",

    
  }
}

{
  $set: {
    observaciones_idioma_extranjero: "...",

    metas_idioma_extranjero: "...",
    
    rep_eva_idioma_extranjero:"...",
}
}