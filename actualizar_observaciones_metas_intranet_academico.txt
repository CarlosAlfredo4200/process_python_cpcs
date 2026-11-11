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
  grupo: { $in: ["7. A", "6. B"] },
  periodo: "TERCER PERIODO",
  tecnologia: { $gte: 3, $lte: 4 }         // cambiar el área
}

{
  grupo: { $in: ["6. B"] },
  periodo: "TERCER PERIODO",
  $expr: {
    $and: [
      {
        $gte: [
          { $convert: { input: "$ciencias_naturales", to: "double", onError: null, onNull: null } },
          3
        ]
      },
      {
        $lte: [
          { $convert: { input: "$ciencias_naturales", to: "double", onError: null, onNull: null } },
          3.95
        ]
      }
    ]
  }
}

despues cambiar texto:

{
  $set: {
    observaciones_ciencias_naturales: "El estudiante obtuvo un desempeño básico en el área de Matemáticas. Necesita fortalecer la comprensión y aplicación de los conceptos relacionados con los números racionales. 2. Requiere afianzar la identificación, representación, comparación y ubicación de números racionales, así como el desarrollo de operaciones con fracciones y números decimales. 3. Es importante fortalecer sus hábitos de estudio y repaso en casa, practicar los procedimientos matemáticos y prestar mayor atención a las explicaciones e instrucciones dadas en clase.",
    
    metas_ciencias_naturales: "Reconocer, representar, comparar y ubicar correctamente los números racionales en diferentes situaciones matemáticas. 2. Fortalecer el manejo de las operaciones básicas con fracciones y números decimales, aplicando correctamente los procedimientos y propiedades. 3. Desarrollar habilidades para resolver situaciones problema que involucren números racionales, mediante la práctica constante y el repaso de los contenidos en casa, buscando superar el promedio obtenido.",

    rep_eva_ciencias_naturales: "Durante el tercer periodo, Isaac Cardona obtuvo un promedio de 3.9 en el área de Matemáticas, alcanzando los desempeños básicos propuestos en el tema de números racionales. Necesita continuar fortaleciendo la comprensión de los conceptos, el manejo de las operaciones con fracciones y números decimales, y la resolución de situaciones problema. Se recomienda dedicar mayor tiempo al repaso en casa, practicar los ejercicios desarrollados en clase y fortalecer su autonomía en los procedimientos matemáticos. Cuenta con las capacidades para alcanzar mejores resultados si mantiene la constancia, la atención y el compromiso con su aprendizaje."
  }
}

{
  $set: {
    observaciones_idioma_extranjero: "...",

    metas_idioma_extranjero: "...",
    
    rep_eva_idioma_extranjero:"...",
}
}