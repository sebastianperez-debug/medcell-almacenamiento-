# 📦 Informe de Almacenamiento — Streamlit

Réplica en Streamlit de un reporte de Power BI de ocupación de bodega
(ubicaciones ocupadas/disponibles por pasillo, nivel y tipo de bodega).

## Estructura del proyecto

```
├── app.py                  # App principal de Streamlit
├── requirements.txt        # Dependencias
├── data/
│   └── Almacenamiento_2026.xlsx   # Datos base (hoja "UBICACIONES")
└── README.md
```

## 1. Probarlo en tu computador

```bash
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá en `http://localhost:8501`.

## 2. Subirlo a GitHub

```bash
git init
git add .
git commit -m "Dashboard de almacenamiento"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

> ⚠️ El archivo `data/Almacenamiento_2026.xlsx` pesa ~1.7 MB, así que se
> puede subir sin problema a GitHub (el límite normal es 100 MB por
> archivo). Si más adelante tus datos crecen mucho, considera usar
> Git LFS o cargar el Excel desde la app (ya incluí un `file_uploader`
> en la barra lateral para eso).

## 3. Desplegarlo en Streamlit Community Cloud

1. Ve a [share.streamlit.io](https://share.streamlit.io) e inicia sesión con GitHub.
2. Clic en **"New app"**.
3. Selecciona tu repositorio, la rama `main` y el archivo `app.py`.
4. Clic en **"Deploy"**. En un par de minutos tendrás una URL pública tipo
   `https://tu-app.streamlit.app`.

Cada vez que hagas `git push` a `main`, la app se actualiza sola.

## 4. Actualizar los datos cada mes

Tienes dos opciones:

- **Reemplazar el archivo en el repo**: sobrescribe
  `data/Almacenamiento_2026.xlsx` con el nuevo export y haz `git push`.
- **Subir el archivo directamente en la app**: usa el cargador de
  archivos de la barra lateral (útil si no quieres tocar GitHub cada vez).

## Notas sobre los datos

- La app lee la hoja **UBICACIONES** del Excel (no lee archivos `.pbix`
  directamente — Power BI no permite eso).
- El filtro **"Código artículo"** del reporte original no está incluido
  porque esa información vive en otra tabla (maestro de artículos) que
  no venía en este Excel. Compárteme esa tabla y la agrego.
- El campo **"Última fecha: fecha proceso"** se muestra como la fecha
  de hoy porque el Excel no trae una columna de fecha de proceso;
  si tu fuente de datos sí la tiene, dime en qué hoja/columna está y la
  conecto.
