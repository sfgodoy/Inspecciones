# Carga de informes de inspección a planilla Excel

App web (Streamlit) que lee un informe PDF y devuelve la planilla Excel de requerimientos.

## Contenido
- `app.py`: interfaz web.
- `casa_inspeccion_a_planilla.py`: lógica de lectura y carga.
- `perfiles/`: un archivo JSON por formato de informe.
- `plantilla/planilla_base.xlsx`: planilla base con catálogos (Recinto, Lugar, Item, Problema).
- `requirements.txt`: dependencias.

## Probar en tu PC
    pip install -r requirements.txt
    streamlit run app.py

## Publicar online (Streamlit Community Cloud, gratis)
1. Crea un repositorio en GitHub (puede ser privado) y sube esta carpeta.
2. Entra a https://share.streamlit.io con tu cuenta de GitHub.
3. "Create app": elige el repositorio, rama principal y archivo `app.py`.
4. Deploy. Obtendrás una URL para usar desde cualquier navegador.

## Agregar un formato nuevo
Copia un perfil de `perfiles/`, ajústalo (ver instrucciones en el archivo de perfiles) y
súbelo al repositorio. Aparecerá solo en el selector de formato.
