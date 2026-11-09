# 🏗️ App de Bitácora Diaria de Obra (Streamlit + Supabase)

Carpeta lista para desplegar en **Streamlit Cloud** o correr en tu computador.

## 📂 Archivos incluidos en esta carpeta
- `app.py`: Código completo de la aplicación (incluye clonación de personal del día anterior, compresión automática de fotos y exportación a Word `.docx`).
- `requirements.txt`: Librerías necesarias para que Streamlit Cloud instale todo automáticamente.
- `schema_supabase.sql`: Código SQL para crear las 4 tablas en Supabase.
- `.streamlit/secrets.toml.example`: Plantilla para colocar tus llaves de Supabase.
- `.gitignore`: Para evitar subir tus llaves secretas a GitHub.

## 🚀 Pasos rápidos para ponerla a funcionar (5 minutos)
1. **En Supabase:**
   - Abre el **SQL Editor**, pega el contenido de `schema_supabase.sql` y dale **Run**.
   - Ve a **Storage**, crea un nuevo Bucket llamado `fotos_bitacora` y márcalo como **Public Bucket** (además agrega una política en Policies para permitir `INSERT` y `SELECT` a usuarios anónimos/`anon`).
2. **En GitHub / Streamlit Cloud:**
   - Sube los archivos de esta carpeta a un repositorio de GitHub.
   - Conéctalo en Streamlit Cloud y en **Advanced Settings -> Secrets** pega tu `SUPABASE_URL` y `SUPABASE_KEY`.
