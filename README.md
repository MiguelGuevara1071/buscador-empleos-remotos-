# 🚀 Buscador de Empleos Remotos por Telegram

Bot automatizado que busca ofertas remotas en español para:
- 🐍 **Python**
- 📊 **Ingeniería de Datos**
- ⌨️ **Data Entry**
- 💻 **Ingeniería de Software**

Fuentes integradas: **LinkedIn (Colombia / LATAM)**, **CompuTrabajo Colombia**, **Get on Board** y **Remotive**.

---

## ☁️ Configuración en GitHub Actions (Ejecución 24/7 Gratis)

Para que el bot corra cada hora en la nube de GitHub con tu computadora apagada:

1. Crea un repositorio en GitHub (puede ser privado o público) y sube este proyecto.
2. En tu repositorio de GitHub, entra a:
   **Settings** > **Secrets and variables** > **Actions**
3. Haz clic en **New repository secret** y añade estos dos secretos (usando los valores que tienes en tu archivo local `.env`):
   - **Nombre:** `TELEGRAM_BOT_TOKEN`
     - **Valor:** *(Tu token de @BotFather, el que está en tu .env)*
   - **Nombre:** `TELEGRAM_CHAT_ID`
     - **Valor:** *(Tu ID numérico de Telegram, el que está en tu .env)*
4. En **Settings** > **Actions** > **General** > **Workflow permissions**, asegúrate de que esté marcado:
   - ✅ **Read and write permissions** (para que guarde el historial de empleos vistos).
5. ¡Listo! El bot se ejecutará cada hora de forma automática y gratuita. Puedes probarlo cuando quieras desde la pestaña **Actions** > **Run workflow**.

---

## 💻 Ejecución Local

1. Crea tu archivo `.env` (o usa el que ya está configurado):
   ```env
   TELEGRAM_BOT_TOKEN="tu_token"
   TELEGRAM_CHAT_ID="tu_chat_id"
   ```
2. Ejecuta el script:
   ```bash
   python buscar_empleos.py
   ```
