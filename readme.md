Primero se deben instalar las dependencias ejecutando los siguientes comandos en la terminal:
pip install openpyxl
pip install -U google-generativeai

Luego es necesario crear archivo con extension .env y sin nombre en la misma carpeta en la que tengamos el programa, dentro del mismo usar la plantilla proporcionada y remplazar los valores con tus propias keys:

OWM_API_KEY = "[tu_owm_api_key]"
GEMINI_API_KEY = "[tu_gemini_api_key]"

A tener en cuenta que es recomendado ejecutar el programa dentro de una carpeta en solitario ya que la creacion de los archivos .csv es automatica por tanto se crearan en la misma direccion donde se este ejecutando el programa

Una vez finalizada la preparacion se puede ejecutar la aplicacion haciendo el siguiente comando: 
python challenge_tec5.py

Una vez ejecutada la aplicacion se encontrara con un menu de inicio de sesion/registro y una vez completado el inicio, tendra acceso al menu de utilizacion del programa con las utilidades:
1. Consultar Clima Actual y Guardar en Historial
2. Ver Mi Historial Personal
3. Estadísticas Globales
4. Consejo IA: ¿Cómo Me Visto Hoy?
5. Acerca De...
6. Cerrar Sesión