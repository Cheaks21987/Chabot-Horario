import os
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from langchain.prompts import PromptTemplate
from langchain.chat_models import ChatOpenAI
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
import pytz
from datetime import datetime, timedelta
import unicodedata
import re

# Cargar variables de entorno
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")

# Función para eliminar acentos
def eliminar_acentos(texto):
    return ''.join(
        char for char in unicodedata.normalize('NFD', texto)
        if unicodedata.category(char) != 'Mn'
    )

# Limpiar el DataFrame
def limpiar_datos(df):
    df = df.drop_duplicates()
    df = df.dropna(how='all')
    df = df.fillna({'DOCENTE': 'desconocido', 'UBICACIÓN': 'no especificado', 'INSTALACIÓN': 'no especificado'})
    columnas_texto = ['DOCENTE', 'CURSO', 'UBICACIÓN', 'INSTALACIÓN', 'DÍA']
    for columna in columnas_texto:
        if columna in df.columns and pd.api.types.is_string_dtype(df[columna]):
            df[columna] = df[columna].str.lower().str.strip().str.title()
            df[columna] = df[columna].apply(eliminar_acentos)
    return df

# Hora y fecha actual en Perú
def obtener_fecha_actual():
    peru_tz = pytz.timezone('America/Lima')
    fecha_actual = datetime.now(peru_tz)
    dia_semana = fecha_actual.strftime("%A").capitalize()
    return fecha_actual.strftime("%Y-%m-%d"), dia_semana

# Obtener el día solicitado
def calcular_dia(dias_offset):
    peru_tz = pytz.timezone('America/Lima')
    fecha_actual = datetime.now(peru_tz)
    dia_calculado = fecha_actual + timedelta(days=dias_offset)
    return dia_calculado.strftime("%A").capitalize()

# Interpretar ubicación en "INSTALACIÓN"
def interpretar_instalacion(valor):
    if 'a' in valor.lower() or 'b' in valor.lower():
        return "Campus Tacna y Arica"
    return "Campus Av. Parra"

# Plantilla de prompt
prompt_template = """
Tienes acceso a los siguientes datos de cursos:
{data}

La hora actual en Perú es {current_time}.

Con base en estos datos y el historial de la conversación, responde con precisión a la siguiente pregunta:
Pregunta: {question}

Historial de la conversación:
{history}
"""

prompt = PromptTemplate(
    input_variables=["data", "question", "history", "current_time"],
    template=prompt_template,
)

# Inicializar el modelo de chat
llm = ChatOpenAI(model="gpt-3.5-turbo", openai_api_key=openai_api_key)

# Crear memoria de conversación
memory = ConversationBufferMemory(return_messages=True)

# Limitar el tamaño del DataFrame
def limitar_datos(df, max_filas=200, max_columnas=12):
    return df.iloc[:max_filas, :max_columnas]

# Función mejorada para extraer información compleja
def extraer_informacion_compleja(df, pregunta):
    pregunta_limpia = eliminar_acentos(pregunta.lower())
    dias_offset = {"hoy": 0, "mañana": 1, "pasado mañana": 2, "ayer": -1}
    
    # Identificar si se pregunta por un día
    for clave, offset in dias_offset.items():
        if clave in pregunta_limpia:
            dia_solicitado = calcular_dia(offset)
            break
    else:
        # Intentar extraer directamente el día de la semana
        dias_semana = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        dia_solicitado = next((dia.capitalize() for dia in dias_semana if dia in pregunta_limpia), None)

    if dia_solicitado:
        df_filtrado = df[df["DÍA"].str.lower() == dia_solicitado.lower()]
        if df_filtrado.empty:
            return f"No hay cursos programados para {dia_solicitado}."
        else:
            cursos = df_filtrado["CURSO"].unique()
            return f"Los cursos programados para {dia_solicitado} son: {', '.join(cursos)}."

    # Identificar si se pregunta por un docente
    if "dicta" in pregunta_limpia or "docente" in pregunta_limpia:
        match = re.search(r'docente\s+([a-záéíóúñ]+)', pregunta_limpia) or re.search(r'dicta\s+([a-záéíóúñ]+)', pregunta_limpia)
        if match:
            docente = eliminar_acentos(match.group(1)).capitalize()
            df_filtrado = df[df["DOCENTE"].str.contains(docente, case=False, na=False)]
            if df_filtrado.empty:
                return f"No se encontraron cursos dictados por el docente {docente}."
            else:
                # Identificar si se pregunta específicamente por los días
                if "qué días" in pregunta_limpia or "que dias" in pregunta_limpia:
                    dias = df_filtrado["DÍA"].unique()
                    return f"El docente {docente} dicta cursos los días: {', '.join(dias)}."
                else:
                    cursos = df_filtrado["CURSO"].unique()
                    return f"Los cursos dictados por {docente} son: {', '.join(cursos)}."

    # Mensaje predeterminado si no se puede determinar la pregunta
    return "Lo siento, no entendí tu pregunta. Por favor, sé más específico."

# Función para responder preguntas del usuario
def responder_pregunta_excel(df, pregunta, texto_conversacion, boton_enviar):
    texto_conversacion.insert(tk.END, "Chatbot: Cargando...\n")
    boton_enviar.config(state=tk.DISABLED)
    texto_conversacion.update()

    # Intentar obtener respuesta directamente
    respuesta = extraer_informacion_compleja(df, pregunta)
    if "no hay" not in respuesta.lower() and "no se encontraron" not in respuesta.lower() and "no entendí" not in respuesta.lower():
        texto_conversacion.delete("end-2l", "end-1l")
        boton_enviar.config(state=tk.NORMAL)
        return respuesta

    # Si no se encuentra respuesta, utilizar el modelo LLM
    df_limitado = limitar_datos(df)
    datos_excel = df_limitado.to_string(index=False)
    current_time, _ = obtener_fecha_actual()
    historia = memory.load_memory_variables({})["history"]

    mensaje = prompt.format(data=datos_excel, question=pregunta, history=historia, current_time=current_time)
    respuesta = llm.invoke(mensaje).content
    memory.save_context({"input": pregunta}, {"output": respuesta})

    texto_conversacion.delete("end-2l", "end-1l")
    boton_enviar.config(state=tk.NORMAL)
    return respuesta

# Interfaz gráfica con Tkinter
def iniciar_interfaz():
    ventana = tk.Tk()
    ventana.title("Chatbot Horario")
    ventana.geometry("500x700")
    ventana.configure(bg="#f0f4f7")

    texto_conversacion = tk.Text(ventana, height=20, width=60, wrap=tk.WORD, bg="#e1effa", fg="#0d1f44", font=("Calibri", 12))
    texto_conversacion.pack(pady=10, expand=True, fill="both")
    texto_conversacion.insert(tk.END, "Chatbot: Hola, ¿sobre qué desea saber de su horario?\n")

    entrada_pregunta = tk.Entry(ventana, width=60, font=("Arial", 12))
    entrada_pregunta.pack(pady=10)

    def enviar_pregunta():
        pregunta = entrada_pregunta.get()
        if pregunta.lower() == "salir":
            ventana.quit()
        else:
            texto_conversacion.insert(tk.END, f"Tú: {pregunta}\n")
            respuesta = responder_pregunta_excel(df, pregunta, texto_conversacion, boton_enviar)
            texto_conversacion.insert(tk.END, f"Chatbot: {respuesta}\n")
            entrada_pregunta.delete(0, tk.END)

    boton_enviar = tk.Button(ventana, text="Enviar", command=enviar_pregunta, font=("Arial", 12), bg="#0066cc", fg="white")
    boton_enviar.pack(pady=5)
    entrada_pregunta.bind("<Return>", lambda event: enviar_pregunta())

    calendario = Calendar(ventana, selectmode='day')
    calendario.pack(pady=10)
    
    ventana.mainloop()

# Cargar archivo Excel y ejecutar
df = pd.read_excel('horarios.xlsx')
df = limpiar_datos(df)
df['INSTALACIÓN'] = df['INSTALACIÓN'].apply(interpretar_instalacion)
iniciar_interfaz()
