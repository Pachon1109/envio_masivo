import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
import time

st.title("📧 Envío masivo de certificados")

# -------------------------
# CONFIGURACION GMAIL
# -------------------------

correo_remitente = st.text_input("Correo Gmail")
password = st.text_input("Contraseña de aplicación", type="password")

# -------------------------
# PLANTILLA DE CORREO
# -------------------------

st.subheader("✉️ Contenido del correo")

asunto = st.text_input(
    "Asunto",
    "Entrega de certificado"
)

mensaje = st.text_area(
    "Mensaje del correo",
    """Hola {nombre},

Adjunto encontrarás tu certificado.

Cordialmente"""
)

st.caption("Puedes usar las variables: {nombre}, {codigo}, {email}")

# -------------------------
# CARGA DE ARCHIVOS
# -------------------------

excel_file = st.file_uploader("Cargar Excel de destinatarios", type=["xlsx"])
pdf_files = st.file_uploader("Cargar certificados PDF", type=["pdf"], accept_multiple_files=True)

# -------------------------
# BOTON ENVIAR
# -------------------------

if st.button("Enviar certificados"):

    if not excel_file or not pdf_files:
        st.warning("Debe cargar el Excel y los certificados")
        st.stop()

    df = pd.read_excel(excel_file)

    # Crear diccionario de PDFs
    pdf_dict = {}
    for pdf in pdf_files:
        nombre_pdf = pdf.name.replace(".pdf", "")
        pdf_dict[nombre_pdf] = pdf

    enviados = []
    fallidos = []

    progress = st.progress(0)

    smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    smtp.login(correo_remitente, password)

    total = len(df)

    for i, row in df.iterrows():

        identificacion = str(row["codigo"])
        nombre = row["nombre"]
        email = row["email"]

        if identificacion not in pdf_dict:
            fallidos.append([codigo, nombre, email, "PDF no encontrado"])
            continue

        try:

            # Reemplazar variables en el mensaje
            texto = mensaje.format(
                nombre=nombre,
                codigo=codigo,
                email=email
            )

            msg = EmailMessage()
            msg["Subject"] = asunto
            msg["From"] = correo_remitente
            msg["To"] = email

            msg.set_content(texto)

            pdf = pdf_dict[codigo]

            msg.add_attachment(
                pdf.read(),
                maintype="application",
                subtype="pdf",
                filename=pdf.name
            )

            smtp.send_message(msg)

            enviados.append([codigo, nombre, email, "Enviado"])

        except Exception as e:
            fallidos.append([codigo, nombre, email, str(e)])
            st.error(f"Error enviando a {email}: {e}")

        progress.progress((i + 1) / total)
        time.sleep(1)

    smtp.quit()

    # -------------------------
    # REPORTE
    # -------------------------

    df_enviados = pd.DataFrame(enviados, columns=["id", "nombre", "email", "estado"])
    df_fallidos = pd.DataFrame(fallidos, columns=["id", "nombre", "email", "error"])

    with pd.ExcelWriter("reporte_envio.xlsx") as writer:
        df_enviados.to_excel(writer, sheet_name="enviados", index=False)
        df_fallidos.to_excel(writer, sheet_name="fallidos", index=False)

    st.success("Proceso finalizado")

    st.write("Enviados:", len(enviados))
    st.write("Fallidos:", len(fallidos))

    with open("reporte_envio.xlsx", "rb") as f:
        st.download_button(
            "Descargar reporte",
            f,
            file_name="reporte_envio.xlsx"

        )
