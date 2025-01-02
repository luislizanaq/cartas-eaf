# %%
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from docxtpl import DocxTemplate

st.title("📄 Cartas por EAF V.1.0")

df_cartas_cuerpo = pd.DataFrame(
    [
        {"EAF": "554-2024", "Nombre": "Nombre del EAF", "Fecha falla": "23-11-2024", "Hora falla": "01:03", "Coordinado": "Transelec S.A."},
        
    ]
)

df_cartas = st.data_editor(df_cartas_cuerpo, num_rows="dynamic")

st.button("Reset", type="primary")
if st.button("Crear Cartas"):
    doc=DocxTemplate("DE a XXXX por EAF XXX-XXX.docx")
    df_reuc=pd.read_excel("datos_reuc.xlsx")



    dic_remp={
    "January":"enero", 
    "February":"febrero", 
    "March":"marzo", 
    "April":"abril",
    "May":"mayo",
    "June":"junio",
    "July":"julio",
    "August":"agosto",
    "September":"septiembre",
    "October":"octubre",
    "November":"noviembre",
    "December":"diciembre"
    }


    for i in df_cartas.index:
        empresa = df_cartas["Coordinado"][i]
        fecha_plazo=datetime.today()+timedelta(days=7)
        fecha_plazo=fecha_plazo.strftime("%d de %B del %Y")
        fecha_hoy=datetime.today().strftime("%d de %B del %Y")
        fecha_falla=df_cartas["Fecha falla"][i].strftime("%d de %B del %Y")

        for palabra,remmplazo in dic_remp.items():
            fecha_falla=fecha_falla.replace(palabra,remmplazo)

        for palabra,remmplazo in dic_remp.items():
            fecha_plazo=fecha_plazo.replace(palabra,remmplazo)

        for palabra,remmplazo in dic_remp.items():
            fecha_hoy=fecha_hoy.replace(palabra,remmplazo)    


  



        dic={
        "empresa":empresa, 
        "enc_titular":df_reuc["Encargado Titular"].loc[df_reuc["Razón Social"]==empresa].to_list()[0], 
        "enc_suplente":df_reuc["Encargado Suplente"].loc[df_reuc["Razón Social"]==empresa].to_list()[0], 
        "hora_falla":df_cartas["Hora falla"][i], 
        "fecha_plazo":fecha_plazo,
        "fecha_hoy":fecha_hoy,
        "fecha_falla":fecha_falla,
        "nom_eaf":df_cartas["Nombre"][i]
        }
        doc.render(dic)
        doc.save(f"0"+str(i)+" DE a "+empresa+" por EAF "+df_cartas["EAF"][i]+".docx")
        with open("0"+str(i)+" DE a "+empresa+" por EAF "+df_cartas["EAF"][i]+".docx", "rb") as docx:
            btn = st.download_button(
                label="Descargar DOCX",
                data=docx,
                file_name="carta.docx",
                mime="image/png",
    
            )

else:
    st.write("Goodbye")




# %%
