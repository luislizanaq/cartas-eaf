# %%
import pandas as pd
from datetime import datetime, timedelta
from docxtpl import DocxTemplate

doc=DocxTemplate("DE a XXXX por EAF XXX-XXX.docx")
df_reuc=pd.read_excel("datos_reuc.xlsx")
df_cartas=pd.read_excel("list_cartas.xlsx")

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



# %%
