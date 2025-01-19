# -*- coding: utf-8 -*-
"""
Created on Mon Jan 22 15:42:00 2024
@author: Millary Antunez
"""

'''
        ETAPA 0: Preparar e importar librerías
        
'''

# Importar librerías

import pandas as pd
from datetime import datetime
import calendar
import locale

# Fecha de la base del último corte

fecha = '28062024'

# Extraer día, mes y año

dia = fecha[0:2]
mes = fecha[2:4]
año = fecha[4:]

# Variables para tablero

fecha_corte = datetime.strptime(año+'-'+mes+'-'+dia, '%Y-%m-%d').strftime('%d/%m/%Y')
fecha_corte_id = año+mes+dia

# Formato Español

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

mes_corte = calendar.month_name[int(mes)].capitalize()
num_mes_corte = int(mes)
num_mes_corte_r = str(num_mes_corte)

# Pasar a minúscula

mes_corte_mi = mes_corte.lower()

# Ruta del directorio

ruta = r'C:\Users\Chrystel\Desktop\7. PRUEBAS NEXUS_AIRSHP'
#ruta = r'C:\Users\ANALISTAUP22\Desktop\7. PRUEBAS NEXUS_AIRSHP'
#ruta = 'B:\\OneDrive - Ministerio de Educación\\unidad_B\\2024\\4. Herramientas de Seguimiento\\3. Plazas CAS IAP\\Nexus'
# Ruta de PEAS_2023_programación

ruta_anexo = ruta+'\Input'

# Ruta de cas fecha de corte

ruta_corte = ruta+'\Input\Diten'

# Ruta Power Bi

ruta_bi = ruta+'\Input\input_pbi'

# Ruta de archivos de salida

ruta_output = ruta+'\Output\\1. Plantillas'

# Ruta de reporte

ruta_rep = ruta+'\Output\\2. Reporte'

'''
        ETAPA 1: Base del último corte

'''

# Importar base del último corte

# Mantener el formato str para código de cargo

#column_types_1 = {'CODCARGO': str}

column_types_1 = {'CODCARGO': str}

corte = pd.read_excel(ruta_corte+'/cas_'+fecha+'.xlsx',sheet_name='Sheet1', header=0)

# Pasar a mayúsculas

corte.columns = corte.columns.str.upper()

########corte['FECINICIO'] = pd.to_datetime(corte['FECINICIO'], format='%d/%m/%Y')
#corte['fecinicio'] = corte['fecinicio'].dt.strftime('%d/%m/%y')

############corte['FECINICIO'] = pd.to_datetime(corte['FECINICIO'], format='%d/%m/%Y')
#corte['fectermino'] = corte['fectermino'].dt.strftime('%d/%m/%y')

# Eliminar espacios de nombres de variables

corte.columns = corte.columns.str.strip()

# Renombrar variable

corte.rename(columns={'SUBTIPOPLA':'SUBTIPOPLAZA'},inplace=True)
corte.rename(columns={'CODUGE':'COD_UGEL_DITEN'},inplace=True)
corte.rename(columns={'NOMBREOOII':'UGEL_DITEN'},inplace=True)
# Almacenar para observaciones

obs = corte.copy()

# Verificar la situación laboral: SITLAB

s_sitlab = corte['SITLAB'].value_counts()
print(s_sitlab)


# Modificar valores

corte.loc[(corte['SITLAB'] == 'V'),'SITLAB'] = 'VACANTE'
corte.loc[(corte['SITLAB'] == 'X'),'SITLAB'] = 'CONTRATADO'
corte.loc[(corte['SITLAB'] == 'C'),'SITLAB'] = 'CONTRATADO'

# Variables que deben contener vacíos caso no se haya realizado la contratación de la plaza
corte.loc[corte['SITLAB'] == 'VACANTE', ['FECINICIO', 'FECTERMINO', 'NUMDOCUM']] = ''

# Para los casos de duplicados de código de plaza, se seleccionarán aquellos que tengan el estado de la plaza "ACTIV" y la situación laboral "CONTRATADO".

#Eliminamos codigos plaza con estado de la plaza diferente a "ACTIV"

corte = corte.query('ESTPLAZA=="ACTIV"')

# Ordenar por código de plaza y situación laboral

corte = corte.sort_values(by=['SITLAB'])

# Verificar duplicados

#corte['duplicado'] = corte.duplicated(subset=['CODPLAZA'])

# Mantener el primer valor de los casos duplicados de código de plaza, estos coindicen con la situación laboral "CONTRATADO", en caso no haya ningún contratado se toma cualquier valor

#corte_unico = corte.drop_duplicates(subset=['CODPLAZA'], keep='first')

# Seleccionar variables de interés

corte_i = corte[['CODPLAZA','SITLAB','FECINICIO','FECTERMINO','NUMDOCUM','APELLIPAT','APELLIMAT','NOMBRES','COD_UGEL_DITEN','UGEL_DITEN']]

'''
        ETAPA 2: Reporte de duplicados 

'''

# Seleccionar los casos duplicados en 'codplaza' con la situación laboral "CONTRATADO"

duplicados = corte[corte.duplicated(subset='CODPLAZA', keep=False)].copy()

'''
        ETAPA 3: PEAS sin SITLAB(contratado/vacante) definido

'''

# Las PEAS que requieren ejecutar el actualizador en el Nexus no presentan información de SITLAB (contratado/vacante)

#a) Importar base PEAS programadas

# Mantener el formato str para código de cargo

peas_prog = pd.read_excel(ruta_anexo+'/PEA_2024_programación - RM 060 2024 - v2.0.xlsx',sheet_name='Anexo1', dtype={'COD_CARGO': str})


peas_prog['Fecha inicio programado'] = pd.to_datetime(peas_prog['Fecha inicio programado'], format='%Y/%m/%d')
peas_prog['Fecha inicio programado'] = peas_prog['Fecha inicio programado'].dt.strftime('%d/%m/%y')
peas_prog.rename(columns={'Fecha inicio programado':'Fecha inicio programado Norma Técnica'},inplace=True)

peas_prog['Fecha fin programado'] = pd.to_datetime(peas_prog['Fecha fin programado'], format='%Y/%m/%d')
peas_prog['Fecha fin programado'] = peas_prog['Fecha fin programado'].dt.strftime('%d/%m/%y')
peas_prog.rename(columns={'Fecha fin programado':'Fecha fin programado Norma Técnica'},inplace=True)

#b) Modificar Huánuco a macro región Oriente, a pedido de Territorial
peas_prog.loc[(peas_prog['Región']=='HUANUCO'),'MACRO_REGION'] = 'ORIENTE'


# Importar código de intervenciones

column_types_3 = {'COD_INT': str}
cod_2024 = pd.read_excel(ruta_anexo+'/Intervenciones.xlsx',sheet_name='intervenciones', header=0, nrows=55)#EDITAR PARA AGREGAR INTERVENCIONES

# Combinar bases
peas_prog_cod=pd.merge(peas_prog, cod_2024, on =['COD_INT'], how ='inner')

#c) Combinar base de PEAS programadas y base del último corte

#Eliminamos variables duplicadas
peas_prog_cod = peas_prog_cod.drop('Intervención_x', axis=1)
peas_prog_cod = peas_prog_cod.drop('Intervención - nombre corto_x', axis=1)
peas_prog_cod = peas_prog_cod.drop('ANEXO_NT', axis=1)

# Renombrar variable
peas_prog_cod.rename(columns={'Código plaza NEXUS':'CODPLAZA'},inplace=True)
peas_prog_cod.rename(columns={'Intervención - nombre corto_y':'Intervención - nombre corto'},inplace=True)
peas_prog_cod.rename(columns={'Intervención_y':'Intervención'},inplace=True)
peas_prog_cod.rename(columns={'ANEXO NT 2024':'ANEXO_NT'},inplace=True)

# Crear variables
peas_prog_cod['NC1'] = 64.19
peas_prog_cod['NC2'] = 50
peas_prog_cod['Sueldo programado'] = peas_prog_cod['Sueldo inicial'] + peas_prog_cod['NC1'] + peas_prog_cod['NC2']


# Comprobar código de plaza
peas_prog_cod_corte = pd.merge(peas_prog_cod, corte_i, on =['CODPLAZA'], how ='outer',indicator=True)

# Filtrar el DataFrame resultante para quedarse con las filas que están en el DataFrame izquierdo y en ambos DataFrames
peas_prog_cod_corte= peas_prog_cod_corte[peas_prog_cod_corte['_merge'].isin(['left_only', 'both'])]


#Renombrar variables y dar formato fecha

#peas_prog_cod_corte['FECINICIO'] = pd.to_datetime(peas_prog_cod_corte['FECINICIO'], format='%Y-%m-%d %H:%M:%S')
#peas_prog_cod_corte['FECINICIO'] = peas_prog_cod_corte['FECINICIO'].dt.strftime('%d/%m/%y')

peas_prog_cod_corte.rename(columns={'FECINICIO':'Fecha inicio NEXUS'},inplace=True)


#peas_prog_cod_corte['FECTERMINO'] = pd.to_datetime(peas_prog_cod_corte['FECTERMINO'], format='%Y/%m/%d')
#peas_prog_cod_corte['FECTERMINO'] = peas_prog_cod_corte['FECTERMINO'].dt.strftime('%d/%m/%y')
peas_prog_cod_corte.rename(columns={'FECTERMINO':'Fecha termino NEXUS'},inplace=True)

#Formato de fecha del NEXUS


#d) Código de plaza que no aparecen en la base del último corte
sin_codplaza = peas_prog_cod_corte.loc[peas_prog_cod_corte['_merge'] == 'left_only']

# Eliminar espacios de nombres de variables
sin_codplaza.columns = sin_codplaza.columns.str.strip()

# Variables de interés

sin_codplaza['Marco Normativo'] = "RM N° 060 - 2024 MINEDU" #crea la columna Marco normativo en el dataframe sin_codplaza  y asignarle a todas las filas el mismo valor, que en este caso es "RM N° 060 - 2024 MINEDU"
sin_codplaza.loc[sin_codplaza['COD_INT'].isin([53, 54]), 'Marco Normativo'] = None #no considera SAE y CUNAS

sin_codplaza_i = sin_codplaza[['Dependencia','Dirección General','Dirección Línea','COD_REGION','Región','MACRO_REGION','COD_PLIEGO','Pliego','COD_UE',
                               'Unidad Ejecutora','COD_UGEL','DRE/UGEL','Código de local','Código modular','Anexo','Nombre de la IE','CODPLAZA','COD_INT',
                               'Intervención','Intervención - nombre corto','COD_CARGO','Cargo','Sueldo programado','Sueldo inicial',
                               'PEAS programadas','Mes inicio programado','N° meses programado','Fecha inicio programado Norma Técnica','Fecha fin programado Norma Técnica','Marco Normativo']]

#e) Exportar

sin_codplaza_i.to_excel(ruta_output+'/nocod_'+fecha+'_plantilla.xlsx', sheet_name=mes_corte , index= False)

'''
        ETAPA 4: Reporte 

'''

#a) Unir bases para agregar información presupuestal

reporte = peas_prog_cod_corte

# Generar variables binarias para representar la situación laboral

reporte['Contratado'] = 0

reporte.loc[(reporte['SITLAB']=='CONTRATADO'),'Contratado'] = 1

reporte['Vacante'] = 0

reporte.loc[(reporte['SITLAB']=='VACANTE'),'Vacante'] = 1

#b) Generar variables binarias para representar el registro en NEXUS: base del último corte

reporte['SITLAB'] = reporte['SITLAB'].fillna('NO REGISTRADO') #Especifica la columna 'SITLAB' dentro del DataFrame reporte, .fillna() es un método que se utiliza para rellenar los valores faltantes (NaN) con el valor especificado. 'NO REGISTRADO': Es el valor que se utilizará para reemplazar los valores NaN en la columna 

reporte['No_Registrado'] = 0

reporte.loc[(reporte['SITLAB']=='NO REGISTRADO'),'No_Registrado'] = 1

#c) Verificar la situación laboral: SITLAB

t_sitlab = reporte['SITLAB'].value_counts()
print(t_sitlab)

# Generar variables binarias para representar las PEAS programadas

reporte['Programadas'] = 1

# Corrección del sueldo

#reporte['Sueldo programado total'] = reporte['Sueldo programado'] + reporte['Incremento DS N° 311-2022-EF']

# Evaluar variables binarias

reporte.rename(columns={'Vacante':'Vacante_inicio'},inplace=True)

reporte['Vacante'] = reporte['Programadas'] - reporte['Contratado']

# Modificar los no registrados por vacante

reporte.loc[reporte['Contratado'] == 1, 'SITLAB'] = 'CONTRATADO'
reporte.loc[reporte['Contratado'] == 0, 'SITLAB'] = 'VACANTE'

# Crear variable
reporte['Marco Normativo'] = "RM N° 060 - 2024 MINEDU"

#d) Agregrar variable SEC_EJEC

# Importar base
sec_ejec = pd.read_excel(ruta_anexo+'/PLIEGO-UE-SEC_EJEC_NOMBRES.xlsx',sheet_name='SIAF', header=0)

# Crear identificador >>>>>>>VERIFICAR QUE LAS VARIABLES TENGAN FORMATO NUMERICO

sec_ejec.PLIEGO = sec_ejec.PLIEGO.astype(str)
sec_ejec.EJECUTORA = sec_ejec.EJECUTORA.astype(str)

sec_ejec['id'] = sec_ejec['PLIEGO'] + sec_ejec['EJECUTORA']

sec_ejec_i = sec_ejec[['id','SEC_EJEC']] #crea un nuevo DataFrame que contiene solo las columnas 'id' y 'SEC_EJEC' del DataFrame original sec_ejec

sec_ejec_i_g = sec_ejec_i.groupby(['id','SEC_EJEC',]).sum().reset_index()

# Arreglos en la base del reporte

reporte['COD_PLIEGO'] = reporte['COD_PLIEGO'].astype(str)
#reporte['COD_PLIEGO'] = reporte['COD_PLIEGO'].str[:-2]

reporte['COD_UE'] = reporte['COD_UE'].astype(str)
#reporte['COD_UE'] = reporte['COD_UE'].str[:-2]

reporte['id'] = reporte['COD_PLIEGO'] + reporte['COD_UE']

# Combinar bases

reporte_sec_ejec = pd.merge(reporte, sec_ejec_i_g, on =['id'], how ='left') #

# Eliminar espacios de nombres de variables

reporte.columns = reporte.columns.str.strip()

reporte_sec_ejec.columns = reporte_sec_ejec.columns.str.strip()

# Crear columnas para comparar el codigo modular y codigo UGEL

reporte['COD_UGEL_DITEN'] = reporte['COD_UGEL_DITEN'].astype('int64')
reporte_sec_ejec['UGEL_DITEN_vs_UGEL_NT'] = reporte_sec_ejec.apply(lambda row: 'SI' if row['COD_UGEL'] == row['COD_UGEL_DITEN'] else 'NO', axis=1)

# Variables de interés
reporte_i = reporte_sec_ejec[['Dependencia','Dirección General','Dirección Línea','COD_REGION','Región','MACRO_REGION','COD_PLIEGO','Pliego','COD_UE',
                              'Unidad Ejecutora','SEC_EJEC','COD_UGEL','DRE/UGEL','Código de local','Código modular','Anexo','Nombre de la IE','COD_INT',
                              'Intervención','Intervención - nombre corto','COD_CARGO','Cargo','COD_PPR','NOM_PPR','COD_PROD',
                              'NOM_PROD','COD_ACT','NOM_ACT','COD_FUN','NOM_FUN','COD_DIV_FUN','NOM_DIV_FUN','COD_GRUPFUN','NOM_GRUPFUN','Fuente de Financiamiento',
                              'Nivel contratación','Sueldo programado','Sueldo inicial','NC1','NC2','Essalud programado','Aguinaldo programado','Mes inicio programado','Mes fin programado',
                              'N° meses programado','Costo estimado anual de Honorario Programado (S/)','Costo estimado anual de Essalud programado (S/)',
                              'Costo estimado anual de Aguinaldo programado (S/)','Costo estimado anual de la contratación programado (S/)',
                              'Fecha inicio programado Norma Técnica','Fecha fin programado Norma Técnica','Fecha inicio NEXUS','Fecha termino NEXUS','NUMDOCUM','APELLIPAT','APELLIMAT','NOMBRES',
                              'CODPLAZA','SITLAB','Contratado','Vacante','Programadas','Marco Normativo','COD_UGEL_DITEN','UGEL_DITEN','UGEL_DITEN_vs_UGEL_NT']]

# Formatear las columnas como fecha corta

#reporte_i['Fecha inicio NEXUS'] = pd.to_datetime(reporte_i['Fecha inicio NEXUS'])
#reporte_i['Fecha inicio NEXUS'] = reporte_i['Fecha inicio NEXUS'].strftime('%d-%m-%Y')
#reporte_i['Fecha inicio programado Norma Técnica'] = reporte_i['Fecha inicio programado Norma Técnica'].dt.strftime('%d-%m-%Y')
#reporte_i['Fecha termino NEXUS'] = reporte_i['Fecha termino NEXUS'].dt.strftime('%d-%m-%Y')
#reporte_i['Fecha fin programado Norma Técnica'] = reporte_i['Fecha fin programado Norma Técnica'].dt.strftime('%d-%m-%Y')

# Forzar fechas
# Convertir las columnas de fecha a tipo datetime
#reporte_i['Fecha inicio programado Norma Técnica'] = pd.to_datetime(reporte_i['Fecha inicio programado Norma Técnica'], errors='coerce')
#reporte_i['Fecha fin programado Norma Técnica'] = pd.to_datetime(reporte_i['Fecha fin programado Norma Técnica'], errors='ignore')
#reporte_i['Fecha inicio NEXUS'] = pd.to_datetime(reporte_i['Fecha inicio NEXUS'], errors='coerce')
#reporte_i['Fecha termino NEXUS'] = pd.to_datetime(reporte_i['Fecha termino NEXUS'], errors='ignore')


# Formatear las fechas en el formato DD/MM/YYYY
#reporte_i['Fecha inicio programado Norma Técnica'] = reporte_i['Fecha inicio programado Norma Técnica'].dt.strftime('%d/%m/%Y')
#reporte_i['Fecha fin programado Norma Técnica'] = reporte_i['Fecha fin programado Norma Técnica'].dt.strftime('%d/%m/%Y')

#reporte_i['Fecha inicio NEXUS'] = reporte_i['Fecha inicio NEXUS'].apply(lambda x: '' if pd.isnull(x) else x.strftime('%d/%m/%Y'))
#reporte_i['Fecha termino NEXUS'] = reporte_i['Fecha termino NEXUS'].apply(lambda x: '' if pd.isnull(x) else x.strftime('%d/%m/%Y'))




#peas_prog_cod_corte['FECINICIO'] = pd.to_datetime(peas_prog_cod_corte['FECINICIO'], format='%Y-%m-%d %H:%M:%S')
#peas_prog_cod_corte['Fecha inicio NEXUS'] = peas_prog_cod_corte['Fecha inicio NEXUS'].dt.strftime('%d/%m/%y')


#e) Exportar reporte

reporte_i.to_excel(ruta_output+'/ReporteNexus_'+fecha+'_plantilla.xlsx', sheet_name=mes_corte , index= False)

xxxx

'''
        ETAPA 5: Insumo para tablero 

'''

#a) Variables de interés

pow_bi = reporte[['Dependencia','Dirección General','Dirección Línea','COD_REGION','Región','MACRO_REGION','COD_PLIEGO','Pliego','COD_UE','Unidad Ejecutora','COD_UGEL',
                  'DRE/UGEL','Código de local','Código modular','Anexo','Nombre de la IE','COD_INT','Intervención','Intervención - nombre corto','COD_CARGO','Cargo',
           'COD_PPR','NOM_PPR','COD_PROD','NOM_PROD','COD_ACT','NOM_ACT','COD_FUN','NOM_FUN','COD_DIV_FUN','NOM_DIV_FUN',
                  'COD_GRUPFUN','NOM_GRUPFUN','Fuente de Financiamiento','Nivel contratación','Sueldo programado','Essalud programado','Aguinaldo programado',
                  'Mes inicio programado','N° meses programado','Costo estimado anual de Honorario Programado (S/)','Costo estimado anual de Essalud programado (S/)',
                  'Costo estimado anual de Aguinaldo programado (S/)','Costo estimado anual de la contratación programado (S/)','Fecha inicio programado Norma Técnica',
                  'Fecha inicio NEXUS','Fecha termino NEXUS','NUMDOCUM','APELLIPAT','APELLIMAT','NOMBRES','CODPLAZA','SITLAB','Contratado','Vacante','Programadas','Marco Normativo']]

# Generar numeración para mes de inicio

pow_bi['num_mes_inicio'] = 1

pow_bi.loc[(pow_bi['Mes inicio programado']=='Febrero'),'num_mes_inicio'] = 2
pow_bi.loc[(pow_bi['Mes inicio programado']=='Marzo'),'num_mes_inicio'] = 3
pow_bi.loc[(pow_bi['Mes inicio programado']=='Abril'),'num_mes_inicio'] = 4
pow_bi.loc[(pow_bi['Mes inicio programado']=='Mayo'),'num_mes_inicio'] = 5
pow_bi.loc[(pow_bi['Mes inicio programado']=='Junio'),'num_mes_inicio'] = 6
pow_bi.loc[(pow_bi['Mes inicio programado']=='Julio'),'num_mes_inicio'] = 7
pow_bi.loc[(pow_bi['Mes inicio programado']=='Agosto'),'num_mes_inicio'] = 8

#b) Generar fecha de corte

pow_bi['fecha_corte_id'] = fecha_corte_id

pow_bi['fecha_corte'] = fecha_corte

# Generar mes de corte

pow_bi['mes_corte'] = mes_corte

pow_bi['num_mes_corte'] = num_mes_corte

# Agrupar

pow_bi_g=pow_bi.groupby(['Dependencia','Dirección General','Dirección Línea','COD_REGION','Región','MACRO_REGION','COD_PLIEGO','Pliego','COD_UE','Unidad Ejecutora','COD_UGEL','DRE/UGEL','COD_INT','Intervención','Intervención - nombre corto','COD_CARGO','Cargo','COD_PPR','NOM_PPR','COD_PROD','NOM_PROD','COD_ACT','NOM_ACT','num_mes_inicio','Mes inicio programado','fecha_corte_id','fecha_corte','num_mes_corte','mes_corte'])[['Contratado','Vacante','Programadas']].sum().reset_index()

# Renombrar variables

pow_bi_g.rename(columns={'Contratado':'Contratado_NEXUS'},inplace=True)
pow_bi_g.rename(columns={'Vacante':'Vacante_NEXUS'},inplace=True)

# Variables de interés

reporte_pow_bi = pow_bi_g[['Dependencia','Dirección General','Dirección Línea','COD_REGION','Región','MACRO_REGION','COD_PLIEGO','Pliego','COD_UE','Unidad Ejecutora','COD_UGEL','DRE/UGEL','COD_INT','Intervención','Intervención - nombre corto','COD_CARGO','Cargo','COD_PPR','NOM_PPR','COD_PROD','NOM_PROD','COD_ACT','NOM_ACT','num_mes_inicio','Mes inicio programado','fecha_corte_id','fecha_corte','num_mes_corte','mes_corte','Contratado_NEXUS','Vacante_NEXUS','Programadas']]

#c) Exportar

reporte_pow_bi.to_excel(ruta_bi+'/reporte_nexus_pbi_'+fecha+'.xlsx', sheet_name='nexus', index= False)



'''
        ETAPA 5: Observaciones

'''

# a) Usar base almacenada del último corte

obs

# Arreglos

obs.CODREG=obs.CODREG.astype(str)
obs['CODREG']= obs['CODREG'].str.zfill(6)

obs.CODUGE=obs.CODUGE.astype(str)
obs['CODUGE']= obs['CODUGE'].str.zfill(6)

peas_prog_cod.COD_UGEL=peas_prog_cod.COD_UGEL.astype(str)
peas_prog_cod['COD_UGEL']= peas_prog_cod['COD_UGEL'].str.zfill(6)

# b) Generar hoja fecha de inicio no compatible

# Cambiar el tipo de variable a fecha

#obs['FECINICIO'] = pd.to_datetime(obs['FECINICIO']).dt.date

finicio_ncomp = obs

# c) Generar cargo no compatible

obs_peas = pd.merge(obs, peas_prog_cod, on = ['CODPLAZA'], how='left', indicator=True)

# Código cargo diferente

obs_peas['dif CODCARGO'] = 'CODCARGO INCORRECTO'
obs_peas.loc[(obs_peas['CODCARGO']==obs_peas['COD_CARGO']),'dif CODCARGO'] = 'CODCARGO CORRECTO'

cargo_ncomp = obs_peas[['CODREG','DESCREG','CODUGE','NOMBREOOII','CODMODCE','CODNIVEDUC','NOMBIE','TIPOPLAZA','SUBTIPOPLAZA','CODPLAZA','ESTPLAZA','SITLAB','FECINICIO','FECTERMINO','NUMDOCUM','APELLIPAT','APELLIMAT','NOMBRES','CODCARGO','DESCARGO','Intervención - nombre corto','COD_CARGO','Cargo','dif CODCARGO']]

# d) Generar UGEL no compatible

obs_peas['dif COD_UGEL'] = 'COD_UGEL INCORRECTO'
obs_peas.loc[(obs_peas['CODUGE']==obs_peas['COD_UGEL']),'dif COD_UGEL'] = 'COD_UGEL CORRECTO'

ugel_ncomp = obs_peas[['CODREG','DESCREG','CODUGE','NOMBREOOII','CODMODCE','CODNIVEDUC','NOMBIE','TIPOPLAZA','SUBTIPOPLAZA','CODPLAZA','ESTPLAZA','SITLAB','FECINICIO','FECTERMINO','NUMDOCUM','APELLIPAT','APELLIMAT','NOMBRES','CODCARGO','DESCARGO','Intervención - nombre corto','COD_UGEL','DRE/UGEL','dif COD_UGEL']]

# e) Exportar

excel_writer_obs = pd.ExcelWriter(ruta_rep+'/cas_'+fecha+'_observaciones - falta evaluar fecha de inicio.xlsx')

# Exportar cada dataframe en una hoja con su nombre respectivo

obs.to_excel(excel_writer_obs, sheet_name='cas_'+fecha, index=False)
finicio_ncomp.to_excel(excel_writer_obs, sheet_name='fecha inicio no comp', index=False)
cargo_ncomp.to_excel(excel_writer_obs, sheet_name='cargo no comp', index=False)
ugel_ncomp.to_excel(excel_writer_obs, sheet_name='ugel no comp', index=False)

# Guarda el archivo Excel

excel_writer_obs.save()

# Cierra el objeto ExcelWriter

excel_writer_obs.close()


'''
        ETAPA 6: Evaluar fechas
'''

#a) Importar la plantilla del reporte

column_types_4 = {'COD_CARGO': str}

rep = pd.read_excel(ruta_output+'\\ReporteNexus_'+fecha+'_plantilla.xlsx',sheet_name=mes_corte, header=0, dtype=column_types_4)

# Seleccionar solo los contratados

rep_con = rep[rep.Contratado == 1]

#b) Crear la variable fecha para asignar identificar aquellas plazas cuya fecha de inicio sea anterior a la fecha de inicio programada por la Norma Técnica

rep_con['observación1'] = (rep_con['Fecha inicio NEXUS'] < rep_con['Fecha inicio programado Norma Técnica']).astype(int)   

#b) Crear la variable fecha para asignar identificar aquellas plazas cuya fecha de inicio sea posterior en 15d a la fecha de corte

rep_con['observación2'] = 0

# Calcular la fecha límite (fecha de hoy más 15 días)
import pandas as pd
import numpy as np  # Make sure to include this import
from datetime import datetime, timedelta

# fecha de CORTE en formato de cadena 'ddmmyyyy'
fecha_corte_str = '28062024'

# Convertir la fecha de hoy a un objeto datetime
fecha_corte_ = datetime.strptime(fecha_corte_str, '%d%m%Y')

# Calcular la fecha límite (fecha de hoy más 15 días)
fecha_limite = fecha_corte_ + timedelta(days=15)

# Reemplazar la variable 'observación2' por 1 si 'Fecha inicio NEXUS' es mayor en 15 días a la fecha de corte

rep_con['Fecha inicio NEXUS'] = pd.to_datetime(rep_con['Fecha inicio NEXUS'], format='%d-%m-%Y')
rep_con['observación2'] = np.where(pd.to_datetime(rep_con['Fecha inicio NEXUS'], format='%d%m%Y') > fecha_limite, 1, rep_con['observación2'])

# Filtrar y mantener solo las observaciones en las que ambas variables no son cero
rep_con = rep_con.loc[(rep_con['observación1'] != 0) | (rep_con['observación2'] != 0)]

#dar formato general a la fecha de inicio NEXUS
rep_con['Fecha inicio NEXUS'] = rep_con['Fecha inicio NEXUS'].dt.strftime('%d-%m-%Y')

#d) Crear diccionario

dic_data = {
    'Columna': ['observación 1', 'observación 2'],
    'Toman el valor de 1:': [
        'Toda aquella PEA cuya fecha de inicio sea anterior a la fecha de inicio programada por la Norma Técnica',
        'Toda aquella PEA cuya fecha de inicio sea posterior en 15 días a la fecha de corte ('+fecha_corte+')'
    ]
}

dic = pd.DataFrame(dic_data)

#e) Exportar

excel_writer_observaciones = pd.ExcelWriter(ruta_rep+'/cas_'+fecha+'_observaciones.xlsx')

# Exportar cada dataframe en una hoja con su nombre respectivo

rep_con.to_excel(excel_writer_observaciones, sheet_name=mes_corte, index=False)
dic.to_excel(excel_writer_observaciones, sheet_name='Diccionario', startrow=1, startcol=1, index=False)

# Guarda el archivo Excel

excel_writer_observaciones.save()

# Cierra el objeto ExcelWriter

excel_writer_observaciones.close()









