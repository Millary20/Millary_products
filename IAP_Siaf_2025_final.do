/*******************************************************************************
Área:			Unidad de Planificación y Presupuesto(UPP)
Responsable: 	Millary Antunez
Objetivos: 		Costear Limpieza y Mantenimiento 2025
Fecha: 			Diciembre, 2024
*******************************************************************************/
**Breve Resumen de los criterios que cambiaron con respecto al 2023

* No cuenta con gasto de pliego
* Hubo actualizaciones en el padrón
* 2025 se actualizó UIT y %impositivo
*******************************************************************************/
clear 
set more off
global ruta 	"B:\OneDrive - Ministerio de Educación\unidad_B\2025\3. Intervenciones y acciones\19. Limpieza y mantenimiento"

global input 	"$ruta/Input"
global pxq 		"$ruta/PxQ"
global output 	"$ruta/Output"

cd "$global"
/***************************************************
		Etapa 0. Definición de variables
****************************************************/

*Variables
local n_mes_cas 11 // meses de cas 
local cas_activo "feb mar abr may jun jul ago sep oct nov dic"
local cas_inactivo "ene"

*Vacaciones truncas
*local cas_activo_vt "ene"
*local cas_inactivo_vt "feb mar abr may jun jul ago sep oct nov dic"

*local cas_vt "oct"
*local no_cas_vt "ene feb mar abr may jun jul ago sep nov dic"

local rem_PEAS 1414.19

*Aguinaldo
local monto_agui 300 // aguinaldo, uno al año
local agui  "jul dic"
local 0_agui "ene feb mar abr may jun ago sep oct nov"

*Essalud - UIT 
local UIT 5350

*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*
*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_WORMHOLE_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*
*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*

/***************************************************
		Etapa 1. Cálculo de metas físicas 
****************************************************/
***Importar el padrón de la dirección*********
import excel "$pxq/PxQ_LyM 2025_ACTUALIZADOS CON EL AUMENTO DE 100 (1)_UIT", sheet("Padrón codlocal") cellrange(B1:M1285) firstrow clear

drop LyM_10 L

gen anexo=0
destring cod_pliego cod_ue codmod codlocal, replace

/*
/*Insertamos el padrón web*/
tempfile padronweb
preserve
use "$input\Padron_web_20231215", clear
keep if ESTADO=="1" // activa
keep if inlist(NIV_MOD,"A1","A2","A3","B0","F0") // Inicial, primaria y secundaria
keep if inlist(GESTION,"1","2") // gestion=="1" pública de gestión directa, "2" pública de gestión privada
rename (COD_MOD CODLOCAL ANEXO CODOOII)  (codmod codlocal anexo codooii)
keep codmod codlocal anexo codooii CEN_EDU TALUMNO TDOCENTE NIV_MOD // nombre del centro educativo y lo que falte
destring codmod codlocal anexo codooii , replace
save `padronweb'
restore
*/

/*Insertamos el padrón web 2024*/
tempfile padronweb
preserve

import dbase using "$input\Padron_web_20241129.dbf", clear
*use "$input\Padron_web_20231215", clear // padron web 2023
keep if ESTADO=="1" // activa
keep if inlist(NIV_MOD,"A1","A2","A3","B0","F0") // Inicial, primaria y secundaria
keep if inlist(GESTION,"1","2") // gestion=="1" pública de gestión directa, "2" pública de gestión privada
rename (COD_MOD CODLOCAL ANEXO CODOOII)  (codmod codlocal anexo codooii)

keep codmod codlocal anexo codooii CEN_EDU TALUMNO TDOCENTE NIV_MOD 

destring codooii , replace
destring codmod   ,replace
destring codlocal   ,replace
destring anexo   ,replace

save `padronweb'
restore

*Unimos el Padrón de LyM con el padrón
destring codooii , replace
merge 1:1  codooii codlocal anexo codmod using `padronweb' 
*tab _m
keep if _merge==3 | _merge==1
drop _merge


***************
*Metas físicas
****************
preserve
order cod_pliego nom_pliego cod_ue nom_ue codooii ugel codlocal codmod alumnos anexo CEN_EDU TALUMNO TDOCENTE
export excel "$output/Metas_fisicas_Lym_2024", firstrow(variables) replace
restore


/***************************************************
		Etapa 2. Cálculo de padrón 
****************************************************/

*Usamo la base de ubigeo*

tempfile ubigeo
preserve
import excel "$input\base_ue_ugel_ubigeo_2023.xlsx", sheet("base") firstrow clear
keep PLIEGO NOM_PLIEGO NOM_UE EJECUTORA UGEL CODOOII 
rename (PLIEGO  EJECUTORA CODOOII ) (cod_pliego  cod_ue codooii)
destring codooii cod_pliego cod_ue, replace

drop if NOM_UE=="301. COLEGIO MILITAR PEDRO RUIZ GALLO" | NOM_UE=="301. COLEGIO MILITAR LEONCIO PRADO" |NOM_UE=="301. COLEGIO MILITAR FRANCISCO BOLOGNESI" | NOM_UE=="310. COLEGIO MILITAR RAMON CASTILLA" | NOM_UE=="301. COLEGIO  MILITAR ELIAS AGUIRRE"
//SE QUITAN COLEGIOS MILITARES 

drop if NOM_UE=="304. INSTITUTOS SUP. DE EDUC. PUBL. REGIONAL DE PIURA"| NOM_UE=="022. INSTITUTO PEDAGOGICO NACIONAL MONTERRICO"
drop if codooii==.
rename codooii cod_ugel
save `ubigeo', replace

*Cuadramos con los datos del pxq
restore

preserve
rename codooii cod_ugel
duplicates drop cod_ugel,force
keep cod_pliego nom_pliego cod_ue nom_ue cod_ugel ugel 

rename nom_ue nom_ue_pxq
rename nom_pliego nom_pliego_pxq
rename cod_pliego cod_pliego_pxq
rename cod_ue cod_ue_pxq

merge 1:1 cod_ugel using `ubigeo' 
keep if _merge==3
restore

*Se revisa si el pliego, ejecutora y codigos coinciden con el pxq
*Se verifica que la única diferencia proviene de la UGEL Rio Mantaron, pero se concluye que la UGEL RIO Mantaro se seguirá considerando en la UE Satipo y no en la UE Rio Mantaro para el 2024. 

*Entonces consideramos el pxq de la direccion

/***************************************************
	*Padron para el PxQ
***************************************************/
rename codooii cod_ugel
preserve
gen PEAS=1
keep cod_pliego nom_pliego cod_ue nom_ue cod_ugel ugel codlocal alumnos codmod anexo TALUMNO TDOCENTE PEAS
export excel "$output/Padron_Lym_anexo.xlsx", firstrow(variables) replace 
restore 

gen PEAS=1

gen cod_grupo_funcional= substr(grupo_funcional, 1, 4)
destring cod_grupo_funcional,replace


keep cod_pliego nom_pliego cod_ue nom_ue cod_ugel ugel cod_grupo_funcional PEAS

collapse (sum) PEAS, by(cod_pliego cod_ue cod_ugel cod_grupo_funcional)

	
/***************************************************
	Etapa 2. Determinacion de costos
****************************************************/

/***************************************************
	Cálculo de costo contratación CAS
****************************************************/

* REMUNERACION

* Personal de Limpieza y Mantenimiento (PEAS)
gen cas_PEAS_anual = `rem_PEAS' * PEAS * `n_mes_cas'

*_t
foreach mes in `cas_activo'{
gen cas_PEAS_`mes' = `rem_PEAS' * PEAS
}

foreach mes in `cas_inactivo'{
gen cas_PEAS_`mes' = 0
}

*AGUINALDO

gen agui_PEAS_anual = PEAS * `monto_agui'*2 

foreach mes in `agui'{
gen agui_PEAS_`mes' = PEAS * `monto_agui'
}

foreach mes in `0_agui'{
gen agui_PEAS_`mes' = 0
}

* ESSALUD banda: 9% * (RMV - 45%UIT)
gen tope_essalud = ceil(0.09 * (0.45 * `UIT'))
*Aporte a Essalud individual
gen essalud_PEAS = ceil(0.09 * `rem_PEAS')
replace essalud_PEAS = tope_essalud if essalud_PEAS > tope_essalud

drop tope_essalud

gen ess_PEAS_anual = PEAS * essalud_PEAS * `n_mes_cas'

foreach mes in `cas_activo'{
gen ess_PEAS_`mes' = PEAS * essalud_PEAS
}

foreach mes in `cas_inactivo'{
gen ess_PEAS_`mes' = 0
}


* VAMOS A REEMPLAZAR LOS VACIOS POR CERO
* Listar las variables numéricas en el conjunto de datos
ds, has(type numeric)

* Generar una lista de las variables numéricas
local num_vars `r(varlist)'

* Reemplazar los valores faltantes por cero en las variables numéricas
foreach var of local num_vars {
	replace `var' = 0 if missing(`var')
}

/*
foreach var in PEAS cas_PEAS_anual cas_PEAS_mar cas_PEAS_abr cas_PEAS_may cas_PEAS_jun cas_PEAS_jul cas_PEAS_ago cas_PEAS_sep cas_PEAS_ene cas_PEAS_feb cas_PEAS_oct cas_PEAS_nov cas_PEAS_dic agui_PEAS_anual agui_PEAS_jul agui_PEAS_ene agui_PEAS_feb agui_PEAS_mar agui_PEAS_abr agui_PEAS_may agui_PEAS_jun agui_PEAS_ago agui_PEAS_sep agui_PEAS_oct agui_PEAS_nov agui_PEAS_dic essalud_PEAS ess_PEAS_anual ess_PEAS_mar ess_PEAS_abr ess_PEAS_may ess_PEAS_jun ess_PEAS_jul ess_PEAS_ago ess_PEAS_sep ess_PEAS_ene ess_PEAS_feb ess_PEAS_oct ess_PEAS_nov ess_PEAS_dic {
  
  
  
   replace `var' = 0 if missing(`var')
}
*/

* =======================================================
*                    Bases SIAF
* ========================================================

*Pasar a nivel ugel

*collapse(sum) cas_* agui_* ess_PEAS_* costo_vt_total_* costo_vt2024_*, by(cod_pliego cod_ue cod_ugel cod_grupo_funcional)
collapse(sum) cas_* agui_* ess_PEAS_*, by(cod_pliego cod_ue cod_ugel cod_grupo_funcional)

****RESHAPE****
rename cas_PEAS_*                 name1_* , replace
rename agui_PEAS_*                name2_* , replace
rename ess_PEAS_*                 name3_* , replace
*rename costo_vt_total_*      	  name4_* , replace
*rename costo_vt2024_*      	  name5_* , replace

rename *ene *1 , replace
rename *feb *2 , replace
rename *mar *3 , replace
rename *abr *4 , replace
rename *may *5 , replace
rename *jun *6 , replace
rename *jul *7 , replace
rename *ago *8 , replace
rename *sep *9 , replace
rename *oct *10, replace
rename *nov *11, replace
rename *dic *12, replace
rename *anual   *13, replace


*gen cod_grupo_funcional=.
*replace cod_grupo_funcional=103 if grupo_funcional=="0103. EDUCACION INICIAL"
*replace cod_grupo_funcional=104 if grupo_funcional=="0104. EDUCACION PRIMARIA"
*replace cod_grupo_funcional=105 if grupo_funcional=="0105. EDUCACION SECUNDARIA"

*drop grupo_funcional

reshape long name, i(cod_pliego cod_ue cod_ugel cod_grupo_funcional) j(s) string

generate valor = real(regexs(0)) if regexm(s, "^[0-9]+")
generate mes = real(regexs(0)) if regexm(s, "[0-9]+$")

drop s

reshape wide name, i(cod_pliego cod_ue cod_ugel cod_grupo_funcional valor) j(mes)

/***************************************************
            BASE SIAF MINEDU
****************************************************/	
**Funcion
gen cod_func= 22
gen funcion ="22. EDUCACION"
**Division funcional
gen cod_divfunc= 47
gen division_funcional="047. EDUCACION BASICA"

**Programa Presupuestal
gen cod_pp=90 
gen programa_presupuestal="0090. LOGROS DE APRENDIZAJE DE ESTUDIANTES DE LA EDUCACION BASICA REGULAR" 

**Producto
gen cod_prod ="3000385" 
gen producto_proy= "3000385.  INSTITUCIONES EDUCATIVAS CON CONDICIONES PARA EL CUMPLIMIENTO DE HORAS LECTIVAS NORMADAS" 

**Actividad
gen cod_act="5005629" 
gen actividad_obra="5005629. CONTRATACION OPORTUNA Y PAGO DEL PERSONAL ADMINISTRATIVO Y DE APOYO DE LAS INSTITUCIONES EDUCATIVAS DE EDUCACION BASICA REGULAR"

gen corr = "1.1.13.1.2." if valor==1 //cas
replace corr = "1.1.9.1.4." if valor==2 //aguinaldo
replace corr = "1.3.1.1.15." if valor==3 //essalud
*replace corr = "1.4.1.1.6." if valor>=4 //Vacaciones truncas


gen componente=""
replace componente="CONTRATACION CAS" if valor>=1 & valor<=3
*replace componente="VACACIONES TRUNCAS 2023" if valor==4
*replace componente="VACACIONES TRUNCAS 2024" if valor==5

display regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")
display regexs(0)
display regexs(1)
display regexs(2) 
display regexs(3)
display regexs(4)
display regexs(5)
gen cod_gen =regexs(1) if regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")
gen cod_subgg =regexs(2) if regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")
gen cod_subgg2 =regexs(3) if regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")
gen cod_espec =regexs(4) if regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")
gen cod_espec2 =regexs(5) if regexm(corr, "([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)\.([0-9]*)")


/***************************************************
           IMPORTANDO ETIQUETAS
****************************************************/	

cd "$input"

preserve

*Importamos etiqueta del corr 

tempfile generica
*import excel "$general/base_generica.xlsx", firstrow clear
import excel "base_generica.xlsx", firstrow clear
save `generica', replace
restore 

merge m:1 corr using `generica'
keep if _merge==3
drop _merge


*Importamos etiqueta de UE, Pliego, UGEL
destring cod_ugel,replace
merge m:1 cod_pliego cod_ue cod_ugel using `ubigeo'
drop if _merge==2
drop _merge

*replace NOM_PLIEGO="450. GOBIERNO REGIONAL DEL DEPARTAMENTO DE JUNIN" if cod_ugel==120014
*replace NOM_UE="302. EDUCACION SATIPO" if cod_ugel==120014
*replace UGEL="UGEL RIO ENE - MANTARO" if cod_ugel==120014


forval x=1/13 {
	
	replace name`x'=0 if name`x'==.
	
}
*	
rename name13 costo_anual


/***************************************************
            VALIDACIÓN DE CÁLCULOS
****************************************************/	

egen anual=rowtotal (name*)

compare costo_anual anual

drop anual

drop if costo_anual==0

rename *e1 enero, replace
rename *e2 febrero, replace
rename *e3 marzo, replace
rename *e4 abril, replace
rename *e5 mayo, replace
rename *e6 junio, replace
rename *e7 julio, replace
rename *e8 agosto, replace
rename *e9 septiembre, replace
rename *e10 octubre, replace
rename *e11 noviembre, replace
rename *e12 diciembre, replace




gen grupo_funcional=""
replace grupo_funcional="0103. EDUCACION INICIAL" if cod_grupo_funcional==103
replace grupo_funcional="0104. EDUCACION PRIMARIA" if cod_grupo_funcional==104
replace grupo_funcional="0105. EDUCACION SECUNDARIA" if cod_grupo_funcional==105


**uniformización Edmar
rename (cod_pliego cod_ue cod_ugel cod_pp cod_prod cod_act cod_func cod_divfunc cod_grupo_funcional cod_gen cod_subgg cod_subgg2 cod_espec cod_espec2) (COD_PLIEGO COD_UE COD_UGEL COD_PPR COD_PROD COD_ACT COD_FUN COD_DIV_FUN COD_GRUPFUN COD_GEN COD_SUBG COD_SUBGT COD_ESP COD_ESPT)
gen COD_FTE=1
gen FUENTE="1. RECURSOS ORDINARIOS"
gen intervencion="Limpieza y Mantenimiento"
gen cod_intervencion="20"

drop valor 

order COD_PLIEGO NOM_PLIEGO  COD_UE NOM_UE COD_UGEL UGEL COD_PPR programa_presupuestal COD_PROD producto_proy COD_ACT actividad_obra  COD_FUN funcion COD_DIV_FUN division_funcional COD_GRUPFUN grupo_funcional COD_FTE FUENTE COD_GEN generica COD_SUBG subgenerica COD_SUBGT subgenerica_det COD_ESP especifica  COD_ESPT especifica_det   corr correlativo intervencion cod_intervencion componente enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre costo_anual        

tostring COD_PLIEGO, replace
tostring COD_UE, replace
tostring COD_UGEL, replace
forval length=1/6 {
			replace COD_UGEL = "0"+ COD_UGEL if length(COD_UGEL)<`length'
	}


preserve
drop if corr == "1.4.1.1.6."
cd "$output"
export excel "Lym_2025_siaf_mod_componente.xlsx", firstrow(variables) replace
save "Lym_2025_siaf_mod_componente.dta", replace

restore
	
/***************************************************
            VACACIONES TRUNCAS
****************************************************/
/*
preserve

keep if corr == "1.4.1.1.6."

*Guardar

cd "$output"
export excel "Lym_2024_vacaciones_truncas.xlsx", firstrow(variables) replace
save "Lym_2024_vacaciones_truncas.dta", replace

restore	
*/
/***************************************************
            BASE SIAF MEF
****************************************************/


tempfile ubigeo_2
preserve
import excel "$input\base_ue_ugel_ubigeo_2023.xlsx", sheet("base") firstrow clear
drop if NOM_UE=="301. COLEGIO MILITAR PEDRO RUIZ GALLO" | NOM_UE=="301. COLEGIO MILITAR LEONCIO PRADO" |NOM_UE=="301. COLEGIO MILITAR FRANCISCO BOLOGNESI" | NOM_UE=="310. COLEGIO MILITAR RAMON CASTILLA" | NOM_UE=="301. COLEGIO  MILITAR ELIAS AGUIRRE"
//SE QUITAN COLEGIOS MILITARES 
drop if NOM_UE=="304. INSTITUTOS SUP. DE EDUC. PUBL. REGIONAL DE PIURA"| NOM_UE=="022. INSTITUTO PEDAGOGICO NACIONAL MONTERRICO"
drop if CODOOII==.
rename CODOOII COD_UGEL
	tostring COD_UGEL, replace
	forval length=1/6 {
			replace COD_UGEL = "0"+ COD_UGEL if length(COD_UGEL)<`length'
	}
	
rename PLIEGO COD_PLIEGO
rename EJECUTORA COD_UE

tostring COD_PLIEGO, replace
tostring  COD_UE, replace

save `ubigeo_2', replace
restore	


	merge m:1 COD_PLIEGO COD_UE COD_UGEL using `ubigeo_2', keepusing(DEPARTAMENTO NOMBRE_DEPARTAMENTO PROVINCIA NOMBRE_PROVINCIA DISTRITO NOMBRE_DISTRITO) 

	*Completamos UGEL RIO ENE - MANTARO
replace DEPARTAMENTO="12" if COD_UGEL=="120014"
replace NOMBRE_DEPARTAMENTO="JUNIN" if COD_UGEL=="120014"
replace PROVINCIA="01" if COD_UGEL=="120014"
replace NOMBRE_PROVINCIA="HUANCAYO" if COD_UGEL=="120014"
replace DISTRITO="35" if COD_UGEL=="120014"
replace NOMBRE_DISTRITO="SANTO DOMINGO DE ACOBAMBA" if COD_UGEL=="120014"
	
drop if _merge==2
drop _merge	
	

	* Generamos variables
	gen ANO_EJE	= 			"2025"
	gen SECTOR	= 			"99"
		replace SECTOR = "10" if COD_PLIEGO == "10"
	gen NOMBRE_SECTOR =		"GOBIERNOS REGIONALES" 
		replace NOMBRE_SECTOR = "EDUCACION" if COD_PLIEGO == "10"
	
	rename COD_FTE FUENTE_FINANC
	rename FUENTE NOMBRE_FUENTE_FINANC
	
	gen RUBRO =	"00"	
	gen NOMBRE_RUBRO = "RECURSOS ORDINARIOS" 	
	gen CAT_GASTO =	"5"
	gen NOMBRE_CAT_GASTO =	"GASTOS CORRIENTES"
	gen TIPO_TRANSACCION =	"2"
	gen NOMBRE_TIPO_TRANSACCION = "GASTOS PRESUPUESTARIOS"	
	gen INTERVENCION = "FORTALECIMIENTO DE LAS INSTITUCIONES EDUCATIVAS FOCALIZADAS PARA CUMPLIR CON LAS CONDICIONES DE BIOSEGURIDAD Y SALVAGUARDAR LA SALUD Y BIENESTAR DE LA COMUNIDAD EDUCATIVA, A TRAVES DE LA CONTRATACION DE PERSONAL DE LIMPIEZA Y MANTENIMIENTO, EN EL MARCO DEL RESTABLECIMIENTO DEL SERVICIO EDUCATIVO EN LAS INSTITUCIONES EDUCATIVAS"

	gen META = "00001"
	gen FINALIDAD = "0161002"
	gen NOMBRE_FINALIDAD = "CONTRATACION OPORTUNA Y PAGO DEL PERSONAL ADMINISTRATIVO Y DE APOYO DE LAS INSTITUCIONES EDUCATIVAS DE EDUCACION BASICA REGULAR"
	gen UNIDAD_MEDIDA = "236"
	gen NOMBRE_UNIDAD_MEDIDA = "INSTITUCION EDUCATIVA"
	
****añadiendo las metas
	*gen CANTIDAD_META = costo_anual/(9*1314.19)
	gen CANTIDAD_META = round(costo_anual/(`n_mes_cas'*`rem_PEAS'))
	*essalud
	replace CANTIDAD_META = costo_anual/(11*128) if COD_ESPT=="15"
	*aguinaldo
	replace CANTIDAD_META = costo_anual/600 if COD_ESPT=="4"
	
	
	* Cambiamos nombres a variables
rename COD_PLIEGO PLIEGO	
rename NOM_PLIEGO NOMBRE_PLIEGO
rename COD_UE EJECUTORA
rename NOM_UE NOMBRE_EJECUTORA
rename  COD_PPR PROG_PRESUPUESTAL
rename programa_presupuestal NOMBRE_PROG_PRESUPUESTAL
rename COD_PROD PRODUCTO_PROYECTO	
rename producto_proy NOMBRE_PRODUCTO_PROY
rename  COD_ACT ACTIVIDAD_OBRA
rename actividad_obra NOMBRE_ACTIVIDAD_NOMBRE	
rename COD_FUN FUNCION	
rename funcion NOMBRE_FUNCION
rename COD_DIV_FUN DIVISION_FUNCIONAL	
rename division_funcional NOMBRE_DIVISION_FUNC
rename COD_GRUPFUN GRUPO_FUNCIONAL
rename grupo_funcional NOMBRE_GRUPO_FUNC
rename COD_GEN GENERICA
rename generica NOMBRE_GENERICA	
rename COD_SUBG SUB_GENERICA	
rename subgenerica NOMBRE_SUB_GENERICA	
rename COD_SUBGT SUB_GENERICA_DET	
rename subgenerica_det NOMBRE_SUB_GENERICA_DET	
rename COD_ESP ESPECIFICA
rename especifica NOMBRE_ESPECIFICA
rename COD_ESPT ESPECIFICA_DET
rename especifica_det NOMBRE_ESPECIFICA_DET
rename componente COMPONENTE
rename enero ENERO
rename febrero FEBRERO
rename marzo MARZO
rename abril ABRIL
rename mayo MAYO
rename junio JUNIO
rename julio JULIO
rename agosto AGOSTO
rename septiembre SEPTIEMBRE
rename octubre OCTUBRE
rename noviembre NOVIEMBRE
rename diciembre DICIEMBRE
rename costo_anual MONTO_PROGRAMADO


replace NOMBRE_PROG_PRESUPUESTAL = substr(NOMBRE_PROG_PRESUPUESTAL,7,10000)
replace NOMBRE_PRODUCTO_PROY =substr(NOMBRE_PRODUCTO_PROY,10,10000)
replace NOMBRE_ACTIVIDAD_NOMBRE =substr(NOMBRE_ACTIVIDAD_NOMBRE,10,10000)
replace NOMBRE_FUNCION =substr(NOMBRE_FUNCION,5,1000)
replace NOMBRE_DIVISION_FUNC=substr(NOMBRE_DIVISION_FUNC,6,10000)
replace NOMBRE_GRUPO_FUNC=substr(NOMBRE_GRUPO_FUNC,7,10000)

replace NOMBRE_EJECUTORA=substr(NOMBRE_EJECUTORA,5,10000)

foreach x in "NOMBRE_PLIEGO" "NOMBRE_FUENTE_FINANC" "NOMBRE_GENERICA" "NOMBRE_SUB_GENERICA" "NOMBRE_SUB_GENERICA_DET" "NOMBRE_ESPECIFICA" "NOMBRE_ESPECIFICA_DET" {

split `x', p(". ")
drop `x'1
replace `x'=`x'2
drop `x'2
}


order ANO_EJE	SECTOR	NOMBRE_SECTOR PLIEGO NOMBRE_PLIEGO	EJECUTORA	NOMBRE_EJECUTORA	PROG_PRESUPUESTAL NOMBRE_PROG_PRESUPUESTAL	PRODUCTO_PROYECTO	NOMBRE_PRODUCTO_PROY ACTIVIDAD_OBRA	NOMBRE_ACTIVIDAD_NOMBRE	FUNCION	NOMBRE_FUNCION	DIVISION_FUNCIONAL	NOMBRE_DIVISION_FUNC GRUPO_FUNCIONAL	NOMBRE_GRUPO_FUNC	META	FINALIDAD	NOMBRE_FINALIDAD	UNIDAD_MEDIDA	NOMBRE_UNIDAD_MEDIDA 	CANTIDAD_META DEPARTAMENTO	NOMBRE_DEPARTAMENTO	PROVINCIA	NOMBRE_PROVINCIA	DISTRITO	NOMBRE_DISTRITO	FUENTE_FINANC	NOMBRE_FUENTE_FINANC	RUBRO	NOMBRE_RUBRO	CAT_GASTO	NOMBRE_CAT_GASTO	TIPO_TRANSACCION	NOMBRE_TIPO_TRANSACCION	GENERICA	NOMBRE_GENERICA SUB_GENERICA	NOMBRE_SUB_GENERICA	SUB_GENERICA_DET	NOMBRE_SUB_GENERICA_DET	ESPECIFICA	NOMBRE_ESPECIFICA	ESPECIFICA_DET	NOMBRE_ESPECIFICA_DET	INTERVENCION COMPONENTE	ENERO	FEBRERO	MARZO	ABRIL	MAYO	JUNIO	JULIO	AGOSTO	SEPTIEMBRE	OCTUBRE	NOVIEMBRE	DICIEMBRE	MONTO_PROGRAMADO

drop COD_UGEL UGEL corr correlativo intervencion cod_intervencion

drop if COMPONENTE=="VACACIONES TRUNCAS 2023" | COMPONENTE=="VACACIONES TRUNCAS 2024"

cd "$output"
export excel "Lym_2025_siaf_mef.xlsx", firstrow(variables) replace
save "Lym_2025_siaf_mef.dta", replace

