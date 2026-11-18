import sys
import cgi
import os
from http import cookies
import json
import locale
import hashlib
from datetime import datetime
from datetime import date
from datetime import timedelta
from io import BytesIO
import qrcode
import openpyxl
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
import random
import string
from os import remove
import csv

import smtplib
from email.message import EmailMessage

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import basename
import shutil
from PyPDF2 import PdfReader,PdfWriter

import socket



sys.path.insert(1,r'C:\inetpub\wwwroot\Portal2\Componentes_2023')
import LibDM_2023
import Dispositivos
locale.setlocale( locale.LC_ALL, '' )
if "HTTP_COOKIE" in os.environ:
    cookievalue = cookies.SimpleCookie(os.environ["HTTP_COOKIE"])
else:
    cookievalue = None
Dir_Raiz = "./Estado_Salidas/inicio.py"
Menu_Activo = "ILC GM"

IP_Zebra = Dispositivos.Zebra_ULC_Saltillo



def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        #Error = DB.Instruccion("INSERT INTO portal.ccorreos_enviar (cce_usuario,cce_asunto,cce_to,cce_contendio) VALUES (2,'TEST','dmarmolejo@goutsi.com','TEST')")
        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-2 mb-2 pb-2 text-center'><i class='mdi mdi-progress-upload'></i> Outbound Status</div>"
        Resultado["Contenido"] += """
        <div id='Res'>
        </div>
        <script>
            $("#fecha").on( "change", function() {
                Cargar_Fecha($("#fecha").val());
            });
            function Cargar_Fecha(Fecha){
                Mostrar_Ventana_Cargando(false);
                var parametros = {"Fun":"Cargar_Fecha","Fecha":Fecha};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Res").html(Resultado["Contenido"]);
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        Mensaje(0,textStatus);
                    }
                });
            }
            $( document ).ready(function() {
                Cargar_Fecha($("#fecha").val());
            });
            
        </script>
        </div>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Cargar_Fecha(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        
        Resultado["Contenido"] += """<div class='text-end pe-1 pt-1'><small class='link-primary' style='cursor:pointer' onclick='Cargar_Fecha($("#fecha").val())'>Actualizar <i class='mdi mdi-refresh'></i></small></div>"""
        Salidas = []
        Salidas_Pendientes = []

        Cajas_Aqui = []
        for Cajas_Liberas_Operaciones in DB.Get_Dato("select * from ulc_saltillo.ccajas where cc_tipo_actual = 'Outbound' AND cc_negocio = 'VOLVO' and cc_activo = 1 and (cc_informacion_actual like '%Despacho%' OR cc_informacion_actual like '%Pending Exit%')"):
            Cajas_Aqui.append(str(Cajas_Liberas_Operaciones["cc_id"]))
        Movimientos = []
        if len(Cajas_Aqui) > 0:
            Movimientos = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master IN ("+str(",".join(Cajas_Aqui))+") ORDER BY cch_fecha_hora")
        
        Correos = []
        if len(Cajas_Aqui) > 0:
            Correos = DB.Get_Dato("SELECT * FROM ilc_saltillo.ccorreos_historico WHERE cch_caja IN ("+str(",".join(Cajas_Aqui))+")")

        for Cajas_Liberas_Operaciones in DB.Get_Dato("select * from ulc_saltillo.ccajas where cc_tipo_actual = 'Outbound' AND cc_negocio = 'VOLVO' and cc_activo = 1 and (cc_informacion_actual like '%Despacho%' OR cc_informacion_actual like '%Pending Exit%') and cc_atrapa_trailer = 0"):
            Caja = Cajas_Liberas_Operaciones
            Info_Caja = json.loads(Caja["cc_informacion_actual"])
            Fecha_Salida = None
            Caja_N = ""
            if Caja is not None:
                Caja_N = "<a style='cursor:pointer' class='link-primary' onclick='Ver_Historico("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")'>"+str(Caja["cc_contenedor"])+"</a>"
            
            Reportes_Ya = 0
            REPORTE_FALTA_CAJA =  ""
            for C in Correos:
                if str(C["cch_caja"]) == str(Caja["cc_id"]):
                    Reportes_Ya += 1
                    
            if Reportes_Ya == 0:
                REPORTE_FALTA_CAJA += "<div class='w-100 text-center'> "+str(Reportes_Ya)+"</div>"
            else:
                REPORTE_FALTA_CAJA += "<div class='w-100 text-center fw-bold'><span onclick='Ver_Historico_Correo("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='link-danger' style='cursor:pointer'> "+str(Reportes_Ya)+" </span></div>"

            Destino = ""
            
            Carrier = ""
            Entrada = ""
            Asignacion = ""
            Operaciones = ""
            Ya_Libero_Operaciones = False
            Man_Doc_Impreso_WMS = ""
            Ya_Man_Doc_Impreso_WMS = False
            Ya_Macro = False
            Correo_Listo = ""
            SELLO_FOTO = ""
            MACRO = ""
            Folio_Prisma = ""
            SALIDA = None
            SALIDA_Fecha = None
            POD_Salida = None
            Cancelado = ""
            if Caja is not None:
                Movimientos_Aqui = []
                for M in Movimientos:
                    if int(M["cch_master"]) == int(Caja["cc_id"]):
                        Movimientos_Aqui.append(M)
                Info_Ahora = json.loads(str(Caja["cc_informacion_actual"]))
                
                Load_ID_N = ""
                Carrier = str(Info_Ahora["Carrier"])
                Entrada = str(Caja["cc_fecha_hora"].strftime("%H:%M<br>%m-%d"))
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "ASIGNAR":
                        Asignacion = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "LIBERA OPERACIONES":
                        Operaciones = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                        Ya_Libero_Operaciones = True
                
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "DOCUMENTOS WMS":
                        Man_Doc_Impreso_WMS = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                        Ya_Man_Doc_Impreso_WMS = True
                
                
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "CORREO CAJA LISTA ENVIADO":
                        Correo_Listo = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                        break
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "SALIDA":
                        SALIDA = M["cch_fecha_hora"]
                        SALIDA_Fecha = M["cch_fecha_hora"]
                    if M["cch_movimiento"] == "ELIMINADO POR YARD":
                        SALIDA = M["cch_fecha_hora"]
                        SALIDA_Fecha = M["cch_fecha_hora"]
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "SELLADA":
                        SELLO_FOTO = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "MACRO":
                        Ya_Macro = True
                        MACRO = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")

                for M in Movimientos_Aqui:
                    if M["cch_movimiento"] == "POD Salida":
                        POD_Salida = M["cch_fecha_hora"].strftime("%H:%M<br>%m-%d")
                
                

                if Entrada != "":
                    Entrada = "<div class='d-flex justify-content-center align-items-center bg-success-subtle fw-bold h-100 w-100'><div><small>"+str(Entrada)+"</small></div></div>"
                if Asignacion != "":
                    Asignacion = "<div class='d-flex justify-content-center align-items-center bg-success-subtle fw-bold h-100 w-100'><div><small>"+str(Asignacion)+"</small></div></div>"
                if Operaciones == "":
                    if Caja["cc_dock"] is not None:
                        Operaciones = "<div class='text-danger bg-warning-subtle fw-bold Pulsa_Texto w-100 h-100 text-denter d-flex justify-content-center align-items-center'><div>"+str(Caja["cc_dock"])+"</div></div>"
                    else:
                        Operaciones = "<div class='text-danger bg-warning-subtle fw-bold Pulsa_Texto w-100 h-100 text-denter d-flex justify-content-center align-items-center'><div><i class='mdi mdi-decagram'></i></div></div>"
                else:
                    Operaciones = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold w-100 h-100'><div><small>"+str(Operaciones)+"</small></div></div>"
                if Ya_Libero_Operaciones == True:
                    if Man_Doc_Impreso_WMS == "":
                        Man_Doc_Impreso_WMS = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='Documentos_WMS_N("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                    else:
                        Man_Doc_Impreso_WMS = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold h-100 w-100'><div><small>"+str(Man_Doc_Impreso_WMS)+"</small></div></div>"
                if Correo_Listo == "":
                    if "Exportacion" in Info_Caja.keys():
                        if Ya_Libero_Operaciones == True and Ya_Man_Doc_Impreso_WMS == True and Ya_Macro == True:
                            Correo_Listo = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='Enviar_Correo_N("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                    else:
                        if Ya_Libero_Operaciones == True and Ya_Man_Doc_Impreso_WMS == True:
                            Correo_Listo = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='Enviar_Correo_N("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                else:
                    Correo_Listo = "<a class='link-warning h-100 w-100' style='curso:pointer' onclick='Re_Enviar_Correo_N("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")'><div class='d-flex justify-content-center align-items-center bg-success fw-bold h-100 w-100'><div><small>"+str(Correo_Listo)+"</small></div></div></a>"
                
                
                if "Exportacion" in Info_Caja.keys():
                    if SALIDA is None:
                        if SELLO_FOTO == "":
                            if Ya_Libero_Operaciones == True:
                                SELLO_FOTO = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='Sellado_Foto("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                        else:
                            SELLO_FOTO = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold h-100 w-100'><div><small>"+str(SELLO_FOTO)+"</small></div></div>"
                    else:
                        SELLO_FOTO = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold h-100 w-100'><div><small>"+str(SELLO_FOTO)+"</small></div></div>"
                else:
                    SELLO_FOTO = "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"
                
                
                
                if MACRO == "":
                    if "Exportacion" in Info_Caja.keys():
                        if Ya_Libero_Operaciones == True and Ya_Man_Doc_Impreso_WMS == True:
                            MACRO = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='MACRO("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                    else:
                        MACRO =  "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"
                else:
                    MACRO = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold h-100 w-100'><div><small><a class='link-light' style='cursor:pointer' onclick='MACRO("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")'>"+str(MACRO)+"</a></small></div></div>"

                if Ya_Libero_Operaciones == True:
                    if POD_Salida is None:
                        POD_Salida = "<div class='d-flex justify-content-center align-items-center bg-warning-subtle fw-bold h-100 w-100'><button onclick='POD_Salida(\""+str(Caja["cc_id"])+"\",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-warning p-0 m-0 ps-1 pe-1'><i class='mdi mdi-check'></i></button></div>"
                        SALIDA = ""
                    else:
                        POD_Salida = "<div class='d-flex justify-content-center align-items-center text-bg-success fw-bold h-100 w-100'><div><small>"+str(POD_Salida)+"</small></div></div>"
                        SALIDA = "<div class='d-flex justify-content-center align-items-center bg-dark-subtle fw-bold h-100 w-100'><button onclick='Imprimir_Pase("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")' class='btn btn-sm btn-dark p-0 m-0 ps-1 pe-1'><i class='mdi mdi-printer'></i></button></div>"
                else:
                    POD_Salida = ""

                if Fecha_Salida is None:
                    Fecha_Salida = "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"
                else:
                    if SALIDA_Fecha is None:
                        pass
                    else:
                        pass
            
            
                if "MilkRun" in Info_Caja.keys():
                    Man_Doc_Impreso_WMS = "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"
                    Correo_Listo = "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"
                    MACRO =  "<div class='d-flex justify-content-center align-items-center text-bg-dark fw-bold h-100 w-100'><div><small></small></div></div>"

            Cancelado = 0
            Estado = ""

            if Caja is not None:
                Salidas.append({
                    "FECHA": str(Info_Caja["Fecha_Salida"]),
                    "CUT_TIME":str(Fecha_Salida),
                    "TIPO":str(Info_Caja["Outbound"]),
                    "RUTA":str(Info_Caja[Info_Caja["Outbound"]]),
                    "BOL":str(Info_Caja["BOL"]),
                    "CAJA":str(Caja_N),
                    "CARRIER":str(Carrier),
                    "ENTRADA ILC": str(Entrada),
                    "ASIGNACION": str(Asignacion),
                    "OPERACIONES": str(Operaciones),
                    "SALIDA" : str(SALIDA),
                    "POD_Salida":str(POD_Salida)
                })
            else:
                pass

        Resultado["Contenido"] += """
        </div>
        """

        Rip_Abiertos = []
        Masters = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip RIP inner join ulc_saltillo.ccajas CAJA ON RIP.crip_idcaja = CAJA.cc_id  WHERE RIP.crip_estado = 1 AND RIP.crip_despacho = 1")
        IDs = []
        for M in Masters:
            IDs.append(str(M["crip_id"]))
        Partes = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip_partes WHERE cripp_master IN ("+",".join(IDs)+")")
        Historico = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip_historico WHERE criph_master IN ("+",".join(IDs)+")")
        for M in Masters:
            ASNs_Aqui = []
            Partes_Aqui = []
            Historico_Aqui =[]
            for P in Partes:
                if int(P["cripp_master"]) == int(M["crip_id"]):
                    if "<div>"+str(P["cripp_asn"])+"</div>" not in ASNs_Aqui:
                        ASNs_Aqui.append("<div>"+str(P["cripp_asn"])+"</div>")
                    if "<div>"+str(P["cripp_asn"])+"</div>" not in Partes_Aqui:
                        Partes_Aqui.append("<div>"+str(P["cripp_parte"])+"</div>")
            for H in Historico:
                if int(M["crip_id"]) == int(H["criph_master"]):
                    Historico_Aqui.append(H)
            Op = """
            <div class="btn-group btn-group-sm" role="group">
            <button class='btn btn-warning btn-sm p-0 ps-1 pe-1' onclick='Ver_Historico_Rip("""+str(M["crip_id"])+""",\""""+str(M["crip_id"]).zfill(5)+"""\");'><i class='mdi mdi-history'></i></button>
            """
            Op += """
            <button class='btn btn-primary btn-sm p-0 ps-1 pe-1' onclick='Modificar_Ruteo("Editar","""+str(M["crip_id"])+""",\""""+str(M["crip_id"]).zfill(5)+"""\");'><i class='mdi mdi-file-document-edit'></i></button>
            """            
            Op += """
            </div>"""
            Despacho = "<div class='text-bg-secondary'>N/A</div>"
            if int(M["crip_despacho"]) == 1:
                Despacho = "<div class='text-bg-warning'>Pendiente</div>"
            if int(M["crip_despacho"]) == 2:
                Despacho = "<div class='text-bg-success'>Solucionado</div>"
            
            Rip_Abiertos.append({'Folio':str(M["crip_id"]).zfill(5),"ASNs":str("".join(ASNs_Aqui)),"Partes":str("".join(Partes_Aqui)),"Fecha":str(M["crip_fecha_hora"]),"Despacho":str(Despacho),"Op":Op,"Caja": str(M["cc_contenedor"])})

        Resultado["Contenido"] += """
        <div class='container-fluid'>


        <h1>SALIDAS <span class='text-warning'>EN PROCESO</span></h1>
        <div id='Tabla_Salida' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Salida;
            var Tabla_Salida = new Tabulator("#Tabla_Salida", {
            height:800,
            layout:"fitColumns",
            data:"""+str(Salidas)+""",
            rowFormatter:function(row){
                var data = row.getData();                    
            },
            columns:[
                {field:"BOL","title":"BOL",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle","width":150},
                {field:"FECHA","title":"FECHA",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle","width":150},
                {field:"TIPO","title":"TIPO",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html","width":150},
                {field:"RUTA","title":"RUTA",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle","width":150},
                {field:"CAJA","title":"CAJA",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle","width":150},
                {field:"CARRIER","title":"CARRIER",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"ENTRADA ILC","title":"ENTRADA ILC",formatter:"html","width":50,headerVertical:true,headerSort:false,headerHozAlign:"center",vertAlign:"middle"},
                {field:"ASIGNACION","title":"ASIGNACION",formatter:"html","width":50,headerVertical:true,headerSort:false,headerHozAlign:"center",vertAlign:"middle"},
                {field:"OPERACIONES","title":"OPERACIONES",formatter:"html","width":50,headerVertical:true,headerSort:false,headerHozAlign:"center",vertAlign:"middle"},
                {field:"POD_Salida","title":"POD SALIDA",formatter:"html","width":50,headerVertical:true,headerSort:false,headerHozAlign:"center",vertAlign:"middle"},
                {field:"SALIDA","title":"SALIDA",formatter:"html","width":50,headerVertical:true,headerSort:false,headerHozAlign:"center",vertAlign:"middle"},
                {field:"CUT_TIME","title":"CUT TIME",headerHozAlign:"center",vertAlign:"middle",formatter:"html","width":50,headerVertical:true},
            ],
            });
        </script>
        <hr>
        """

        Resultado["Contenido"] += """
        <script>
            /*------------------------------------------------------------------------------------------------------------*/
            function Documentos_WMS_N(ID_Caja,Caja){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Documentos WMS ["+Caja+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Documentos_WMS_N","ID_Caja":ID_Caja};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Editar_CutTime(ID){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Editar CutTime");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Editar_CutTime","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Enviar_Correo_N(ID,Caja){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Enviar Correos de CAJA LISTA  de la caja ["+Caja+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg');
                var parametros = {"Fun":"Enviar_Correo_N","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Re_Enviar_Correo_N(ID,Caja){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Re-Enviar Correos de CAJA LISTA  de la caja ["+Caja+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg');
                var parametros = {"Fun":"Re_Enviar_Correo_N","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Sellado_Foto(ID,Caja){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Sello y Foto ["+Caja+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Sellado_Foto","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Imprimir_Pase(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                var parametros = {"Fun":"Imprimir_Pase","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        if(Resultado["Estado"] == 1){
                            if("Pase Salida" in Resultado)
                            {
                                var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                win1.focus();
                            }
                            if("Manifiesto" in Resultado)
                            {
                                var win1 = window.open(Resultado["Manifiesto"], '_blank');
                                win1.focus();
                            }
                            Mensaje(2);
                        }else{
                            Mensaje(0,Resultado["Contenido"]);
                        }
                        
                    },
                    error: function (jqXHR, textStatus, errorThrown ){Mensaje(0,textStatus);}
                });
            }
            function Asignar_Load_ID_N(ID,Fecha){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Asignar LOAD ID ["+Fecha+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Asignar_Load_ID_N","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Asignar_Caja_N(ID,Fecha){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Asignar Pre-Caja ["+Fecha+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Asignar_Caja_N","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Activar_Ruta(ID,Tipo,Ruta){
            Swal.fire({
            title: 'Estas seguro de Activar la Ruta ['+Tipo+' '+Ruta+']?',
            buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Yes",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
            customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
            preConfirm: () => {
                Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Activar_Ruta","ID":ID};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
                })
            }
            function MACRO(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Macro ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"MACRO","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Agregar_Adicional(){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Agregar adicional");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Agregar_Adicional"};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Cargar_Load_ID(){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Agregar adicional");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Cargar_Load_ID"};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Crear_Programacion(){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-plus'></i> Crear Programación de Día");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Crear_Programacion"};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }
            function Folio_Prisma_N(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Folio de PRISMA ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Folio_Prisma_N","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function POD_Salida(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> POD DE SALIDA ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"POD_Salida","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Folio_Prisma(Load_ID,Caja,ID){
                Swal.fire({
                title: 'Que folio de Prisma es para LOAD ID ['+Load_ID+'] y Caja ['+Caja+']?',
                input: 'text',buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Yes",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                preConfirm: (Comentario) => {
                    
                    if(Comentario.trim() == ""){
                        Mensaje(1,"Agrega comentario");
                    }else{
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Folio_Prisma","Load_ID":Load_ID,"Comentario":Comentario,"ID":ID};
                        $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response)
                            {
                                var Resultado = JSON.parse(response);
                                if(Resultado["Estado"] == 1)
                                {
                                    $("#Vent_1").modal("hide");
                                    Mensaje(2);
                                    Cargar_Fecha($("#fecha").val());
                                }
                                else
                                    Mensaje(0,Resultado["Contenido"]);
                                    
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Mensaje(0,textStatus);
                            }
                        });
                    }
                    
                    
                }
                })
            }
            function Cancelar_Ruta(ID,Tipo,Ruta){
                Swal.fire({
                title: 'Estas seguro de Cacelar la Ruta ['+Tipo+' '+Ruta+']?',
                input: 'text',buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Yes",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                preConfirm: (Comentario) => {
                    
                    if(Comentario.trim() == ""){
                        Mensaje(1,"Agrega comentario");
                    }else{
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Cancelar_Ruta","ID":ID,"Comentario":Comentario};
                        $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response)
                            {
                                var Resultado = JSON.parse(response);
                                if(Resultado["Estado"] == 1)
                                {
                                    $("#Vent_1").modal("hide");
                                    Mensaje(2);
                                    Cargar_Fecha($("#fecha").val());
                                }
                                else
                                    Mensaje(0,Resultado["Contenido"]);
                                    
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Mensaje(0,textStatus);
                            }
                        });
                    }
                    
                    
                }
                })
            }

            function Ver_Historico(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> History ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-xl');
                var parametros = {"Fun":"Ver_Historico","ID":ID};
                $.ajax({data:  parametros,url:"./Container_Control/Container.py",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Ver_Historico_Correo(ID,LoadID){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> History NO TRAILER ["+LoadID+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-xl');
                var parametros = {"Fun":"Ver_Historico_Correo","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        $("#Vent_1").modal("show").find(".modal-body").html(Resultado["Contenido"]);
                        $("#Vent_1").find(".modal-footer").find("button").attr('onclick',"$('#Vent_1').modal('hide'); delete table; ")
                        swal.close();
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        $("#Vent_1").modal("show").find(".modal-body").html("<i class='mdi mdi-alert'></i> "+ textStatus);
                        swal.close();
                    }
                });
            }

            function Ver_Historico_Rip(ID,Folio){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg')
                $("#Vent_1").modal("show").find(".modal-body").html("<div class='w-100 text-center'><div class='spinner-grow' role='status'></div><b> Loading...</b></div>").parent().find(".modal-title").html("<i class='mdi mdi-history'></i> History ["+Folio+"]");
                var parametros = {"Fun":"Ver_Historico_Rip","ID":ID};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    type:  "post",
                    success:  function (response){var Resultado = JSON.parse(response); $("#Vent_1").find(".modal-body").html(Resultado["Contenido"]);swal.close();},
                    error: function (jqXHR, textStatus, errorThrown ){Mensaje(0,'Process error ['+ jqXHR.status + " | " + textStatus  + " | " + errorThrown +']');}
                });
            }
            function Modificar_Ruteo(Accion,ID,Folio,De_Donde,Problemas){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg')
                $("#Vent_1").modal("show").find(".modal-body").html("<div class='w-100 text-center'><div class='spinner-grow' role='status'></div><b> Loading...</b></div>").parent().find(".modal-title").html("<i class='mdi mdi-file-document-edit'></i> Edit ["+Folio+"] "+Problemas);
                var parametros = {"Fun":"Modificar_Ruteo","Accion":Accion,"ID":ID,"De_Donde":De_Donde};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    type:  "post",
                    success:  function (response){var Resultado = JSON.parse(response); $("#Vent_1").find(".modal-body").html(Resultado["Contenido"]);swal.close();},
                    error: function (jqXHR, textStatus, errorThrown){$("#Vent_1").find(".modal-body").html("<span style='color:red;'><b>ERROR</b></span> <hr>"+textStatus);}
                });
            }

            function Enviar_Nuevo_Reporte(){
                
                var info = {};
                info["IDs"] = []
                Tabla_Salida_Pendientes.getSelectedData().forEach(element => {
                    info["IDs"].push(element.ID)
                });
                if(info["IDs"].length > 0){
                    Mostrar_Ventana_Cargando(false);
                    $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg')
                    $("#Vent_1").modal("show").find(".modal-body").html("<div class='w-100 text-center'><div class='spinner-grow' role='status'></div><b> Loading...</b></div>").parent().find(".modal-title").html("<i class='mdi mdi-history'></i> Reporte por falta de Caja");
                    var parametros = {"Fun":"Enviar_Nuevo_Reporte","Info":JSON.stringify(info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        type:  "post",
                        success:  function (response){var Resultado = JSON.parse(response); $("#Vent_1").find(".modal-body").html(Resultado["Contenido"]);swal.close();},
                        error: function (jqXHR, textStatus, errorThrown ){Mensaje(0,'Process error ['+ jqXHR.status + " | " + textStatus  + " | " + errorThrown +']');}
                    });
                }
            }

        </script>
        </div>
        """
        Resultado["Contenido"] += """
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


#-----------------------------------------------------------------------------------------------------
def Documentos_WMS_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "WMS" }
        Ruta = json.loads(DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID_Caja"])+"'")[0]["cc_informacion_actual"])
        Formulario["Campos"].append({"tipo":"archivo","campo":"WMS","titulo":"Parts Manifest (CSV) / WMS PRODUCTIVO","Requerido":1,"Col":12,"min":0,"max":1,"tipo_archivo":["application/vnd.ms-excel","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","text/csv"]})
        Formulario["Campos"].append({"tipo":"archivo","campo":"CMS","titulo":"Parts Manifest (CSV) / CMS VACIO","Requerido":1,"Col":12,"min":0,"max":1,"tipo_archivo":["application/vnd.ms-excel","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","text/csv"]})


        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Documentos_WMS_Guardar_N("""+str(Datos["ID_Caja"])+""")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Documentos_WMS_Guardar_N(ID_Caja){
                var Info = Dame_Formulario(".WMS",true);
                if(Info['WMS'] == "" && Info['CMS'] == "" ){
                    Info = null;
                }
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Documentos_WMS_Guardar_N","Info":JSON.stringify(Info),"ID_Caja":ID_Caja};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }else{
                    Swal.fire({icon: 'warning',position: 'top-end',title: 'Check required fields.',showConfirmButton: false,toast: true,background : "#ffeb96",timer: 1500,timerProgressBar: true});
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Documentos_WMS_Guardar_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info_Datos = json.loads(str(Datos["Info"]))
        
        if "WMS" in Info_Datos.keys() and len(Info_Datos["WMS"]) > 0:
            CSV_Contenido = None
            Archivo =  open("I:/Portal_File/"+str(Info_Datos["WMS"][0]), newline='', encoding="utf-8-sig")
            CSV_Contenido = csv.reader(Archivo)
            R = ""
            Aux = []
            Resultado = {}
            if CSV_Contenido is not None:
                for row in CSV_Contenido:
                    Aux.append(row)
            Info_Necesario = []
            index = 0
            for A in Aux:
                if index > 0:
                    if len(A) >= 26:
                        for i in range(0,26):
                            if "E+" in A[i] or "e+" in A[i]:
                                A[i] = format(Decimal(A[i]), 'f')  # ← quita E+

                        if str(A[21]).strip() == "":
                            A[21] = "0"
                        if str(A[22]).strip() == "":
                            A[22] = "0"
                        Aux_2 = {}
                        Aux_2["Trailer_Number"] = A[0]
                        Aux_2["Destination"] = A[1]
                        Aux_2["ASC_Number"] = A[2]
                        Aux_2["Bill_of_Lading"] = A[3]
                        Aux_2["Outbound_ID"] = A[4]
                        Aux_2["Bill_of_Lading_2"] = A[6]
                        Aux_2["Outbound_ID_2"] = int(str(A[7]).replace(",",""))
                        Aux_2["Packing_Slip"] = A[9]
                        Aux_2["Pro_Number"] = A[10]
                        Aux_2["Weight"] = float(str(A[12]).replace(",",""))
                        Aux_2["B_L"] = A[13]
                        Aux_2["Vendor"] = A[17]
                        Aux_2["PartOrContainer1"] = A[18]
                        Aux_2["ContainerQuantity"] = int(str(A[21]).replace(",",""))
                        Aux_2["PartQuantity"] = int(str(A[22]).replace(",",""))
                        Aux_2["Destination_1"] = A[23]
                        Aux_2["Total_Containers3"] = int(str(A[24]).replace(",",""))
                        Aux_2["Total_Parts2"] = int(str(A[25]).replace(",",""))
                        Aux_2["DataSet1"] = A[26]
                        Aux_2["Description_"] = A[19]
                        Info_Necesario.append(Aux_2)
                index += 1
            

            Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID_Caja"])+"'")[0]
            Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
            Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
            Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
            Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
            Informacion_Actual = json.loads(str(Info_Ahora["cc_informacion_actual"]))
            Informacion_Actual["Archivos WMS"] = Info_Datos["WMS"]
            Error = DB.Instruccion(""" 
            INSERT INTO ulc_saltillo.ccajas_moviemiento
            (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
            VALUES
            ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Informacion_Actual))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','DOCUMENTOS WMS','"""+str(Info_Ahora["cc_info_ruta"])+"""')
            """)
            ID_Moviemiento = DB.Get_Dato("SELECT MAX(cch_fecha_hora) AS Ultimo FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Info_Ahora["cc_id"])+"' AND cch_movimiento = 'DOCUMENTOS WMS'")[0]["Ultimo"]

            if Error == "":
                for P in Info_Necesario:
                    Error += DB.Instruccion("""
                    INSERT INTO ilc_saltillo.cpartes_enviandas
                    (cpe_trailer_number,cpe_destination,cpe_asc_number,cpe_bill_of_loading,cpe_outbounID,cpe_bill_of_loading_2,cpe_outboundID_2,cpe_packing_slip,cpe_pro_number,cpe_weight,cpe_b_l,cpe_vendor,cpe_partorcontainer1,cpe_containerquantity,cpe_partquantity,cpe_destination1,cpe_total_container,cpe_total_part2,cpe_dataset1,cpe_movieminto_master,cpe_movieminto_fecha_hora,cpe_descripcion)
                    VALUES
                    ('"""+str(P["Trailer_Number"])+"""','"""+str(P["Destination"])+"""','"""+str(P["ASC_Number"])+"""','"""+str(P["Bill_of_Lading"])+"""','"""+str(P["Outbound_ID"])+"""','"""+str(P["Bill_of_Lading_2"])+"""','"""+str(P["Outbound_ID_2"])+"""','"""+str(P["Packing_Slip"])+"""','"""+str(P["Pro_Number"])+"""','"""+str(P["Weight"])+"""','"""+str(P["B_L"])+"""','"""+str(P["Vendor"])+"""','"""+str(P["PartOrContainer1"])+"""','"""+str(P["ContainerQuantity"])+"""','"""+str(P["PartQuantity"])+"""','"""+str(P["Destination_1"])+"""','"""+str(P["Total_Containers3"])+"""','"""+str(P["Total_Parts2"])+"""','"""+str(P["DataSet1"])+"""','"""+str(Info_Ahora["cc_id"])+"""','"""+str(ID_Moviemiento)+"""','"""+str(P["Description_"])+"""')
                    """)
                    if Error != "":
                        break
                
                if Error != "":
                    DB.Instruccion("DELETE FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Info_Ahora["cc_id"])+"' AND cch_fecha_hora = '"+str(ID_Moviemiento)+"'")
                    DB.Instruccion("DELETE FROM ilc_saltillo.cpartes_enviandas WHERE cpe_movieminto_master = '"+str(Info_Ahora["cc_id"])+"' AND cpe_movieminto_fecha_hora = '"+str(ID_Moviemiento)+"'")
                else:
                    pass
            if Error == "":
                Resultado["Estado"] = 1
                Completo(Caja_ID=Datos["ID_Caja"])
                os.remove("I:/Portal_File/"+str(Info_Datos["WMS"][0]))
            else:
                Resultado["Contenido"] = Error
        if "CMS" in Info_Datos.keys() and len(Info_Datos["CMS"]) > 0:
            CSV_Contenido = None
            Archivo =  open("I:/Portal_File/"+str(Info_Datos["CMS"][0]), newline='', encoding="utf-8-sig")
            CSV_Contenido = csv.reader(Archivo)
            R = ""
            Aux = []
            Resultado = {}
            if CSV_Contenido is not None:
                for row in CSV_Contenido:
                    Aux.append(row)
            Info_Necesario = []
            index = 0
            for A in Aux:
                if index > 0:
                    if len(A) >= 26:

                        for i in range(0,26):
                            if "E+" in A[i] or "e+" in A[i]:
                                A[i] = format(Decimal(A[i]), 'f')  # ← quita E+
    
                        Aux_2 = {}
                        Aux_2["Copia"] = A[0]
                        Aux_2["Trailer_Number"] = A[2]
                        Aux_2["Carrier"] = A[4]
                        Aux_2["Proveedor_Codigo"] = A[10]
                        Aux_2["Proveedor"] = A[11]
                        Aux_2["Contenedor"] = A[16]
                        Aux_2["Planta"] = A[17]
                        Aux_2["Manifiesto"] = A[18]
                        Aux_2["PakingSlip"] = A[19]
                        Aux_2["Bol"] = A[20]
                        Aux_2["SID"] = A[21]
                        Aux_2["Qty"] = A[25]
                        Aux_2["WGT"] = A[26]
                        Info_Necesario.append(Aux_2)
                index += 1
            

            Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID_Caja"])+"'")[0]
            Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
            Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
            Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
            Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
            Informacion_Actual = json.loads(str(Info_Ahora["cc_informacion_actual"]))
            Informacion_Actual["Archivos CMS"] = Info_Datos["CMS"]
            Error = DB.Instruccion(""" 
            INSERT INTO ulc_saltillo.ccajas_moviemiento
            (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
            VALUES
            ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Informacion_Actual))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','DOCUMENTOS WMS','"""+str(Info_Ahora["cc_info_ruta"])+"""')
            """)
            ID_Moviemiento = DB.Get_Dato("SELECT MAX(cch_fecha_hora) AS Ultimo FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Info_Ahora["cc_id"])+"' AND cch_movimiento = 'DOCUMENTOS WMS'")[0]["Ultimo"]

            if Error == "":
                for P in Info_Necesario:
                    if str(P["Copia"]).strip() == "ILC":
                        Error += DB.Instruccion("""
                        INSERT INTO ilc_saltillo.cvacio_enviado
                        (cve_idcaja,cve_trailer,cve_carrier,cve_proveedor_codigo,cve_proveedor,cve_contenedor,cve_planta,cve_manifiesto,cve_pakingslip,cve_bool,cve_sid,cve_qty,cve_peso)
                        VALUES
                        ('"""+str(Datos["ID_Caja"])+"""','"""+str(P["Trailer_Number"])+"""','"""+str(P["Carrier"])+"""','"""+str(P["Proveedor_Codigo"])+"""','"""+str(P["Proveedor"])+"""','"""+str(P["Contenedor"])+"""','"""+str(P["Planta"])+"""','"""+str(P["Manifiesto"])+"""','"""+str(P["PakingSlip"])+"""','"""+str(P["Bol"])+"""','"""+str(P["SID"])+"""','"""+str(P["Qty"])+"""','"""+str(P["WGT"])+"""')
                        """)
                        if Error != "":
                            break
                
                if Error != "":
                    DB.Instruccion("DELETE FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Info_Ahora["cc_id"])+"' AND cch_fecha_hora = '"+str(ID_Moviemiento)+"'")
                    DB.Instruccion("DELETE FROM ilc_saltillo.cpartes_enviandas WHERE cpe_movieminto_master = '"+str(Info_Ahora["cc_id"])+"' AND cpe_movieminto_fecha_hora = '"+str(ID_Moviemiento)+"'")
                else:
                    pass
            if Error == "":
                Resultado["Estado"] = 1
                Completo(Caja_ID=Datos["ID_Caja"])
                os.remove("I:/Portal_File/"+str(Info_Datos["CMS"][0]))
            else:
                Resultado["Contenido"] = Error
        
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Editar_CutTime(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        
        Load_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"' ")[0]
        if Load_ID["clid_load_id"] is None:
            Load_ID["clid_load_id"] = ""
        Resultado["Contenido"] += """
        <div>Tipo: <span class='fw-bold'>"""+str(Load_ID["clid_ruta_tipo_2"])+"""</span></div>
        <div>Ruta: <span class='fw-bold'>"""+str(Load_ID["clid_ruta_tipo_ruta"])+"""</span></div>
        <div>Fecha de Ruta: <span class='fw-bold'>"""+str(Load_ID["clid_fecha"])+"""</span></div>
        <div>Load Id: <span class='fw-bold'>"""+str(Load_ID["clid_load_id"])+"""</span></div>
        <div>CutTime Actual: <span class='fw-bold'>"""+str(Load_ID["clid_cut_time"])+"""</span></div>
        <hr>
        """
        Formulario = {"Col":"12", "Campos": [],"Clase": "Editar_CutTime" }
        Formulario["Campos"].append({"tipo":"fecha-hora","campo":"Nuevo_CutTime","titulo":"Nuevo Cut Time","Requerido":1,"min":1,"max":150,"valor":str(Load_ID["clid_cut_time"])})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Editar_CutTime_Guardar("""+str(Datos["ID"])+""")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Editar_CutTime_Guardar(ID){
                var Info = Dame_Formulario(".Editar_CutTime",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Editar_CutTime_Guardar","Info":JSON.stringify(Info),"ID":ID};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Editar_CutTime_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info_Datos = json.loads(str(Datos["Info"]))
        Error += DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_cut_time = '"+str(Info_Datos["Nuevo_CutTime"])+"' WHERE clid_id = '"+str(Datos["ID"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


def Enviar_Correo_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Encri = LibDM_2023.Pass()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Corres_Pre = ""
        Programacion = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Confi_Ruta = json.loads(str(Programacion["clid_ruta_configuracion"]))
        if "DL" in Confi_Ruta.keys():
            Corres_Pre = Confi_Ruta["DL"]
        Carta_Porte = []
        Falta_Info = False
        for Parte in DB.Get_Dato("SELECT * FROM ilc_saltillo.cpartes_enviandas PARTES left join  ilc_saltillo.ccarta_porte_info PORTE on PARTES.cpe_partorcontainer1 = PORTE.cParte where PARTES.cpe_movieminto_master = "+str(Programacion["clid_id_caja"])+" and PARTES.cpe_activo = 1"):
            # if Parte["cCarta_Porte_Code"] is None:
            #     Falta_Info = True
            for K in Parte.keys():
                if Parte[K] is None:
                    Parte[K] = ""
            if str(Parte["CDescripcion"]) == "":
                Parte["CDescripcion"] = 'NO EXISTE EN EL CATALOGO'
            if str(Parte["cCarta_Porte_Code"]) == "":
                Parte["cCarta_Porte_Code"] = "01010101"
            if str(Parte["CCarta_Porte_Name"]) == "":
                Parte["CCarta_Porte_Name"] = "No existe en el catalógo"
            if str(Parte["cTAX_ID"]) == "":
                Parte["cTAX_ID"] = "XAXX010101000"
            if str(Parte["cTariff_Code"]) == "":
                Parte["cTariff_Code"] = "01010101"

            Carta_Porte.append({
                "Packing_Slip":str(Parte["cpe_packing_slip"]),
                "Part_Number":str(Parte["cpe_partorcontainer1"]),
                "Quantity":str(Parte["cpe_partquantity"]),
                "Description":str(Parte["CDescripcion"]),
                "Weight":str(Parte["cpe_weight"]),
                "UOM":"LBS",
                "DUNS":str(Parte["cpe_vendor"]),
                "Destination":str(Parte["cpe_destination"]),
                "Carta_Porte_Code":str(Parte["cCarta_Porte_Code"]),
                "Carta_Porte_Name":str(Parte["CCarta_Porte_Name"]),
                "TAX_ID":str(Parte["cTAX_ID"]),
                "Tariff_Code":str(Parte["cTariff_Code"]),
                "Clave_Unidad":"H87"
            })
        Resultado["Contenido"] += """
        <div class='h3'>CARTA PORTE</div>
        <div id='Carta_Porte' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Carta_Porte;
            var Carta_Porte = new Tabulator("#Carta_Porte", {
            height:400,
            layout:"fitColumns",
            data:"""+str(Carta_Porte)+""",
            rowFormatter:function(row){
                var data = row.getData();
                if(data.Carta_Porte_Code == ""){
                    row.getElement().style.backgroundColor = "#ff72727e";
                }
                
            },
            columns:[
                {field:"Packing_Slip","title":"Packing Slip / ASN",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"},
                {field:"Part_Number","title":"Part Number",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Quantity","title":"Quantity",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"},
                {field:"Description","title":"Description",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Weight","title":"Weight",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"UOM","title":"UOM",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"DUNS","title":"DUNS",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Destination","title":"Destination",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Carta_Porte_Code","title":"Carta Porte Code",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Carta_Porte_Name","title":"Carta Porte Name",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"TAX_ID","title":"TAX_ID",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Tariff_Code","title":"Tariff Code",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Clave_Unidad","title":"Clave Unidad",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
            ],
            });
        </script>
        <hr>
        """
        if Falta_Info == False:
            Resultado["Contenido"] += "<hr>"
            Formulario = {"Col":"12", "Campos": [],"Clase": "Correo" }
            Formulario["Campos"].append({"tipo":"texto","campo":"Manifiesto WMS","titulo":"Manifiesto WMS","Requerido":1,"min":1,"max":1500,"valor":""})
            Formulario["Campos"].append({"tipo":"lista","tipo_item":"email","titulo_nuevo_item":"Correo","campo":"Correo","titulo":"Correo","Requerido":1,"min":1,"max":50,"valor":Corres_Pre.split(";")})
            #Formulario["Campos"].append({"tipo":"multitexto","campo":"Correo","titulo":"Correo (Separado con ;)","Requerido":1,"min":1,"max":1500,"valor":Corres_Pre})
            Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
            Resultado["Contenido"] += """
            <div class='text-center row'>
                <div class='col-6'><button class='btn btn-success w-75' onclick='Enviar_Correo_Enviar_N(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Enviar</button></div>
                <div class='col-6'><button class='btn btn-primary w-75' onclick='Generar_Excel_N(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-file-excel'></i> Descargar Carta Porte</button></div>
            </div>
            <script>
                function Enviar_Correo_Enviar_N(ID){
                    var Info = Dame_Formulario(".Correo",true);
                    if(Info != null)
                    {
                        Info["ID"] = ID;
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Enviar_Correo_Enviar_N","Info":JSON.stringify(Info)};
                        $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response)
                            {
                                var Resultado = JSON.parse(response);
                                if(Resultado["Estado"] == 1)
                                {
                                    $("#Vent_1").modal("hide");
                                    Mensaje(2);
                                    Cargar_Fecha($("#fecha").val());
                                }
                                else
                                    Mensaje(0,Resultado["Contenido"]);
                                    
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Mensaje(0,textStatus);
                            }
                        });
                    }
                }
                function Generar_Excel_N(ID)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Generar_Excel_N","ID":ID};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                var win1 = window.open("http://10.4.7.219:8080/Portal_File/Manifest_SODC.xlsx", '_blank');
                                win1.focus();
                                Mensaje(2);
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            </script>
            """
        else:
            Resultado["Contenido"] += "<div class='h3 fw-bold text-danger text-center'>COMPLETA LA INFORMACIÓN DELA CARGA PORTE PARA PODER ENVIAR LE CORREO</div>"
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Enviar_Correo_Enviar_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])
        HTML = ""
        SCAC = ""
        RUTA = ""
        Variables = {}
        Variables["LoadID"] = str(Info["ID"])
        Variables["Manifiesto"] = str(Info["Manifiesto WMS"])

        Load_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Info["ID"])+"'")[0]
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Load_ID["clid_id_caja"])+"'")[0]
        Confi_Ruta = json.loads(str(Load_ID["clid_ruta_configuracion"]))
        if "SCAC" in Confi_Ruta.keys():
            SCAC = Confi_Ruta["SCAC"]
        RUTA = Load_ID["clid_ruta_tipo_ruta"]
        if Load_ID["clid_load_id"] is not None:
            HTML += "Load ID: "+str(Load_ID["clid_load_id"])
        Asunto = "SALTILLO ILC CAJA LISTA "+str(Info["Manifiesto WMS"])+"/TRANSP: "+str(SCAC)+" /: "+str(SCAC)+" "+str(Caja["cc_contenedor"])+" /DOCK: "+str(RUTA)
        Error = DB.Instruccion("INSERT INTO portal.ccorreos_enviar (cce_usuario,cce_asunto,cce_to,cce_contendio,cce_funcion,cce_variables) VALUES ('"+str(Datos["ID_User"])+"','"+str(Asunto)+"','"+str(",".join(Info["Correo"]))+"','"+str(HTML).replace("'","''")+"','Caja_Lista_ODC_N','"+str(json.dumps(Variables))+"')")
        if Error == "":
            Load_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Info["ID"])+"'")[0]
            DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_manifiesto_wms = '"+str(Info["Manifiesto WMS"])+"' WHERE clid_id = '"+str(Info["ID"])+"'")
            Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Load_ID["clid_id_caja"])+"'")[0]
            if Error == "":
                Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Caja["cc_id"])+"'")[0]
                Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
                Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
                Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
                Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
                Info_Nueva = json.loads(str(Info_Ahora["cc_informacion_actual"]))
                Info_Nueva["Correo Caja Lista"] = "OK"
                Info_Nueva["Correo Caja Lista DESTINOS"] = Info["Correo"]
                Error += DB.Instruccion(""" 
                INSERT INTO ulc_saltillo.ccajas_moviemiento
                (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
                VALUES
                ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Info_Nueva))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','CORREO CAJA LISTA ENVIADO','"""+str(Info_Ahora["cc_info_ruta"])+"""')
                """)
            if Error != "":
                Resultado["Estado"] = 0
                Resultado["Contenido"] = str(Error)
            else:
                Completo(CLoadID=Info["ID"])
        else:
            Resultado["Estado"] = 0
            Resultado["Contenido"] = Error
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Generar_Excel_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Error = ""
        SCAC = ""
        RUTA = ""
        Load_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Partes = DB.Get_Dato("SELECT * FROM ilc_saltillo.cpartes_enviandas PARTES inner join  ilc_saltillo.ccarta_porte_info PORTE on PARTES.cpe_partorcontainer1 = PORTE.cParte where PARTES.cpe_movieminto_master = "+str(Load_ID["clid_id_caja"])+" and PARTES.cpe_activo = 1")
        wb_obj = openpyxl.load_workbook("C:/inetpub/wwwroot/Portal2/ODC/Estado_Cajas/Manifiesto_Porte.xlsx")
        sheet_obj = wb_obj.worksheets[0]
        sheet_obj.cell(row = 3, column = 2).value = Partes[0]["cpe_trailer_number"]
        sheet_obj.cell(row = 3, column = 4).value = datetime.now()
        sheet_obj.cell(row = 3, column = 6).value = str(RUTA)
        Row = 5
        for Parte in Partes:
            for K in Parte.keys():
                if Parte[K] is None:
                    Parte[K] = ""
            sheet_obj.cell(row = Row, column = 1).value = str(Parte["cpe_packing_slip"])
            sheet_obj.cell(row = Row, column = 2).value = str(Parte["cpe_partorcontainer1"])
            sheet_obj.cell(row = Row, column = 3).value = str(Parte["cpe_partquantity"])
            sheet_obj.cell(row = Row, column = 4).value = str(Parte["CDescripcion"])
            sheet_obj.cell(row = Row, column = 5).value = str(Parte["cpe_weight"])
            sheet_obj.cell(row = Row, column = 6).value = "LBS"
            sheet_obj.cell(row = Row, column = 7).value = str(Parte["cpe_vendor"])
            sheet_obj.cell(row = Row, column = 8).value = str(Parte["cpe_destination"])
            sheet_obj.cell(row = Row, column = 9).value = str(Parte["cCarta_Porte_Code"])
            sheet_obj.cell(row = Row, column = 10).value = str(Parte["CCarta_Porte_Name"])
            sheet_obj.cell(row = Row, column = 11).value = str(Parte["cTAX_ID"])
            sheet_obj.cell(row = Row, column = 12).value = str(Parte["cTariff_Code"])
            sheet_obj.cell(row = Row, column = 13).value = "H87"
            Row += 1
        for col in sheet_obj.columns:
            max_length = 0
            column = col[2].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length)
            sheet_obj.column_dimensions[column].width = adjusted_width
        wb_obj.save("I:/Portal_File/Manifest_SODC.xlsx")
        wb_obj.close()
        if Error != "":
            Resultado["Estado"] = 0
            Resultado["Contenido"] = str(Error)
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Re_Enviar_Correo_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Encri = LibDM_2023.Pass()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Corres_Pre = ""
        Programacion = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Confi_Ruta = json.loads(str(Programacion["clid_ruta_configuracion"]))
        if "DL" in Confi_Ruta.keys():
            Corres_Pre = Confi_Ruta["DL"]
        Carta_Porte = []
        Re_Envios = []

        Movimientos_Aqui = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = "+str(Programacion["clid_id_caja"])+" ")
        for M in Movimientos_Aqui:
            if M["cch_movimiento"] == "CORREO CAJA LISTA ENVIADO":
                Info_Nueva = json.loads(str(M["cch_informacion_actual"]))
                Re_Envios.append({
                    "Fecha_Hora": str(M["cch_fecha_hora"]),
                    "Enviado_Por": str(Compartido_2023.Dame_Nombre(M["cch_usuario"])),
                    "Destinos": str(Info_Nueva["Correo Caja Lista DESTINOS"]),
                })
                break

        Falta_Info = False
        for Parte in DB.Get_Dato("SELECT * FROM ilc_saltillo.cpartes_enviandas PARTES inner join  ilc_saltillo.ccarta_porte_info PORTE on PARTES.cpe_partorcontainer1 = PORTE.cParte where PARTES.cpe_movieminto_master = "+str(Programacion["clid_id_caja"])+" and PARTES.cpe_activo = 1"):
            if Parte["cCarta_Porte_Code"] is None:
                Falta_Info = True
            for K in Parte.keys():
                if Parte[K] is None:
                    Parte[K] = ""
            Carta_Porte.append({
                "Packing_Slip":str(Parte["cpe_packing_slip"]),
                "Part_Number":str(Parte["cpe_partorcontainer1"]),
                "Quantity":str(Parte["cpe_partquantity"]),
                "Description":str(Parte["CDescripcion"]),
                "Weight":str(Parte["cpe_weight"]),
                "UOM":"LBS",
                "DUNS":str(Parte["cpe_vendor"]),
                "Destination":str(Parte["cpe_destination"]),
                "Carta_Porte_Code":str(Parte["cCarta_Porte_Code"]),
                "Carta_Porte_Name":str(Parte["CCarta_Porte_Name"]),
                "TAX_ID":str(Parte["cTAX_ID"]),
                "Tariff_Code":str(Parte["cTariff_Code"]),
                "Clave_Unidad":"H87"
            })
        Resultado["Contenido"] += """
        <div class='h3'>CARTA PORTE</div>
        <div id='Carta_Porte' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Carta_Porte;
            var Carta_Porte = new Tabulator("#Carta_Porte", {
            height:400,
            layout:"fitColumns",
            data:"""+str(Carta_Porte)+""",
            rowFormatter:function(row){
                var data = row.getData();
                if(data.Carta_Porte_Code == ""){
                    row.getElement().style.backgroundColor = "#ff72727e";
                }
                
            },
            columns:[
                {field:"Packing_Slip","title":"Packing Slip / ASN",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"},
                {field:"Part_Number","title":"Part Number",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Quantity","title":"Quantity",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"},
                {field:"Description","title":"Description",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Weight","title":"Weight",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"UOM","title":"UOM",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"DUNS","title":"DUNS",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Destination","title":"Destination",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Carta_Porte_Code","title":"Carta Porte Code",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Carta_Porte_Name","title":"Carta Porte Name",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"TAX_ID","title":"TAX_ID",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Tariff_Code","title":"Tariff Code",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Clave_Unidad","title":"Clave Unidad",headerFilter:"input",hozAlign:"center",formatter:"html",headerHozAlign:"center",vertAlign:"middle"},
            ],
            });
        </script>
        <hr>
        <div class='h3'>Enviados</div>
        <div id='Re_Envios' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Re_Envios;
            var Re_Envios = new Tabulator("#Re_Envios", {
            height:400,
            layout:"fitColumns",
            data:"""+str(Re_Envios)+""",
            rowFormatter:function(row){
                var data = row.getData();
                if(data.Carta_Porte_Code == ""){
                    row.getElement().style.backgroundColor = "#ff72727e";
                }
            },
            columns:[
                {field:"Fecha_Hora","title":"Fecha y Hora",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"},
                {field:"Enviado_Por","title":"Enviado_Por",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle"},
                {field:"Destinos","title":"Destinos",headerFilter:"input",headerHozAlign:"center",vertAlign:"middle",formatter:"html"}
            ],
            });
        </script>
        <hr>
        """
        if Falta_Info == False:
            Resultado["Contenido"] += "<hr>"
            Formulario = {"Col":"12", "Campos": [],"Clase": "Correo" }
            Formulario["Campos"].append({"tipo":"texto","campo":"Manifiesto WMS","titulo":"Manifiesto WMS","Requerido":1,"min":1,"max":1500,"valor":str(Programacion["clid_manifiesto_wms"])})
            Formulario["Campos"].append({"tipo":"lista","tipo_item":"email","titulo_nuevo_item":"Correo","campo":"Correo","titulo":"Correo","Requerido":1,"min":1,"max":50,"valor":Corres_Pre.split(";")})
            #Formulario["Campos"].append({"tipo":"multitexto","campo":"Correo","titulo":"Correo (Separado con ;)","Requerido":1,"min":1,"max":1500,"valor":Corres_Pre})
            Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
            Resultado["Contenido"] += """
            <div class='text-center row'>
                <div class='col-6'><button class='btn btn-success w-75' onclick='Enviar_Correo_Enviar_N(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Enviar</button></div>
                <div class='col-6'><button class='btn btn-primary w-75' onclick='Generar_Excel_N(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-file-excel'></i> Descargar Carta Porte</button></div>
            </div>
            <script>
                function Enviar_Correo_Enviar_N(ID){
                    var Info = Dame_Formulario(".Correo",true);
                    if(Info != null)
                    {
                        Info["ID"] = ID;
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Enviar_Correo_Enviar_N","Info":JSON.stringify(Info)};
                        $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response)
                            {
                                var Resultado = JSON.parse(response);
                                if(Resultado["Estado"] == 1)
                                {
                                    $("#Vent_1").modal("hide");
                                    Mensaje(2);
                                    Cargar_Fecha($("#fecha").val());
                                }
                                else
                                    Mensaje(0,Resultado["Contenido"]);
                                    
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Mensaje(0,textStatus);
                            }
                        });
                    }
                }
                function Generar_Excel_N(ID)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Generar_Excel_N","ID":ID};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                var win1 = window.open("http://10.4.7.219:8080/Portal_File/Manifest_SODC.xlsx", '_blank');
                                win1.focus();
                                Mensaje(2);
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            </script>
            """
        else:
            Resultado["Contenido"] += "<div class='h3 fw-bold text-danger text-center'>COMPLETA LA INFORMACIÓN DELA CARGA PORTE PARA PODER ENVIAR LE CORREO</div>"
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Sellado_Foto(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Seal" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Sello","titulo":"Seal","Requerido":1,"min":1,"max":50})
        Formulario["Campos"].append({"tipo":"archivo","campo":"Fotos","titulo":"Photos","Requerido":1,"Col":12,"min":1,"max":5,"tipo_archivo":["image/*"],"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Sellado_Foto_Guardar("""+str(Datos["ID"])+""")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Sellado_Foto_Guardar(ID){
                var Info = Dame_Formulario(".Seal",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Sellado_Foto_Guardar","Info":JSON.stringify(Info),"ID":ID};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
         Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Sellado_Foto_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info = json.loads(str(Datos["Info"]))
        Error = ""
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
        Info_Ahora = json.loads(str(Caja["cc_informacion_actual"]))
        Info_Ahora["Etapa"] = "Despacho"
        Info_Ahora["Fotos"] = Info["Fotos"]
        Info_Ahora["Sello Blanco"] = Info["Sello"]
        Codigo = str(hashlib.md5(str("LOGISTICS INSIGHT CORPORATION - PASE SALIDA - "+str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))+" - "+str(Datos["ID"])).encode()).hexdigest())
        Folio = ''.join(random.SystemRandom().choice(string.ascii_uppercase + string.digits) for _ in range(8))
        Folio += str(Datos["ID"])
        Error += DB.Instruccion("UPDATE ulc_saltillo.ccajas SET cc_informacion_actual = '"+str(json.dumps(Info_Ahora))+"',cc_qr_salida = '"+str(Codigo)+"', cc_folio_salida = '"+str(Folio)+"',cc_bloquear = 1 WHERE cc_id = '"+str(Datos["ID"])+"'")
        if Error == "":
            Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
            Caja["cc_dock"] = 'null' if Caja["cc_dock"] is None else "'"+str(Caja["cc_dock"])+"'"
            Caja["cc_tipo_actual"] = 'null' if Caja["cc_tipo_actual"] is None else "'"+str(Caja["cc_tipo_actual"])+"'"
            Caja["cc_zona"] = 'null' if Caja["cc_zona"] is None else "'"+str(Caja["cc_zona"])+"'"
            Caja["cc_negocio"] = 'null' if Caja["cc_negocio"] is None else "'"+str(Caja["cc_negocio"])+"'"
            Error += DB.Instruccion(""" 
            INSERT INTO ulc_saltillo.ccajas_moviemiento
            (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
            VALUES
            ('"""+str(Caja["cc_id"])+"""',NOW(),'"""+str(Caja["cc_contenedor"])+"""','"""+str(Caja["cc_ubicacion"])+"""','"""+str(Caja["cc_informacion_actual"])+"""',"""+str(Caja["cc_dock"])+""","""+str(Caja["cc_tipo_actual"])+""","""+str(Caja["cc_zona"])+""","""+str(Caja["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','SELLADA','"""+str(Caja["cc_info_ruta"])+"""')
            """)
        if Error == "":
            Completo(Caja_ID=Datos["ID"])
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Folio_Prisma(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = str(Datos)
        Error = DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_folio_prisma = '"+str(Datos["Comentario"])+"'  WHERE clid_id = '"+str(Datos["ID"])+"' ")
        ID_Caja = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]["clid_id_caja"]
        Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(ID_Caja)+"'")[0]
        Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
        Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
        Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
        Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
        Error = DB.Instruccion(""" 
        INSERT INTO ulc_saltillo.ccajas_moviemiento
        (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
        VALUES
        ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(Info_Ahora["cc_informacion_actual"])+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','FOLIO PRISMA','"""+str(Info_Ahora["cc_info_ruta"])+"""')
        """)                
        if Error == "":
            Completo(CLoadID = Datos["ID"])
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Imprimir_Pase(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        # Resultado["Pase Salida"] = Generar_Pase_Salida(Datos["ID"],"Pase_Salida_Aduana")
        # Pase_Salida_Ticket(Datos["ID"],Datos["ID_User"])
        
        Resultado["Pase Salida"] = "http://10.4.7.219:8080/Portal_File/"+str(Compartido_2023.Pase_Salida_Excel(Datos["ID"],"Pase_Salida_Estado_Cajas_ULC_Saltillo","ULC_Saltillo"))
        Compartido_2023.Pase_Salida_Ticket_2(IP_Zebra,Datos["ID"],Datos["ID_User"],"ULC_Saltillo")
        
        Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
        Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
        Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
        Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
        Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
        DB.Instruccion(""" 
        INSERT INTO ulc_saltillo.ccajas_moviemiento
        (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
        VALUES
        ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(Info_Ahora["cc_informacion_actual"])+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','IMPRIMIR PASE DE SALIDA','"""+str(Info_Ahora["cc_info_ruta"])+"""')
        """)

        Resultado["Estado"] = 1
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Generar_Pase_Salida(ID_Caja,Nombre_Archivo):
    DB = LibDM_2023.DataBase()
    try:
        Error = ""
        CAJA = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(ID_Caja)+"'")[0]

        if CAJA["cc_qr_salida"] is None:
            Codigo = str(hashlib.md5(str("LOGISTICS INSIGHT CORPORATION - PASE SALIDA - "+str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))+" - "+str(ID_Caja)).encode()).hexdigest())
            Folio = ''.join(random.SystemRandom().choice(string.ascii_uppercase + string.digits) for _ in range(8))
            Folio += str(ID_Caja)
            Error += DB.Instruccion("UPDATE ulc_saltillo.ccajas SET cc_qr_salida = '"+str(Codigo)+"', cc_folio_salida = '"+str(Folio)+"',cc_bloquear = 1 WHERE cc_id = '"+str(ID_Caja)+"'")

        CAJA = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(ID_Caja)+"'")[0]
        if CAJA["cc_qr_salida"] is not None:
            #Codigo = str(CAJA["cc_qr_salida"])
            Codigo = str(CAJA["cc_qr_salida"])+"|"+str(CAJA["cc_qr_entarda"])
            Folio = str(CAJA["cc_folio_salida"])
            Info_Cajas_Dtalle = json.loads(str(CAJA["cc_informacion_actual"]))
            buffered = BytesIO()
            img = qrcode.make(Codigo)
            f = open(str(os.path.dirname(os.path.realpath(__file__)))+"/"+str(Nombre_Archivo)+".png", "wb")
            img.save(f)
            f.close()
            wb_obj = openpyxl.load_workbook("C:/inetpub/wwwroot/Portal2/WTC/Plantilla_Pase_Salida.xlsx")
            sheet_obj = wb_obj.worksheets[0]

            Sello = ""
            try:
                K = None
                for Key in Info_Cajas_Dtalle.keys():
                    if "Sello" in str(Key):
                        K = str(Key)
                if K is not None:
                    Sello = str(Info_Cajas_Dtalle[K])
            except:
                pass

            sheet_obj.cell(row = 7, column = 3).value = str(CAJA["cc_contenedor"])[-4:]
            sheet_obj.cell(row = 11, column = 3).value = str(Sello)
            sheet_obj.cell(row = 15, column = 3).value = str(CAJA["cc_contenedor"])
            sheet_obj.cell(row = 15, column = 8).value = str(Info_Cajas_Dtalle["Carrier"])
            sheet_obj.cell(row = 16, column = 3).value = str(CAJA["cc_negocio"])

            sheet_obj.cell(row = 17, column = 3).value = str(CAJA["cc_ultimo_mov"])

            try:
                Ruta = ""
                Ruta = str(CAJA["cc_tipo_actual"])
                if CAJA["cc_tipo_actual"] in Info_Cajas_Dtalle.keys():
                    Ruta = Info_Cajas_Dtalle[CAJA["cc_tipo_actual"]]
                    if Info_Cajas_Dtalle[CAJA["cc_tipo_actual"]] in Info_Cajas_Dtalle.keys():
                        Ruta = Info_Cajas_Dtalle[Info_Cajas_Dtalle[CAJA["cc_tipo_actual"]]]
                Tipo = Tipo[""]
            except:
                pass
            
            sheet_obj.cell(row = 16, column = 8).value = str(Ruta)

            Manifiesto = DB.Get_Dato("SELECT * FROM xdock.cmanifest_master WHERE cmam_caja = '"+str(ID_Caja)+"'")
            if len(Manifiesto) > 0:
                Manifiesto = Manifiesto[0]["cmam_manifest"]
            else:
                Manifiesto = ""
            sheet_obj.cell(row = 17, column = 8).value = str(Manifiesto)
            
            # sheet_obj.cell(row = 17, column = 3).value = datetime.now().strftime("%Y-%m-%d %H:%M")
            # MF = DB.Get_Dato("SELECT * FROM lc_slp.cmanifest_master WHERE cmam_caja = '"+str(ID_Caja)+"' ")
            # if len(MF) > 0:
            #     sheet_obj.cell(row = 17, column = 8).value = str(MF[0]["cmam_manifest"])
            # else:
            #     sheet_obj.cell(row = 17, column = 8).value = ""


            sheet_obj.cell(row = 36, column = 3).value = str(Folio)
            img = openpyxl.drawing.image.Image(str(os.path.dirname(os.path.realpath(__file__)))+'/'+str(Nombre_Archivo)+'.png')
            img.height = 310
            img.width = 310
            img.anchor = 'C20'
            sheet_obj.add_image(img)
            sheet_obj.protection.sheet = True
            sheet_obj.protection.password = 'LC_2022'
            sheet_obj.protection.formatColumns = False
            wb_obj.security  = WorkbookProtection(workbookPassword = 'LC_2022', lockStructure = True)
            wb_obj.save("C:/inetpub/wwwroot/Portal2/ODC/"+str(Nombre_Archivo)+".xlsx")
            return str(Nombre_Archivo)+".xlsx"
    except:
        return "ERROR: " + str(sys.exc_info())

def Asignar_Load_ID_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Aqui = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Formulario = {"Col":"12", "Campos": [],"Clase": "Asignar" }
        Opciones = []
        Resultado["Contenido"] += "<div class='text-center h1'>"+str(Info_Aqui["clid_ruta_tipo_2"])+"-"+str(Info_Aqui["clid_ruta_tipo_ruta"])+"</div>"
        Formulario["Campos"].append({"tipo":"texto","campo":"Load ID","titulo":"Load ID","Requerido":1,"min":1,"max":150,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Asignar_Load_ID_N_Guardar(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Asignar_Load_ID_N_Guardar(ID){
                var Info = Dame_Formulario(".Asignar",true);
                if(Info != null)
                {
                    Info["ID"] = ID;
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Asignar_Load_ID_N_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Asignar_Load_ID_N_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])
        Error += DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_load_id = '"+str(Info["Load ID"])+"', clid_usuario_asignacion = '"+str(Datos["ID_User"])+"', clid_fecha_hora_asignacion = NOW() WHERE  clid_id = '"+str(Info["ID"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Asignar_Caja_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Aqui = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Formulario = {"Col":"12", "Campos": [],"Clase": "Asignar" }
        Opciones = []
        Resultado["Contenido"] += "<div class='text-center h1'>"+str(Info_Aqui["clid_ruta_tipo_2"])+"-"+str(Info_Aqui["clid_ruta_tipo_ruta"])+"</div>"
        Formulario["Campos"].append({"tipo":"texto","campo":"Caja","titulo":"Caja","Requerido":1,"min":1,"max":150,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Asignar_Caja_N_Guardar(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Asignar_Caja_N_Guardar(ID){
                var Info = Dame_Formulario(".Asignar",true);
                if(Info != null)
                {
                    Info["ID"] = ID;
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Asignar_Caja_N_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Asignar_Caja_N_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])
        Error += DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_caja_pre_asignada = '"+str(Info["Caja"])+"' WHERE  clid_id = '"+str(Info["ID"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def MACRO(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Movimientos = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Datos["ID"])+"' ORDER BY cch_fecha_hora")

        Info_Macro = None
        for M in Movimientos:
            if M["cch_movimiento"] == "MACRO":
                Info_Macro = json.loads(M["cch_informacion_actual"])
        
        Info_Documentos = None
        for M in Movimientos:
            if M["cch_movimiento"] == "DOCUMENTOS WMS":
                Info_Documentos = json.loads(M["cch_informacion_actual"])

        


        Resultado["Contenido"] += """
        <div>
            <button class='btn btn-dark w-100' onclick='Descargar_Macro("""+str(Datos["ID"])+""")' ><i class='mdi mdi-cloud-download '></i> Descargar Excel</button>
        </div>
        <hr>
        """
        if Info_Macro is None:
            if Info_Documentos is not None:
                Formulario = {"Col":"6", "Campos": [],"Clase": "Mgo_Actualizar" }
                Formulario["Campos"].append({"tipo":"archivo","campo":"Excel","titulo":"Excel Final","Requerido":1,"Col":12,"min":1,"max":1,"tipo_archivo":["application/vnd.ms-excel","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]})
                if "Archivos WMS" in Info_Documentos.keys():
                    Formulario["Campos"].append({"tipo":"archivo","campo":"Fotos","titulo":"Captura Champs","Requerido":1,"Col":12,"min":1,"max":5,"tipo_archivo":["image/*"],"valor":""})
                Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
                Resultado["Contenido"] += """
                <hr>
                <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='MACRO_Guardar(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Guardar</button></div>
                <script>

                    function Descargar_Macro(ID)
                    {
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Descargar_Macro","ID":ID};
                        $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response)
                            {
                                var Resultado = JSON.parse(response);
                                if(Resultado["Estado"] == 1)
                                {
                                    var win1 = window.open("http://10.4.7.219:8080/Portal_File/Parts_ILC.xlsx", '_blank');
                                    win1.focus();
                                    Mensaje(2);
                                }
                                else
                                    Mensaje(0,Resultado["Contenido"]);
                                    
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Mensaje(0,textStatus);
                            }
                        });
                    }


                    function MACRO_Guardar(ID){
                        var Info = Dame_Formulario(".Mgo_Actualizar",true);
                        if(Info != null)
                        {
                            Info["ID"] = ID;
                            Mostrar_Ventana_Cargando(false);
                            var parametros = {"Fun":"MACRO_Guardar","Info":JSON.stringify(Info)};
                            $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                                success:  function (response)
                                {
                                    var Resultado = JSON.parse(response);
                                    if(Resultado["Estado"] == 1)
                                    {
                                        if("Pase Salida" in Resultado)
                                        {
                                            var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                            win1.focus();
                                        }
                                        $("#Vent_1").modal("hide");
                                        Mensaje(2);
                                        Cargar_Fecha($("#fecha").val());
                                    }
                                    else
                                        Mensaje(0,Resultado["Contenido"]);
                                        
                                },
                                error: function (jqXHR, textStatus, errorThrown )
                                {
                                    Mensaje(0,textStatus);
                                }
                            });
                        }
                    }
                </script>
                """
            else:
                Resultado["Contenido"] += """<h1>Documento WMS/CSM no encontrados</h1>"""
        else:
            Resultado["Contenido"] += """
            <div>
                <a class='btn btn-success w-100' href='http://10.4.7.219:8080/Portal_File/"""+str(Info_Macro["Macro"])+"""' ><i class='mdi mdi-cloud-download '></i> Descargar Excel (Final)</a>
            </div>
            <hr>
            """
            if "Champs" in Info_Macro.keys():
                Resultado["Contenido"] += """<h3>Champs</h3>"""
                for Champs in Info_Macro["Champs"]:
                    Resultado["Contenido"] += """
                    <div class="text-center">
                        <img src="http://10.4.7.219:8080/Portal_File/"""+str(Champs)+"""" class="rounded" alt="..." width="auto" height="200">
                    </div>
                    """

    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def MACRO_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])

        # for k in Info.keys():
        #     if "Pallets_" in str(k):
        #         id = str(k).split("_")[1]
        #         DB.Instruccion("UPDATE ilc_saltillo.cpartes_enviandas SET cpe_pallets = '"+str(Info[k])+"' WHERE cpe_id IN ("+str(id)+")")

        Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Info["ID"])+"'")[0]
        Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
        Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
        Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
        Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
        Info_Nueva = json.loads(str(Info_Ahora["cc_informacion_actual"]))
        Info_Nueva["Macro"] = Info["Excel"][0]
        if "Fotos" in Info.keys():
            Info_Nueva["Champs"] = Info["Fotos"]
        Error += DB.Instruccion(""" 
        INSERT INTO ulc_saltillo.ccajas_moviemiento
        (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
        VALUES
        ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Info_Nueva))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','MACRO','"""+str(Info_Ahora["cc_info_ruta"])+"""')
        """)
        
        if Error != "":
            Resultado["Estado"] = 0
            Resultado["Contenido"] = str(Error)
        else:
            Completo(CLoadID = Info["ID"])
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Descargar_Macro(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Error = ""
        SCAC = ""
        RUTA = ""
        Proveedores = DB.Get_Dato("SELECT * FROM ilc_saltillo.cproveedores WHERE crilcpr_activo = 1")
        Partes = DB.Get_Dato("SELECT * FROM ilc_saltillo.cpartes_enviandas where cpe_movieminto_master = "+str(Datos["ID"])+" and cpe_activo = 1 AND cpe_descripcion != ''")
        VacioS = DB.Get_Dato("SELECT * FROM ilc_saltillo.cvacio_enviado WHERE cve_idcaja = '"+str(Datos["ID"])+"' ")
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
    
        
        wb_obj = openpyxl.load_workbook("C:/inetpub/wwwroot/Portal2/ULC_ILC_Saltillo/Estado_Cajas/Macro.xlsx")
        sheet_obj = wb_obj.worksheets[0]
        sheet_obj.cell(row = 26, column = 3).value = Caja["cc_contenedor"]
        sheet_obj.cell(row = 20, column = 12).value = datetime.now().strftime("%d-%b-%y")
        sheet_obj.cell(row = 21, column = 12).value = datetime.now().strftime("%H:%M")

        Num = 1
        Row = 32
        for Parte in Partes:
            for K in Parte.keys():
                if Parte[K] is None:
                    Parte[K] = ""
            Proveedor = ""
            for Pro in Proveedores:
                if str(Pro["crilcpr_codigo"]) == str(Parte["cpe_vendor"]):
                    Proveedor = Pro["crilcpr_nombre"]
                    break
            sheet_obj.cell(row = Row, column = 1).value = str(Num)
            sheet_obj.cell(row = Row, column = 2).value = str(Parte["cpe_packing_slip"])
            sheet_obj.cell(row = Row, column = 3).value = str(Parte["cpe_packing_slip"])
            sheet_obj.cell(row = Row, column = 4).value = str(Proveedor)
            sheet_obj.cell(row = Row, column = 5).value = str(Parte["cpe_vendor"])
            sheet_obj.cell(row = Row, column = 6).value = str(Parte["cpe_partquantity"])
            sheet_obj.cell(row = Row, column = 7).value = str(Parte["cpe_destination1"])
            sheet_obj.cell(row = Row, column = 8).value = str(Parte["cpe_partorcontainer1"])
            sheet_obj.cell(row = Row, column = 9).value = str(Parte["cpe_descripcion"])
            sheet_obj.cell(row = Row, column = 10).value = str(Parte["cpe_asc_number"])
            sheet_obj.cell(row = Row, column = 11).value = ""
            sheet_obj.cell(row = Row, column = 12).value = ""
            sheet_obj.cell(row = Row, column = 13).value = ""
            sheet_obj.cell(row = Row, column = 14).value = ""
            Row += 1
            Num += 1
        for Vacio in VacioS:
            for K in Vacio.keys():
                if Vacio[K] is None:
                    Vacio[K] = ""
            Proveedor = ""
            for Pro in Proveedores:
                if str(Pro["crilcpr_codigo"]) == str(Vacio["cve_proveedor_codigo"]):
                    Proveedor = Pro["crilcpr_nombre"]
                    break
            sheet_obj.cell(row = Row, column = 1).value = str(Num)
            sheet_obj.cell(row = Row, column = 2).value = str(Vacio["cve_pakingslip"])
            sheet_obj.cell(row = Row, column = 3).value = str(Vacio["cve_pakingslip"])
            sheet_obj.cell(row = Row, column = 4).value = str(Vacio["cve_proveedor"])
            sheet_obj.cell(row = Row, column = 5).value = str(Vacio["cve_proveedor_codigo"])
            sheet_obj.cell(row = Row, column = 6).value = str(Vacio["cve_qty"])
            sheet_obj.cell(row = Row, column = 7).value = str(Vacio["cve_planta"])
            sheet_obj.cell(row = Row, column = 8).value = str(Vacio["cve_contenedor"])
            sheet_obj.cell(row = Row, column = 9).value = ""
            sheet_obj.cell(row = Row, column = 10).value = ""
            sheet_obj.cell(row = Row, column = 11).value = ""
            sheet_obj.cell(row = Row, column = 12).value = ""
            sheet_obj.cell(row = Row, column = 13).value = ""
            sheet_obj.cell(row = Row, column = 14).value = ""
            Row += 1
            Num += 1


        col = 2  # columna A
        fila_inicio = 32  # ignoramos encabezado

        alineacion = Alignment(vertical="center")
        
        valor_actual = sheet_obj.cell(row=fila_inicio, column=col).value
        inicio_merge = fila_inicio
        rangos_iguales = []
        for fila in range(fila_inicio + 1, sheet_obj.max_row + 1):
            valor = sheet_obj.cell(row=fila, column=col).value

            if valor != valor_actual:
                if inicio_merge < fila - 1:
                    sheet_obj.merge_cells(start_row=inicio_merge, start_column=col,end_row=fila - 1, end_column=col)
                    sheet_obj.cell(row=inicio_merge, column=col).alignment = alineacion

                    for C in [11,12,13,14]:
                        sheet_obj.merge_cells(start_row=inicio_merge, start_column=C,end_row=fila - 1, end_column=C)
                        sheet_obj.cell(row=inicio_merge, column=C).alignment = alineacion
                    
                    rangos_iguales.append((valor_actual, inicio_merge, fila - 1))

                valor_actual = valor
                inicio_merge = fila

        # Último grupo
        if inicio_merge < sheet_obj.max_row:
            sheet_obj.merge_cells(start_row=inicio_merge, start_column=col,end_row=sheet_obj.max_row, end_column=col)
            sheet_obj.cell(row=inicio_merge, column=col).alignment = alineacion

            for C in [11,12,13,14]:
                sheet_obj.merge_cells(start_row=inicio_merge, start_column=C,end_row=sheet_obj.max_row, end_column=C)
                sheet_obj.cell(row=inicio_merge, column=C).alignment = alineacion

            
            rangos_iguales.append((valor_actual, inicio_merge, sheet_obj.max_row))

        wb_obj.save("I:/Portal_File/Parts_ILC.xlsx")
        wb_obj.close()
        if Error != "":
            Resultado["Estado"] = 0
            Resultado["Contenido"] = str(Error)
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)



def Folio_Prisma_N(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Info_Aqui = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Datos["ID"])+"'")[0]
        Formulario = {"Col":"12", "Campos": [],"Clase": "Mgo_Actualizar" }
        Resultado["Contenido"] += "<div class='text-center h3'>"+str(Info_Aqui["clid_ruta_tipo_2"])+"-"+str(Info_Aqui["clid_ruta_tipo_ruta"])+"</div>"
        Formulario["Campos"].append({"tipo":"texto","campo":"Folio","titulo":"Folio","Requerido":1,"min":1,"max":50})
        Formulario["Campos"].append({"tipo":"archivo","campo":"Fotos","titulo":"Photos","Requerido":1,"Col":12,"min":1,"max":10,"tipo_archivo":["image/*"],"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Folio_Prisma_N_Guardar(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Folio_Prisma_N_Guardar(ID){
                var Info = Dame_Formulario(".Mgo_Actualizar",true);
                if(Info != null)
                {
                    Info["ID"] = ID;
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Folio_Prisma_N_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Folio_Prisma_N_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Info = json.loads(Datos["Info"])
        Error = ""
        Load_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(Info["ID"])+"'")[0]
        Error = DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_folio_prisma = '"+str(Info["Folio"])+"'  WHERE clid_id = '"+str(Info["ID"])+"' ")
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Load_ID["clid_id_caja"])+"'")[0]
        Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Caja["cc_id"])+"'")[0]
        Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
        Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
        Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
        Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
        Info_Nueva = json.loads(str(Info_Ahora["cc_informacion_actual"]))
        Info_Nueva["Fotos"] = Info["Fotos"]
        Info_Nueva["Folio Prisma"] = Info["Folio"]
        Error += DB.Instruccion(""" 
        INSERT INTO ulc_saltillo.ccajas_moviemiento
        (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
        VALUES
        ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Info_Nueva))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','FOLIO PRISMA','"""+str(Info_Ahora["cc_info_ruta"])+"""')
        """)
        if Error != "":
            Resultado["Estado"] = 0
            Resultado["Contenido"] = str(Error)
        else:
            Completo(CLoadID = Info["ID"])
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


def Activar_Ruta(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_activo = 1,clid_id_usuario_cancelar = null,clid_fecha_hora_cancelacion = null,clid_comentario_cancelacion = null WHERE clid_id = '"+str(Datos["ID"])+"' ")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Completo(CLoadID = None,Caja_ID = None):
    Error = ""
    try:
        DB = LibDM_2023.DataBase()
        Compartido_2023 = LibDM_2023.Compartido()
        # if CLoadID is not None and Caja_ID is None :
        #     Caja_ID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(CLoadID)+"' AND clid_activo = 1")[0]["clid_id_caja"]
        # if CLoadID is None and Caja_ID is not None:
        #     CLoadID = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id_caja = '"+str(Caja_ID)+"' AND clid_activo = 1")[0]["clid_id"]

        Movimientos = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas_moviemiento WHERE cch_master = '"+str(Caja_ID)+"' ORDER BY cch_fecha_hora")
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas  WHERE cc_id = '"+str(Caja_ID)+"' ")[0]
        Info_Ahora = json.loads(Caja["cc_informacion_actual"])

        Asignacion = None
        Operaciones = None
        Man_Doc_Impreso_WMS = None
        Correo_Listo = None
        SELLO_FOTO = None
        Actualziar_MGO = None
        Folio_Prisma = None
        SALIDA = None
        POD_Salida = None
        
        for M in Movimientos:
            if M["cch_movimiento"] == "ASIGNAR":
                Asignacion = M["cch_fecha_hora"]
        for M in Movimientos:
            if M["cch_movimiento"] == "LIBERA OPERACIONES":
                Operaciones = M["cch_fecha_hora"]
        for M in Movimientos:
            if M["cch_movimiento"] == "DOCUMENTOS WMS":
                Man_Doc_Impreso_WMS = M["cch_fecha_hora"]
        for M in Movimientos:
            if M["cch_movimiento"] == "CORREO CAJA LISTA ENVIADO":
                Correo_Listo = M["cch_fecha_hora"]
        for M in Movimientos:
            if "Exportacion" in Info_Ahora.keys():
                if M["cch_movimiento"] == "SELLADA":
                    SELLO_FOTO = M["cch_fecha_hora"]
        for M in Movimientos:
            if M["cch_movimiento"] == "MGO ACTUALIZADO":
                Actualziar_MGO = M["cch_fecha_hora"]
        for M in Movimientos:
            if "Exportacion" in Info_Ahora.keys():
                if M["cch_movimiento"] == "FOLIO PRISMA":
                    Folio_Prisma = M["cch_fecha_hora"]
        if "Exportacion" in Info_Ahora.keys():
            if Folio_Prisma is None:
                Escrito = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id_caja = '"+str(Caja_ID)+"'")
                if len(Escrito) > 0:
                    if Escrito[0]["clid_folio_prisma"] is not None:
                        Folio_Prisma = datetime.now()
        for M in Movimientos:
            if M["cch_movimiento"] == "POD Salida":
                POD_Salida = M["cch_fecha_hora"]
        
        
        for M in Movimientos:
            if M["cch_movimiento"] == "SALIDA":
                SALIDA = M["cch_fecha_hora"]
            if M["cch_movimiento"] == "ELIMINADO POR YARD":
                SALIDA = M["cch_fecha_hora"]
        if Info_Ahora["Outbound"] == "Exportacion":
            if Asignacion is not None and Operaciones is not None and Man_Doc_Impreso_WMS is not None and Correo_Listo is not None and Actualziar_MGO is not None and SALIDA is not None and SELLO_FOTO is not None and Folio_Prisma is not None and POD_Salida is not None:
                Error += DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_completos = 1 WHERE clid_id = '"+str(CLoadID)+"'")
        else:
            if Asignacion is not None and Operaciones is not None and Man_Doc_Impreso_WMS is not None and Correo_Listo is not None and Actualziar_MGO is not None and SALIDA is not None and POD_Salida is not None:
                Error += DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_completos = 1 WHERE clid_id = '"+str(CLoadID)+"'")
        
        Error += str(Info_Ahora["Outbound"]) + " ["+str(Caja_ID)+"] - " +str(Asignacion)+"/"+str(Operaciones)+"/"+str(Man_Doc_Impreso_WMS)+"/"+str(Correo_Listo)+"/"+str(Actualziar_MGO)+"/"+str(SALIDA)+"/"+str(SELLO_FOTO)+"/"+str(Folio_Prisma)
    except:
        Error += str(sys.exc_info())
    return Error


def Agregar_Adicional(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Adicional" }
        Rutas_Outbond = json.loads(DB.Get_Dato("SELECT * FROM ilc_saltillo.crutas WHERE cr_tipo = 'Outbound'")[0]["cr_niveles"])
        Opciones = []
        for K in Rutas_Outbond.keys():
            for Ruta in Rutas_Outbond[K].keys():
                if K != "MilkRun":
                    Opciones.append(str(K) + "-" + str(Ruta))
        Formulario["Campos"].append({"tipo":"fecha","id":"fecha","campo":"Fecha","titulo":"Fecha","Requerido":1,"Col":12,"min":1,"max":10,"valor":datetime.now().strftime("%Y-%m-%d")})
        Formulario["Campos"].append({"tipo":"seleccion","id":"Ruta","campo":"Ruta","titulo":"Ruta","Requerido":1,"Tipo_Opciones":"Opciones","Opciones":Opciones,"valor":"","Col":12})
        Formulario["Campos"].append({"tipo":"texto","id":"Load_ID","campo":"Load ID","titulo":"Load ID","Requerido":1,"min":1,"max":150,"valor":""})
        Formulario["Campos"].append({"tipo":"fecha-hora","id":"Cut_Time","campo":"Cut Time","titulo":"Cut Time","Requerido":1,"min":1,"max":150,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","id":"Caja","campo":"Caja","titulo":"Caja","Requerido":1,"min":1,"max":150,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Agregar_Adicional_Guardar()'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            $("#Ruta").on( "change", function() {
                if($(this).val().includes("Nacional")){
                    $("#Load_ID").removeClass('Adicional');
                    $("#Load_ID").parent().hide();
                    $("#Cut_Time").removeClass('Adicional');
                    $("#Cut_Time").parent().hide();
                    $("#Caja").removeClass('Adicional');
                    $("#Caja").parent().hide();
                }
                else{
                    $("#Load_ID").addClass('Adicional');
                    $("#Load_ID").parent().show();
                    $("#Cut_Time").addClass('Adicional');
                    $("#Cut_Time").parent().show();
                    $("#Caja").addClass('Adicional');
                    $("#Caja").parent().show();
                }
            } );
            function Agregar_Adicional_Guardar(){
                var Info = Dame_Formulario(".Adicional",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Agregar_Adicional_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Agregar_Adicional_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])
        #Error += str(Info)
        Rutas = DB.Get_Dato("SELECT * FROM ilc_saltillo.crutas WHERE cr_tipo = 'Outbound'")
        Rutas = json.loads(str(Rutas[0]["cr_niveles"]))
        Tipo = str(Info["Ruta"]).split("-")[0]
        Ruta = str(Info["Ruta"]).replace(str(Tipo)+"-","")
        Aux_Ya = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_ruta_tipo_2 = '"+str(Tipo)+"' AND clid_fecha = '"+str(Info["Fecha"])+"' AND clid_ruta_tipo_ruta LIKE '%"+str(Ruta)+"/A%' AND clid_activo = 1")

        Cut_Time = "null"
        if str(Rutas[Tipo][Ruta]["CUTTIME_OUT"]) != "":
            Cut_Time = datetime.strptime(str(Info["Fecha"])+" "+str(Rutas[Tipo][Ruta]["CUTTIME_OUT"])+":00","%Y-%m-%d %H:%M:%S")
            if str(Rutas[Tipo][Ruta]["CUTTIME_OUT"]) != "":
                Cut_Time = Cut_Time + timedelta(days=int(Rutas[Tipo][Ruta]["CUTTIME_OUT_DIA"]))
            Cut_Time = "'"+str(Cut_Time.strftime("%Y-%m-%d %H:%M:%S"))+"'"

        if "Load ID" in Info.keys():
            Error += DB.Instruccion("INSERT INTO ilc_saltillo.cload_id (clid_load_id,clid_fecha,clid_usuario_asignacion,clid_fecha_hora_asignacion,clid_ruta_tipo_1,clid_ruta_tipo_2,clid_ruta_tipo_ruta,clid_caja_pre_asignada,clid_cut_time,clid_ruta_configuracion,clid_usuario_cuendo_crea) VALUES ('"+str(Info["Load ID"])+"','"+str(Info["Fecha"])+"','"+str(Datos["ID_User"])+"',NOW(),'Outbound','"+str(Tipo)+"','"+str(Ruta)+"/A"+str(len(Aux_Ya)+1)+"','"+str(Info["Caja"])+"',"+str(Cut_Time)+",'"+str(json.dumps(Rutas[Tipo][Ruta]))+"',NOW())")
        else:
            Error += DB.Instruccion("INSERT INTO ilc_saltillo.cload_id (clid_fecha,clid_usuario_asignacion,clid_fecha_hora_asignacion,clid_ruta_tipo_1,clid_ruta_tipo_2,clid_ruta_tipo_ruta,clid_ruta_configuracion,clid_cut_time,clid_usuario_cuendo_crea) VALUES ('"+str(Info["Fecha"])+"','"+str(Datos["ID_User"])+"',NOW(),'Outbound','"+str(Tipo)+"','"+str(Ruta)+"/A"+str(len(Aux_Ya)+1)+"','"+str(json.dumps(Rutas[Tipo][Ruta]))+"',"+str(Cut_Time)+",NOW())")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)



def Crear_Programacion(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Adicional" }
        Formulario["Campos"].append({"tipo":"fecha","campo":"Fecha a Programar","titulo":"Fecha a Programar","Requerido":1,"min":1,"max":150,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Crear_Programacion_Guardar()'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            function Crear_Programacion_Guardar(){
                var Info = Dame_Formulario(".Adicional",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Crear_Programacion_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Crear_Programacion_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info = json.loads(Datos["Info"])
        Rutas = DB.Get_Dato("SELECT * FROM ilc_saltillo.crutas WHERE cr_tipo = 'Outbound'")
        Rutas = json.loads(str(Rutas[0]["cr_niveles"]))
        Aux_Ya = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_fecha = '"+str(Info["Fecha a Programar"])+"'")
        if len(Aux_Ya) >= 5:
            Error = "La fecha ["+str(Info["Fecha a Programar"])+"] ya está programada"
        else:
            for Tipo in Rutas.keys():
                if Tipo != "MilkRun":
                    for Ruta in Rutas[Tipo].keys():
                        if "CUTTIME_OUT_DIA" not in Rutas[Tipo][Ruta].keys():
                            Rutas[Tipo][Ruta]["CUTTIME_OUT_DIA"] = 0
                        Cut_Time = "null"
                        if str(Rutas[Tipo][Ruta]["CUTTIME_OUT"]) != "":
                            Cut_Time = datetime.strptime(str(Info["Fecha a Programar"])+" "+str(Rutas[Tipo][Ruta]["CUTTIME_OUT"])+":00","%Y-%m-%d %H:%M:%S")
                            if str(Rutas[Tipo][Ruta]["CUTTIME_OUT"]) != "":
                                Cut_Time = Cut_Time + timedelta(days=int(Rutas[Tipo][Ruta]["CUTTIME_OUT_DIA"]))
                            Cut_Time = "'"+str(Cut_Time.strftime("%Y-%m-%d %H:%M:%S"))+"'"
                        Error += DB.Instruccion("INSERT INTO ilc_saltillo.cload_id (clid_fecha,clid_ruta_tipo_1,clid_ruta_tipo_2,clid_ruta_tipo_ruta,clid_usuario_crea,clid_usuario_cuendo_crea,clid_ruta_configuracion,clid_cut_time) VALUES ('"+str(Info["Fecha a Programar"])+"','Outbound','"+str(Tipo)+"','"+str(Ruta)+"','"+str(Datos["ID_User"])+"',NOW(),'"+str(json.dumps(Rutas[Tipo][Ruta]))+"',"+str(Cut_Time)+")")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)



def Cancelar_Ruta(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_activo = 0,clid_id_usuario_cancelar = '"+str(Datos["ID_User"])+"',clid_fecha_hora_cancelacion = NOW(),clid_comentario_cancelacion = '"+str(Datos["Comentario"])+"'  WHERE clid_id = '"+str(Datos["ID"])+"' ")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if "Duplicate" in str(Error):
                Resultado["Contenido"] = "Este LOAD ID ya esta en RUTA"
            else:
                Resultado["Contenido"] = Error
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Cargar_Load_ID(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Cargar_Excel_Ryder" }
        Formulario["Campos"].append({"tipo":"archivo","campo":"Archivo_Ryder","titulo":"Archivo Ryder (ODC Linehaul Schedule.xlsx)","Requerido":1,"Col":12,"min":1,"max":1,"tipo_archivo":["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Cargar_Load_ID_Guardar()'><i class='mdi mdi-floppy'></i> Procesar</button></div>
        """
        Resultado["Contenido"] += """
        <script>
            function Cargar_Load_ID_Guardar(){
                var Info = Dame_Formulario(".Cargar_Excel_Ryder",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Cargar_Load_ID_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Cargar_Load_ID_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info_Datos = json.loads(str(Datos["Info"]))

        Load_IDs = []
        wb_obj = openpyxl.load_workbook("I:/Portal_File/"+str(Info_Datos["Archivo_Ryder"][0]))
        Shell = wb_obj.active

        index = 5
        while Shell["F"+str(index)].value is not None:
            Load_IDs.append({
                "Load No": Shell["F"+str(index)].value,
                "Ruta": str(Shell["D"+str(index)].value)+"-"+str(Shell["E"+str(index)].value),
                "Cut_Time": Shell["H"+str(index)].value,
                "Status" : Shell["J"+str(index)].value
            })
            index += 1

        for Load_ID in Load_IDs:
            if str(Load_ID["Status"]).strip() == "Cancelled":
                DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_fecha_hora_asignacion = NOW(), clid_usuario_asignacion = '"+str(Datos["ID_User"])+"', clid_load_id = '"+str(Load_ID["Load No"])+"', clid_activo = 0, clid_fecha_hora_cancelacion = NOW(), clid_comentario_cancelacion = 'Ryder File', clid_id_usuario_cancelar = '"+str(Datos["ID_User"])+"'  WHERE clid_ruta_tipo_ruta = '"+str(Load_ID["Ruta"])+"' AND clid_fecha = '"+str(Load_ID["Cut_Time"].strftime("%Y-%m-%d"))+"' AND clid_activo = 1 AND clid_id_caja IS NULL")
            else:
                DB.Instruccion("UPDATE ilc_saltillo.cload_id SET clid_fecha_hora_asignacion = NOW(), clid_usuario_asignacion = '"+str(Datos["ID_User"])+"', clid_load_id = '"+str(Load_ID["Load No"])+"'  WHERE clid_ruta_tipo_ruta = '"+str(Load_ID["Ruta"])+"' AND clid_fecha = '"+str(Load_ID["Cut_Time"].strftime("%Y-%m-%d"))+"' AND clid_activo = 1 AND clid_id_caja IS NULL")



        if Error == "":
            os.remove("I:/Portal_File/"+str(Info_Datos["Archivo_Ryder"][0]))
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] = str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


def POD_Salida(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "POD_Salida" }
        Archivos_Sin_Asignar = []
        for file in os.listdir("I:/Portal_File/Rutas ILC/"): 
            if file != "Asignadas":
                Archivos_Sin_Asignar.append(file)
        # for file in os.listdir("\\\\10.21.131.117\\Users\\lincAdmin\\Documents\\Proceso_REQUERIDO_1\\Script\\Procesos\\Escaneos\\"): 
        #     if file != "Asignadas":
        #         Archivos_Sin_Asignar.append("LOCAL_"+str(file))
        Formulario["Campos"].append({"tipo":"seleccion","id":"seleccinar_archivo","campo":"Documentos","titulo":"Documentos","Requerido":1,"Tipo_Opciones":"Opciones","Opciones":Archivos_Sin_Asignar,"valor":""})
        Formulario["Campos"].append({"titulo":"<a id='Ver_Archivo' href='' target='_blank' class='btn btn-primary btn-sm'><i class='mdi mdi-eye'></i> Ver Archivo</a>","Separacion":1})

        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='POD_Salida_Guardar(\""""+str(Datos["ID"])+"""\")'><i class='mdi mdi-floppy'></i> Save</button></div>
        <script>
            Actualizar_Cambia_Texto();
            $("#seleccinar_archivo").on( "change", function() {
                if( $(this).find('option:selected').val().includes("LOCAL_") )
                    $("#Ver_Archivo").attr('href','http://10.21.131.117/Portal_File/'+$(this).find('option:selected').val().replace("LOCAL_",""))
                else
                    $("#Ver_Archivo").attr('href','http://10.4.7.219:8080/Portal_File/Rutas%20ILC/'+$(this).find('option:selected').val())
            } );
            function POD_Salida_Guardar(ID){
                var Info = Dame_Formulario(".POD_Salida",true);
                if(Info != null)
                {
                    Info["ID"] = ID;
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"POD_Salida_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def POD_Salida_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Info = json.loads(Datos["Info"])
        Error = ""
        Info_Ahora = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(Info["ID"])+"'")[0]
        Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
        Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
        Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
        Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
        Info_Nueva = json.loads(str(Info_Ahora["cc_informacion_actual"]))
        Info_Nueva["Archivos"] = [str(Info["Documentos"]).replace("LOCAL_","")]
        Error += DB.Instruccion(""" 
        INSERT INTO ulc_saltillo.ccajas_moviemiento
        (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_info_ruta)
        VALUES
        ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(json.dumps(Info_Nueva))+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','POD Salida','"""+str(Info_Ahora["cc_info_ruta"])+"""')
        """)
        Archivos_Entrada = None
        if "Documentos" in Info.keys():
            Archivos_Entrada = str(Info["Documentos"])
        if Archivos_Entrada is not None:
            Archivos_Entrada = Archivos_Entrada.replace("'","")

            pdf_reader = PdfReader("I:/Portal_File/Rutas ILC/"+str(Archivos_Entrada))
            output = PdfWriter()
            for P in pdf_reader.pages:
                Texto = P.extract_text()
                Byts = 0
                for I in P.images:
                    Byts += len(I.data)
                if Byts > 250 or Texto != "":
                    P.compress_content_streams()
                    output.add_page(P)
            with open("I:/Portal_File/Rutas ILC/"+str(Archivos_Entrada), 'wb') as f:
                output.write(f)
                
            shutil.move("I:/Portal_File/Rutas ILC/"+str(Archivos_Entrada), "I:/Portal_File/"+str(Archivos_Entrada),)
        if Error != "":
            Resultado["Estado"] = 0
            Resultado["Contenido"] = str(Error)
        else:
            Completo(CLoadID = Info["ID"])
        
        

    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Pase_Salida_Ticket(IDCaja,Usuario):
    try:
        DB = LibDM_2023.DataBase()
        Compartido_2023 = LibDM_2023.Compartido()
        mysocket = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
        mysocket.connect((IP_Zebra, 9100))
        Caja = DB.Get_Dato("SELECT * FROM ulc_saltillo.ccajas WHERE cc_id = '"+str(IDCaja)+"'")[0]
        Info_Cajas_Dtalle = json.loads(str(Caja["cc_informacion_actual"]))

        Nombre = str(DB.Dame_Nombre_IDUsuario(Usuario))
        Valor = str(Caja["cc_qr_salida"])
        Sello = ""
        for K in Info_Cajas_Dtalle.keys():
            if "Sello" in str(K):
                Sello = Info_Cajas_Dtalle[K]
        Imprimir = """
        ^XA
        ^FO130,50^GFA,5966,5966,38,,:::::::hS01LFE,hQ01PF8,hP0SF8,hN07UF8,hM07WF,hL0YFC,hJ01gGF,hI01gHFC,hH01gIFE,hG01gKF,hG0QFJ0QF8,h0OFCN0OFC,gY07MFCQ07MFE,gX03MF8S0NF,gW01MFU01MF8,gV01LFEW07LF8,gV07KFEX01LFC,gU03KFEg0LFC,gT01KFEgG07KFC,gT0LFgH01KFE,gS07KFgI01KFE,gR01KF8gJ0KFE,gR0KFCgK07JFE,gQ03JFEgL03JFE,gP01JFEgM03JFE,gP07JF8gM03JFE,gO03JFCgN01JFE,gO0JFEgO01JFE,gN03JF8gO01JFE,gM01JFCgP01JFC,gM03IFEgQ01JFC,gM03IF8gQ01JFC,gI03F801FFEgR01JFC,gI07FC01FFgS01JF8,gI0FFE00FCgS01JF8,gH01IF00FgT01JF,007FF8007FFQ01IFgW03JFJ0IF,007FF800IFQ03IF8gV03IFEI01IF,00IF800IFQ03IF8gV03IFEI01FFE,00IF801IFQ03IF8gV07IFCI01FFE,:00IF001IFQ01IFgW07IF8I01FFE,00IF001FFEQ01IFgW0JFJ03FFE,01IF001FFER0FFEhK03FFC,01IF003FFER07FChK03FFC,01IF003FFER03F8hK03FFC,01FFE003FFEi03FFC,01FFE003FFCi07FFC,03FFE007FFCi07FF8,03FFE007FFCK01FFY01FFEN0FC001FFEK01IF8I07FF8,03FFE007FFC07FF07FF8007FF83FFE007FF800JF8003FFC3FC00JFCJ0JFEI07FF8,03FFC007FF807FF1IFC007FF83FFE00IF003JFE003FFC7FC03KFI03KF800IF8,07FFC007FF80IF3IFE00IF83FFE00IF00LF003FF8FF80LFI07KFC00IF,07FFC00IF80IF7IFE00IF03FFE00FFE01LF807FF9FF81LF800LFC00IF,07FFC00IF80NF00IF03FFE01FFE03LF807FFBFF83LFC01LFC00IF,07FF800IF80NF00IF03FFC01FFC07LF807KF87LFC03LFE00IF,07FF800IF01NF00IF03FFC03FFC07FFC7FFC07KF87FF87FFC07LFE01IF,0IF800IF01NF01IF03FFC03FF80IF03FFC07KF0IF01FFC07FF83FFE01FFE,0IF801IF01IFC7IF01FFE03FFC03FF81FFE01FFC0LF0FFE01FFC0IF01FFE01FFE,0IF801IF01IF03FFE01FFE03FFC07FF81FFC01FFC0LF1FFE01FFC0FFE01FFE01FFE,0IF001IF01IF01FFE01FFE03FFC07FF03FFC01FFC0JF801FFC01FFC0FFE01FFC01FFE,0IF001FFE03FFE01FFE01FFE03FFC07FF03FFC01FFC0IFC001FFC01FF81FFC01FFC03FFE,1IF003FFE03FFC01FFE03FFC03FFC0FFE03FF801FFC1IF8001FFEP01FFC03FFC,1IF003FFE03FFC01FFE03FFC01FFC0FFE07FF803FFC1IFI03IFP03FFC03FFC,1IF003FFE03FFC03FFC03FFC01FFC1FFC07FF803FFC1IFI03IFCO07FFC03FFC,1FFE003FFC03FFC03FFC03FFC01FFC1FFC07FF807FF81FFEI03JF8M01IF807FFC,3FFE003FFC07FFC03FFC07FFC01FFC1FF80NF81FFEI01KFL03JF807FF8,3FFE007FFC07FF803FFC07FF801FFC3FF80NF83FFEI01KFCJ03KF807FF8,3FFE007FFC07FF807FF807FF801FFC3FF00NF83FFCI01KFEI01LF807FF8,3FFC007FFC07FF807FF807FF801FFC3FF00NF83FFCJ0LF8007LF007FF8,3FFC007FF80IF807FF807FF801FFC7FE01NF83FFCJ07KF800MF00IF8,7FFC007FF80IF007FF80IF801FFC7FE01NF03FFCJ03KFC01IFE7FF00IF,7FFC007FF80IF007FF80IF001FFCFFE01FFEL07FFCK0KFC03IF07FF00IF,7FFC00IF80IF00IF80IF001FF8FFC01FFEL07FF8K03JFE07FFC0IF00IF,7FF800IF80IF00IF00IF001FF8FFC01FFEL07FF8L07IFE0IF00FFE00IF,7FF800IF01IF00IF00IF001FF9FF801FFEL07FF8M0IFE0FFE00FFE01IF,7FF801IF01FFE00IF01FFE001FF9FF803FFCL0IF8M07FFE1FFE00FFE01FFE,IF801IF01FFE00IF01FFE001FF9FF003FFC01FFE0IFN03FFE1FFE01FFE01FFE,IF801FFE01FFE01FFE01FFE001FFBFF003FFC01FFC0IFI01FFC01FFC1FFC01FFE01FFE,IF801FFE01FFE01FFE01FFE001FFBFE003FFC01FFC0IFI01FFC01FFC3FFC01FFC03FFE,IF803FFE03FFE01FFE01FFE001JFE003FFC03FFC0IFI03FFC01FFC3FFC03FFC03FFC,IF803FFC03FFC01FFE03FFC001JFC003FFC03FF81IFI03FFC03FF83FFC03FFC03FFC,IF807FFC03FFC03FFE03FFC001JFC003FFC07FF81FFEI03FFC03FF83FFE07FFC03FFC,7FFE1IF803FFC03FFC03FFC001JF8003FFE07FF01FFEI03FFC07FF83FFE1IF803FFC,7MF007FFC03FFC03FFC001JF8001IF1IF01FFEI03FFE1IF03MF807FFC,7MF007FF803FFC03FFC001JF8001LFE01FFEI03LFE03MF807FF8,7LFE007FF803FFC07FF8001JFI01LFC03FFEI01LFC03MF807FF8,3LFC007FF807FFC07FF8001JFJ0LF803FFCI01LF803JFBFF807FF8,1LFI07FF807FF807FF8001IFEJ07KF003FFCJ0LF001JF3FF807FF8,0KFEI0IF807FF807FF8001IFEJ03JFE003FFCJ07JFE001IFE7FF80IF8,03JF8I0IF007FF80IFI01IFCJ01JF8007FFCJ03JF8I0IF87FF80IF,007FF8gP03FF8R07FFCJ03FE,,:::::I03IFEgW03IFC,I03IFCgW0JF,I07IFCgV03IFE,I07IF8gV0JF8,I0JF8gU03IFE,I0JF8gU0JF8,001JFgU03IFE,001JFgT01JF8,003JFgT07IFE,003JFgS01JF,007JFgS0JFC,007JFgR03JF,007JFgR0JFC,00KFgQ07JF,00KFgP03JF8,00KFgP0JFE,00KFgO07JF8,00KFgN03JFC,00KF8gL01KF,00KF8gL0KF8,00KFCgK07JFE,00KFEgJ03KF,00KFEgI01KFC,00LFgH01KFE,007KF8gG0LF,007KFEg0LF8,007LFY0LFC,003LFCW0LFE,003MFU01MF,001MFES03MF8,I0NFCQ0NFC,I07NFCN07NFE,I03PFEI01PFE,I01gKF,J0gJF,J07gHF,J01gGF,K07XFE,K01WFC,L03UFC,M07RFE,N03OFE,P0MF,,:::::::^FS
        ^FO30,220^GB495,3,3^FS
        ^CF0,30
        ^FO170,230^FDPASE DE SALIDA^FS
        ^CF0,20
        ^FO183,260^FD"""+str(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))+"""^FS
        ^FO145,270 ^BY1,1.0,20^BQN,2,9^FD,,"""+str(Valor)+"""^FS
        ^CF0,20
        ^FO50,600^FDCAJA^FS
        ^CF0,50
        ^FO50,630^FD"""+str(Caja["cc_contenedor"])+"""^FS
        ^CF0,20
        ^FO50,710^FDSELLO^FS
        ^CF0,50
        ^FO50,740^FD"""+str(Sello)+"""^FS
        ^FO30,810^GB495,3,3^FS
        ^CF0,18
        ^FO100,960^FD"""+str(Nombre)+"""^FS
        ^FO90,975^GB400,1,1^FS
        ^CF0,25
        ^FO200,980^FDFIRMA Y NOMBRE^FS
        ^CF0,20
        ^FO247,1000^FDDespacho^FS
        ^CF0,20
        ^FO50,1040^FDFECHA Y HORA^FS
        ^FO60,1060^FDDE SALIDA^FS
        ^FO180,2000^GB320,1,1^FS
        ^XZ
        """


        mysocket.send(Imprimir.encode())
        mysocket.close ()
    except:
        pass


def Ver_Historico_Rip(Datos):
    Compartido_2023 = LibDM_2023.Compartido()
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Estado":0,"Contenido":""}
    try:
        Historico = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip_historico WHERE  criph_master = '"+str(Datos["ID"])+"'")
        Tabla_Datos = []
        Columnas = []
        Columnas.append({"title":"Fecha y Hora", "field":"Fecha y Hora"})
        Columnas.append({"title":"Usuario", "field":"Usuario"})
        Columnas.append({"title":"Tipo", "field":"Tipo"})
        Columnas.append({"title":"Comentario", "field":"Comentario"})
        Columnas.append({"title":"Archivo(s)", "field":"Archivos","formatter":"html"})
        for H in Historico:
            Aux_Datos ={}
            Aux_Datos["Fecha y Hora"] = H["criph_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S")
            Aux_Datos["Usuario"] = str(DB.Dame_Nombre_IDUsuario(H["criph_user"]))
            Aux_Datos["Tipo"]  = H["criph_tipo"]
            Aux_Datos["Comentario"]  = H["criph_comentario"]
            Aux_Datos["Archivos"] = ""
            for Archivo in str(H["criph_archivo"]).split(","):
                if Archivo.strip() != "":
                    if "pdf" in Archivo:
                        Aux_Datos["Archivos"] += "<a href='http://10.4.7.219:8080/Portal_File/"+str(Archivo)+"' target='_blank' class='mdi mdi-file-pdf-box ms-1'></a>"
                    else:
                        Aux_Datos["Archivos"] += "<a href='http://10.4.7.219:8080/Portal_File/"+str(Archivo)+"' target='_blank' class='mdi mdi-image ms-1'></a>"
            Tabla_Datos.append(Aux_Datos)

        Resultado["Contenido"] += """
        <div id='Tabla-Detalles' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete table;
            var Col = """+str(Columnas)+""";
            var table = new Tabulator("#Tabla-Detalles", {
                minHeight:300,
                data:"""+str(Tabla_Datos)+""",
                layout:"fitColumns",
                pagination:"local",
                rowFormatter:function(row){
                if(row.getData().Estado == 1){
                        row.getElement().style.backgroundColor = "#5dff67";
                    }
                },
                paginationSize:50,  
                paginationCounter:"rows",
                columns:Col,
                selectable:false
            });
            document.getElementById("download-xlsx-ruteo").addEventListener("click", function(){
                table.download("xlsx", "Ruteo.xlsx", {sheetName:"My Data"});
            });
            $( document ).ready(function() {
                $('.tabulator-header-contents').addClass('bg-body-secondary').find('.tabulator-col').addClass('bg-body-secondary');
            })
        </script>
        """
    except:
        Resultado["Contenido"] += str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Modificar_Ruteo(Datos):
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Estado":0,"Contenido":""}
    DB = LibDM_2023.DataBase()
    try:
        Tabla = []
        Partes = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip_partes WHERE cripp_master = '"+str(Datos["ID"])+"'")
        for P in Partes:
            Tabla.append({"ASN":str(P["cripp_asn"]),"Parte":str(P["cripp_parte"]),"Cantidad":str(P["cripp_cantidad"]),"DUNS":str(P["cripp_duns"]),"Destino":str(P["cripp_destino"])})
        Resultado["Contenido"] += """
        <h1>Partes</h1>
        <div id="example-table"></div>
        <script>
            delete tabledata;
            var tabledata = """+str(Tabla)+""";
            delete table;
            var table = new Tabulator("#example-table", {
                height:"311px",
                layout:"fitColumns",
                reactiveData:true, //turn on data reactivity
                data: tabledata,
                columns:[
                    {title:"ASN", field:"ASN", hozAlign:"center", headerFilter: 'input'},
                    {title:"Parte", field:"Parte", hozAlign:"center", headerFilter: 'input'},
                    {title:"Cantidad", field:"Cantidad", hozAlign:"center", headerFilter: 'input'},
                    {title:"DUNS", field:"DUNS", hozAlign:"center", headerFilter: 'input'},
                    {title:"Destino", field:"Destino", hozAlign:"center", headerFilter: 'input'},
                ],
            });
            table.on("tableBuilt", function(){ 
                setTimeout(() => {
                    table.clearFilter();
                    table.redraw();
                }, 100);
            });
        </script>
        <hr>
        <h1>Historico</h1>
        """
        Historico = DB.Get_Dato("SELECT * FROM ilc_saltillo.crip_historico WHERE  criph_master = '"+str(Datos["ID"])+"'")
        Tabla_Datos = []
        Columnas = []
        Columnas.append({"title":"Fecha y Hora", "field":"Fecha y Hora"})
        Columnas.append({"title":"Usuario", "field":"Usuario"})
        Columnas.append({"title":"Tipo", "field":"Tipo"})
        Columnas.append({"title":"Comentario", "field":"Comentario"})
        Columnas.append({"title":"Archivo(s)", "field":"Archivos","formatter":"html"})
        for H in Historico:
            Aux_Datos ={}
            Aux_Datos["Fecha y Hora"] = H["criph_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S")
            Aux_Datos["Usuario"] = str(DB.Dame_Nombre_IDUsuario(H["criph_user"]))
            Aux_Datos["Tipo"]  = H["criph_tipo"]
            Aux_Datos["Comentario"]  = H["criph_comentario"]
            Aux_Datos["Archivos"] = ""
            for Archivo in str(H["criph_archivo"]).split(","):
                if Archivo.strip() != "":
                    if "pdf" in Archivo:
                        Aux_Datos["Archivos"] += "<a href='http://10.4.7.219:8080/Portal_File/"+str(Archivo)+"' target='_blank' class='mdi mdi-file-pdf-box ms-1'></a>"
                    else:
                        Aux_Datos["Archivos"] += "<a href='http://10.4.7.219:8080/Portal_File/"+str(Archivo)+"' target='_blank' class='mdi mdi-image ms-1'></a>"
            Tabla_Datos.append(Aux_Datos)

        Resultado["Contenido"] += """
        <div id='Tabla-Detalles' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete table;
            var Col = """+str(Columnas)+""";
            var table = new Tabulator("#Tabla-Detalles", {
                minHeight:300,
                data:"""+str(Tabla_Datos)+""",
                layout:"fitColumns",
                pagination:"local",
                rowFormatter:function(row){
                if(row.getData().Estado == 1){
                        row.getElement().style.backgroundColor = "#5dff67";
                    }
                },
                paginationSize:50,  
                paginationCounter:"rows",
                columns:Col,
                selectable:false
            });
            document.getElementById("download-xlsx-ruteo").addEventListener("click", function(){
                table.download("xlsx", "Ruteo.xlsx", {sheetName:"My Data"});
            });
            $( document ).ready(function() {
                $('.tabulator-header-contents').addClass('bg-body-secondary').find('.tabulator-col').addClass('bg-body-secondary');
            })
        </script>
        <hr>
        """
        Formulario = {"Col":"", "Campos": [],"Clase": "Formato" }
        Formulario["Campos"].append({"tipo":"multitexto","campo":"Comentario","titulo":"Comentario","editable":True,"Requerido":1,"min":1,"max":1500,"valor":"","Col":12})
        Formulario["Campos"].append({"tipo":"archivo","campo":"Archivos","titulo":"Archivo(s)","Requerido":1,"Col":12,"min":0,"max":5,"tipo_archivo":["application/pdf","image/png","image/jpeg","image/gif"],"valor":"","editable":True})
        #Formulario["Campos"].append({"tipo":"radio","campo":"Enviar a Despacho?","titulo":"Enviar a Despacho?","Tipo_Opciones":"Opciones","Opciones":["Si","No"],"editable":True,"Requerido":1,"min":1,"max":1500,"valor":"No","Col":12})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <br>
        <div class='w-100 text-center mb-2'><button id='btn_guardar' class='btn btn-warning w-50' onclick='Modificar("Modificar","""+str(Datos["ID"])+""")'><i class='mdi mdi-floppy'></i> Modificar</button></div>
        <script>
            function Modificar(Tipo,ID){
                var Info = Dame_Formulario(".Formato",true);
                if(Info != null){
                    Info["ID"] = ID;
                    Info["Tipo"] = Tipo;
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Modificar","Info":JSON.stringify(Info)};
                        $.ajax({
                            data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                            success:  function (response){
                                var obj = JSON.parse(response);
                                if(obj["Estado"] == 1)
                                {
                                    $("#Vent_1").modal("hide");
                                    Swal.fire({icon: 'success',position: 'top-end',title: '¡Guardado con éxito!',showConfirmButton: false,toast: true,background : "#c9fad7",timer: 1500,timerProgressBar: true});
                                    Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");
                                }
                                else
                                {
                                    Swal.fire({icon: 'error',position: 'top-end',title: 'Process error ['+ obj["Contenido"] +']',showConfirmButton: false,toast: true,background : "#fac9c9",timer: 3500,timerProgressBar: true});
                                }
                            },
                            error: function (jqXHR, textStatus, errorThrown )
                            {
                                Swal.fire({icon: 'error',position: 'top-end',title: 'Process error ['+ textStatus +']',showConfirmButton: false,toast: true,background : "#fac9c9",timer: 3500,timerProgressBar: true});
                            }
                    });
                }
            }
        </script>
        """
    except:
        Resultado["Contenido"] += str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Modificar(Datos):
    DB = LibDM_2023.DataBase()
    Cur = "Content-Type: text/html;charset=ISO-8859-1\n\n"
    Resultado = { "Estado" :0, "Contenido":""}
    try:
        ID_Maste = None
        Error = ""
        Info = json.loads(Datos["Info"])
        if Info["Tipo"] == "Modificar":
            Error += DB.Instruccion("INSERT INTO ilc_saltillo.crip_historico (criph_master,criph_fecha_hora,criph_comentario,criph_archivo,criph_user,criph_tipo) VALUES ('"+str(Info["ID"])+"',NOW(),'"+str(Info["Comentario"])+"','"+str(",".join(Info["Archivos"]))+"','"+str(Datos["ID_User"])+"','DESPACHO')")
            Error += DB.Instruccion("UPDATE ilc_saltillo.crip SET crip_despacho = 2 WHERE crip_id = '"+str(Info["ID"])+"'")
        if Info["Tipo"] == "Solucionar":
            Error += DB.Instruccion("INSERT INTO ilc_saltillo.crip_historico (criph_master,criph_fecha_hora,criph_comentario,criph_archivo,criph_user,criph_tipo) VALUES ('"+str(Info["ID"])+"',NOW(),'"+str(Info["Comentario"])+"','"+str(",".join(Info["Archivos"]))+"','"+str(Datos["ID_User"])+"','SOLUCION')")
            Error += DB.Instruccion("UPDATE ilc_saltillo.crip SET crip_estado = 2, crip_solucionado_user = '"+str(Datos["ID_User"])+"',crip_solucinado_fecha_hora = NOW() WHERE crip_id = '"+str(Info["ID"])+"'")
        if Info["Tipo"] == "Despacho":
            Error += DB.Instruccion("INSERT INTO ilc_saltillo.crip_historico (criph_master,criph_fecha_hora,criph_comentario,criph_archivo,criph_user,criph_tipo) VALUES ('"+str(Info["ID"])+"',NOW(),'"+str(Info["Comentario"])+"','"+str(",".join(Info["Archivos"]))+"','"+str(Datos["ID_User"])+"','ENVIAR A DESPACHO')")
            Error += DB.Instruccion("UPDATE ilc_saltillo.crip SET crip_despacho = 1, crip_idcaja = '"+str(Info["Caja"])+"' WHERE crip_id = '"+str(Info["ID"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            if ID_Maste is not None:
                DB.Instruccion("DELETE FROM ilc_saltillo.crip_partes WHERE cripp_master = '"+str(ID_Maste)+"'")
                DB.Instruccion("DELETE FROM ilc_saltillo.crip_historico WHERE criph_master = '"+str(ID_Maste)+"'")
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Enviar_Nuevo_Reporte(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Gen = []
        Ultima_Actualizacion = DB.Get_Dato("SELECT ultima_zona_wms_ocs FROM servicios.cvariables_general")[0]["ultima_zona_wms_ocs"]
        if Ultima_Actualizacion + timedelta(hours=1) < datetime.now():
            #Resultado["Contenido"] += "<div class='text-center'><button class='btn btn-primary mb-1 me-2' onclick='Actualizar_Zona()'><i class='mdi mdi-cloud-arrow-down'></i> Actualizar Zonas</button></div>"
            Resultado["Contenido"] += "<div class='text-center text-danger fw-bold'>ACTUALIZA LAS ZONAS PARA PODER ENVIAR EL CORREO DE REPORTE</div>"
        else:
            Info = json.loads(str(Datos["Info"]))
            LoadIDs = DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id IN ("+",".join(Info["IDs"])+")")
            for LoadID in LoadIDs:
                Rutas_Info = json.loads(LoadID["clid_ruta_configuracion"])
                Item = {}
                # Item["Ruta"] = str(LoadID["clid_ruta_tipo_ruta"])
                Item["Carrier"] = str(Rutas_Info["SCAC"])
                Item["Destino"] = str(Rutas_Info["NOMBRE"])
                Item["Zona"] = str(Rutas_Info["ZONA"])
                Item["LoadIDs"] = []
                Item["Plantas"] = []
                LoadID["Plantas"] = []
                for Z in DB.Get_Dato("SELECT cwz_zona,cwz_last_scan,cwz_paking_slip,cwz_part,cwz_serial,cwz_pack_qty,cwz_rout_date,cwz_plant,cwz_weight,cwz_supplier_code FROM ilc_saltillo.cwms_zonas WHERE cwz_actual = 1 and cwz_zona = 'Z"+str(Rutas_Info["ZONA"])+"'"):
                    if str(Z["cwz_plant"]).split("-")[0] not in Item["Plantas"]:
                        Item["Plantas"].append(str(Z["cwz_plant"]).split("-")[0])
                        LoadID["Plantas"].append(str(Z["cwz_plant"]).split("-")[0])
                if Item not in Info_Gen:
                    Info_Gen.append(Item)
            
            for LoadID in LoadIDs:
                Rutas_Info = json.loads(LoadID["clid_ruta_configuracion"])
                Planta = []
                for Item in Info_Gen:
                    if str(Item["Carrier"]) == str(Rutas_Info["SCAC"]) and str(Item["Destino"]) == str(Rutas_Info["NOMBRE"]) and Item["Plantas"] == LoadID["Plantas"]:
                        Item["LoadIDs"].append({"LoadID":str(LoadID["clid_load_id"]),"Caja Pre-Cargada":str(LoadID["clid_caja_pre_asignada"]),"Fecha del LOAD ID":str(LoadID["clid_fecha"]),"Tipo":str(LoadID["clid_ruta_tipo_2"]),"Ruta":str(LoadID["clid_ruta_tipo_ruta"]),"Cut Time":str(LoadID["clid_cut_time"]),"ID":str(LoadID["clid_id"])})


            for Item in Info_Gen:
                Resultado["Contenido"] += "<div>Destino: <b>"+str(Item["Destino"])+"</b></div>"
                Resultado["Contenido"] += "<div class='row'><div class='col'>SCAC: <b>"+str(Item["Carrier"])+"</b></div><div class='col'>Plantas: <b>"+str(",".join(Item["Plantas"]))+"</b></div><div class='col'>Zona: <b>"+str(Item["Zona"])+"</b></div></div>"
                
                Resultado["Contenido"] += "<div class='row'>"
                for LoadID in Item["LoadIDs"]:
                    Resultado["Contenido"] += """
                    <div class='col-auto p-1'>
                        <div class='w-100 h-100 border bg-warning-subtle p-1'>
                        <div>LoadID: <b>"""+str(LoadID["LoadID"])+"""</b></div>
                        <div>Caja Pre-Cargada: <b>"""+str(LoadID["Caja Pre-Cargada"])+"""</b></div>
                        <div>Fecha del LOAD ID: <b>"""+str(LoadID["Fecha del LOAD ID"])+"""</b></div>
                        <div>Ruta: <b>"""+str(LoadID["Ruta"])+"""</b></div>
                        <div>Cut Time: <b>"""+str(LoadID["Cut Time"])+"""</b></div>
                        </div>
                    </div>
                    """
                Resultado["Contenido"] += "</div><hr>"

            Resultado["Contenido"] += """
            <hr>
            <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Enviar_Nuevo_Reporte_Enviar()'><i class='mdi mdi-email'></i> Enviar</button></div>
            <script>
                function Enviar_Nuevo_Reporte_Enviar(){
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Enviar_Nuevo_Reporte_Enviar","Info":JSON.stringify("""+str(Info_Gen)+""")};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
                                if("Pase Salida" in Resultado)
                                {
                                    var win1 = window.open(Resultado["Pase Salida"], '_blank');
                                    win1.focus();
                                }
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Cargar_Fecha($("#fecha").val());
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            </script>
            """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Enviar_Nuevo_Reporte_Enviar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":1}
    try:
        Resultado["Contenido"] += str(Datos)
        Error = ""
        Info = json.loads(Datos["Info"])
        for ITEM in Info:
            HTML = ""
            SCAC = ""
            RUTA = ""
            Destinos = ["dmarmolejo@universallogistics.com","LVictoriaArevalo@universallogistics.com","DLopez2@universallogistics.com","LHernandez@universallogistics.com","VJesusAlberto@universallogistics.com","NHernandezDiaz@universallogistics.com","VMorales@universallogistics.com","edeheza@universallogistics.com"]
            #Destinos = ["dmarmolejo@universallogistics.com"]
            for LoadID in ITEM["LoadIDs"]:
                ConfLoadID = json.loads(DB.Get_Dato("SELECT * FROM ilc_saltillo.cload_id WHERE clid_id = '"+str(LoadID["ID"])+"'")[0]["clid_ruta_configuracion"])
                if "CORREOS_FALTA" in ConfLoadID.keys():
                    for C in str(ConfLoadID["CORREOS_FALTA"]).split(";"):
                        if str(C) not in Destinos:
                            Destinos.append(str(C))
                    

            Zonas_Actual  = DB.Get_Dato("select cwz_zona,group_concat(distinct(cwz_paking_slip)) as PS, MIN(cwz_rout_date) as PS_Min,cwz_fecha_hora from ilc_saltillo.cwms_zonas where cwz_actual = 1 and cwz_zona = 'Z"+str(ITEM["Zona"])+"' group by cwz_zona")
            Partes_En_Zona = DB.Get_Dato("SELECT * FROM ilc_saltillo.cwms_zonas WHERE cwz_actual = 1 and cwz_zona = 'Z"+str(ITEM["Zona"])+"'")
            Zonas = []
            for Z in Partes_En_Zona:
                Aux = {}
                for k in Z.keys():
                    if isinstance(Z[k], datetime) or isinstance(Z[k], date):
                        Z[k] = str(Z[k].strftime("%Y-%m-%d %H:%M"))
                    Aux[k] = Z[k]
                Zonas.append(Aux)
            LoadID_Texto = []
            for LoadID in ITEM["LoadIDs"]:
                if str(LoadID["LoadID"]) not in LoadID_Texto:
                    LoadID_Texto.append(str(LoadID["LoadID"]))
            Rutas_Texto = []
            for LoadID in ITEM["LoadIDs"]:
                if str(LoadID["Ruta"]) not in Rutas_Texto:
                    Rutas_Texto.append(str(LoadID["Ruta"]))

            Asunto = "SLP ILC NO TRAILER / LOAD ID: "+str(",".join(LoadID_Texto))+" / SCAC: "+str(ITEM["Carrier"])+" / ROUTE: "+str(",".join(Rutas_Texto))+" / DESTINATION: "+str(ITEM["Destino"])
            
            HTML += """
            <div style='color:#f10000; background:#edf100; font-weight:bold;'>
            ILC will be not responsible for any impact related to the material listed below. We are a droop and hook operation and the trailer should be in our locations 24 hours prior to its cut window. If the GM plant needs or want to expedite all or partially material please provide the AA number
            </div>
            """
            

            HTML += "<div style='color:#ffffff; background:#000000; font-weight:bold;'><h3>GENERAL INFORMATION</h3></div>"
            for LoadID in ITEM["LoadIDs"]: 
                HTML += "<div>LOAD ID:<span style='text-decoration:underline; font-weight:bold'>"+str(LoadID["LoadID"])+"</span></div>"
                #HTML += "<div>CONTAINER:<span style='text-decoration:underline; font-weight:bold'>"+str(LoadID["Caja Pre-Cargada"])+"</span></div>"
                HTML += "<div>ROUTE:<span style='text-decoration:underline; font-weight:bold'>"+str(LoadID["Ruta"])+"</span></div>"
                HTML += "<div>SCAC:<span style='text-decoration:underline; font-weight:bold'>"+str(ITEM["Carrier"])+"</span></div>"
                HTML += "<div>CUT TIME:<span style='text-decoration:underline; font-weight:bold'>"+str(LoadID["Cut Time"])+"</span></div>"
                HTML += "<div>DESTINATION:<span style='text-decoration:underline; font-weight:bold'>"+str(ITEM["Destino"])+"</span></div>"
                HTML += "<div>PLANS:<span style='text-decoration:underline; font-weight:bold'>"+str(",".join(ITEM["Plantas"]))+"</span></div><br>"
                # HTML += "<br><div style=' background:#edf100;'>REQUIRED TRAILER: <span style='text-decoration:underline; font-weight:bold'> "+str(Info["Cajas necesarias"])+" </span></div>"
            


            HTML += "<div style='color:#ffffff; background:#000000; font-weight:bold;'><h3>ASNs AFFECTED</h3></div> <h3>ZONE: <span style='color:#f10000;'><b>Z"+str(ITEM["Zona"])+"</b></span> / DESTINATION: <span style='color:#f10000;'><b>"+str(ITEM["Destino"])+"</b></span> / CARRIER: <span style='color:#f10000;'><b>"+str(ITEM["Carrier"])+"</b></span></h3>"
            HTML += """
            <table style='border: 1px solid black; border-collapse: collapse; table-layout: fixed;'>
                <thead>
                    <tr>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>ASN</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>PART</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>SERIAL</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>QTY</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>ROUTE INBOUND</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>DUNS</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>CISCO</th>
                    </tr>
                </thead>
                <tbody>
            """
            for Z in Zonas:
                HTML += """
                <tr>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_paking_slip"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_part"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_serial"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_pack_qty"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_route"])+""" ["""+str(Z["cwz_rout_date"])+"""]</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_supplier_code"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_plant"])+"""</td>
                </tr>
                """
            HTML += "</tbody></table>"
            HTML += """
            <div style='color:#ffffff; background:#000000; font-weight:bold;'><h3>ADDITIONAL INFORMATION</h3></div>
            
            <div style='color:#ff0000; padding:2px; font-weight:bold;'>Material in OSD waiting to be resolved by the supplier that cannot be shipped at this time.</div>

            <h3>ZONE: <span style='color:#f10000;'><b>OS&D</b></span> / CISCOS: <span style='color:#f10000;'><b>"""+str(",".join(ITEM["Plantas"]))+"""</b></span></h3>
            """
            HTML += """
            <table style='border: 1px solid black; border-collapse: collapse; table-layout: fixed;'>
                <thead>
                    <tr>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>ASN</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>PART</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>SERIAL</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>QTY</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>ROUTE INBOUND</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>DUNS</th>
                        <th style='border: 1px solid black; border-collapse: collapse; background:#DDDDDD;'>CISCO</th>
                    </tr>
                </thead>
                <tbody>
            """
            Partes_En_Zona_WMS = DB.Get_Dato("SELECT * FROM ilc_saltillo.cwms_zonas where cwz_actual = 1 and cwz_zona = 'ZOSD'")
            Zonas = []
            for Z in Partes_En_Zona_WMS:
                Aux = {}
                for k in Z.keys():
                    if isinstance(Z[k], datetime) or isinstance(Z[k], date):
                        Z[k] = str(Z[k].strftime("%Y-%m-%d %H:%M"))
                    Aux[k] = Z[k]
                if str(Z["cwz_plant"]).split("-")[0] in ITEM["Plantas"]:
                    Zonas.append(Aux)

            for Z in Zonas:
                HTML += """
                <tr>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_paking_slip"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_part"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_serial"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_pack_qty"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_route"])+""" ["""+str(Z["cwz_rout_date"])+"""]</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_supplier_code"])+"""</td>
                    <td style='border: 1px solid black; border-collapse: collapse; padding:5px;'>"""+str(Z["cwz_plant"])+"""</td>
                </tr>
                """
            HTML += "</tbody></table>"
            HTML += """
            <hr>
            Atentamente: ILC/Universal
            <div style='font-weight: bold;'>Do not reply to this email.</div>
            """

            Error = DB.Instruccion("INSERT INTO portal.ccorreos_enviar (cce_usuario,cce_asunto,cce_to,cce_contendio) VALUES ('"+str(Datos["ID_User"])+"','"+str(Asunto)+"','"+str(",".join(Destinos))+"','"+str(HTML).replace("'","''")+"')")
            for LoadID in ITEM["LoadIDs"]: 
                Error += DB.Instruccion("INSERT INTO ilc_saltillo.ccorreos_historico (cch_fecha_hora,cch_usuario,cch_tipo,cch_destinos,cch_loadid) VALUES (NOW(),'"+str(Datos["ID_User"])+"','FALTA CAJA','"+str(",".join(Destinos))+"','"+str(LoadID["ID"])+"')")
            
            if Error == "":
                Resultado["Estado"] = 1
            else:
                Resultado["Estado"] = 0
                Resultado["Contenido"] = Error
    except:
        Resultado["Estado"] = 0
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Ver_Historico_Correo(Datos):
    Compartido_2023 = LibDM_2023.Compartido()
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Estado":0,"Contenido":""}
    try:
        Historico = DB.Get_Dato("SELECT * FROM ilc_saltillo.ccorreos_historico WHERE cch_loadid = '"+str(Datos["ID"])+"'")
        Tabla_Datos = []
        Columnas = []
        Columnas.append({"title":"Fecha y Hora", "field":"Fecha y Hora"})
        Columnas.append({"title":"Usuario", "field":"Usuario"})
        Columnas.append({"title":"Tipo", "field":"Tipo"})
        Columnas.append({"title":"Correos", "field":"Correos"})
        for H in Historico:
            Aux_Datos ={}
            Aux_Datos["Fecha y Hora"] = H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S")
            Aux_Datos["Usuario"] = str(DB.Dame_Nombre_IDUsuario(H["cch_usuario"]))
            Aux_Datos["Tipo"]  = H["cch_tipo"]
            Aux_Datos["Correos"]  = H["cch_destinos"]
            Tabla_Datos.append(Aux_Datos)

        Resultado["Contenido"] += """
        <div id='Tabla-Detalles' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete table;
            var Col = """+str(Columnas)+""";
            var table = new Tabulator("#Tabla-Detalles", {
                minHeight:300,
                data:"""+str(Tabla_Datos)+""",
                layout:"fitColumns",
                pagination:"local",
                rowFormatter:function(row){
                if(row.getData().Estado == 1){
                        row.getElement().style.backgroundColor = "#5dff67";
                    }
                },
                paginationSize:50,  
                paginationCounter:"rows",
                columns:Col,
                selectable:false
            });
            document.getElementById("download-xlsx-ruteo").addEventListener("click", function(){
                table.download("xlsx", "Ruteo.xlsx", {sheetName:"My Data"});
            });
            $( document ).ready(function() {
                $('.tabulator-header-contents').addClass('bg-body-secondary').find('.tabulator-col').addClass('bg-body-secondary');
            })
        </script>
        """
    except:
        Resultado["Contenido"] += str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


def Actualizar_Zona(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "WMS" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Usuario_WMS","titulo":"Usuario del WMS","Requerido":1,"min":1,"max":150,"valor":"Usuario WMS"})
        Formulario["Campos"].append({"tipo":"texto","campo":"Pass_WMS","titulo":"Contraseña del WMS","Requerido":1,"min":1,"max":150,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Actualizar_Zona_Procesar()'><i class='mdi mdi-cloud-arrow-down'></i> Automático</button></div>
        <span id='Error'></span>
        <hr>
        <div class='w-100 text-center'><button class='btn btn-dark w-75' onclick='Cargar_Excel()'><i class='mdi mdi-microsoft-excel'></i> Cargar Excel</button></div>
        <script>
        var TimeOut_Rutas = null;
        function Actualizar_Zona_Procesar(ID){
            var Info = Dame_Formulario(".WMS",true);
            if(Info != null)
            {
                Mostrar_Ventana_Cargando(false);
                var parametros = {"Fun":"Actualizar_Zona_Procesar","Info":JSON.stringify(Info)};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        
                        var Resultado = JSON.parse(response);
                        if(Resultado["Estado"] == 1)
                        {
                            Esperando_Respuesta(Resultado["ID"]);
                        }
                        else{
                            Mensaje(0,Resultado["Contenido"]);
                            $("#Error").html(Resultado["Contenido"])
                        }

                        
                    },
                    error: function (jqXHR, textStatus, errorThrown )
                    {
                        Mensaje(0,textStatus);
                    }
                });
            }
        }
        function Esperando_Respuesta(ID){
            var parametros = {"Fun":"Esperando_Respuesta","ID":ID};
            $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                success:  function (response)
                {
                    var Resultado = JSON.parse(response);
                    if(Resultado["Estado"] == 1)
                    {
                        if(Resultado["Respuesta"] == 0){
                            if(TimeOut_Rutas)
                                clearTimeout(TimeOut_Rutas);
                            TimeOut_Rutas =  setTimeout(function() { Esperando_Respuesta(ID); }, 500);
                        }
                        else{
                            if(Resultado["Contenido"] == "OK"){
                                $("#Vent_1").modal("hide");
                                Mensaje(2);
                                Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");
                            }else{
                                Mensaje(0);
                                $("#Error").html(Resultado["Contenido"])
                            }
                        
                        }
                        
                    }
                    else{
                       if(Resultado["Respuesta"] == 0){
                            if(TimeOut_Rutas)
                                clearTimeout(TimeOut_Rutas);
                            TimeOut_Rutas =  setTimeout(function() { Esperando_Respuesta(ID); }, 500);
                        }
                    }                        
                },
                error: function (jqXHR, textStatus, errorThrown )
                {
                    Mensaje(0,textStatus);
                }
            });
        }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Actualizar_Zona_Procesar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    Error = ""
    try:
        Info_Datos = json.loads(str(Datos["Info"]))
        Variables = {"Usr":str(Info_Datos["Usuario_WMS"]),"Pass":str(Info_Datos["Pass_WMS"])}
        Error = str(DB.Instruccion("INSERT INTO servicios.web (cw_funcion,cw_parametros,cw_fecha_hora) VALUES ('WMS_Zonas_ILC','"+str(json.dumps(Variables))+"',NOW())"))
        if Error == "":
            Resultado["ID"] = DB.Get_Dato("SELECT MAX(cw_id) AS ID FROM servicios.web WHERE cw_parametros = '"+str(json.dumps(Variables))+"' AND cw_funcion = 'WMS_Zonas_ILC'")[0]["ID"]
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Esperando_Respuesta(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0,"Respuesta":0}
    Error = ""
    try:
        Resultado_DB = DB.Get_Dato("SELECT * FROM servicios.web WHERE cw_id = '"+str(Datos["ID"])+"'")[0]
        if Resultado_DB["cw_respuesta"] is None:
            Resultado["Respuesta"] = 0
        else:
            Res = json.loads(Resultado_DB["cw_respuesta"])
            Resultado["Respuesta"] = 1
            if "incorrecta" in str(Res["Resultado"]):
                Resultado["Contenido"] = "<div class='text-center h3 text-danger'>Usuario y/o contarse incorrecta</div><div class='text-center text-bg-danger'>Antes de intentar de nueva entra directamente al WMS y corrobora tu usuarios y contraseña. Evita bloquear tu usuario</div>"
            else:
                Resultado["Contenido"] = Res["Resultado"]
        Resultado["Estado"] = 1
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Cargar_Excel(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Cargar_Excel" }
        Formulario["Campos"].append({"tipo":"archivo","campo":"Excel","titulo":"Zone Status Report (Excel)","Requerido":1,"Col":12,"min":1,"max":1,"tipo_archivo":["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Cargar_Excel_Guardar()'><i class='mdi mdi-floppy'></i> Procesar</button></div>
        """
        Resultado["Contenido"] += """
        <script>
            function Cargar_Excel_Guardar(){
                var Info = Dame_Formulario(".Cargar_Excel",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Cargar_Excel_Guardar","Info":JSON.stringify(Info)};
                    $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                        success:  function (response)
                        {
                            var Resultado = JSON.parse(response);
                            if(Resultado["Estado"] == 1)
                            {
        """
        if int(Datos["Cargar"]) == 1:
            Resultado["Contenido"] += """
                $("#Vent_1").modal("hide");
                Mensaje(2);
                Imprimir_Zona();
            """
        else:
            Resultado["Contenido"] += """
                $("#Vent_1").modal("hide");
                Mensaje(2);
                Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");
            """
        Resultado["Contenido"] += """
                            }
                            else
                                Mensaje(0,Resultado["Contenido"]);
                                
                        },
                        error: function (jqXHR, textStatus, errorThrown )
                        {
                            Mensaje(0,textStatus);
                        }
                    });
                }
            }
            function Imprimir_Zona(){
                Mostrar_Ventana_Cargando(false);
                var parametros = {"Fun":"Imprimir_Zona"};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        if(Resultado["Estado"] == 1){
                            if("Archivo" in Resultado)
                            {
                                var win1 = window.open(Resultado["Archivo"], '_blank');
                                win1.focus();
                            }
                            Mensaje(2);
                        }else{
                            Mensaje(0,Resultado["Contenido"]);
                        }
                        
                    },
                    error: function (jqXHR, textStatus, errorThrown ){Mensaje(0,textStatus);}
                });
            }
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Cargar_Excel_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Error = ""
        Info_Datos = json.loads(str(Datos["Info"]))
        Max_Nulls = 5
        Cuantos_Null = 0
        wb_obj = openpyxl.load_workbook("I:/Portal_File/"+str(Info_Datos["Excel"][0]))
        Shell = wb_obj.active
        Datos = []
        Rows = 4
        while Cuantos_Null < Max_Nulls:
            Zona = Shell["A"+str(Rows)].value
            if Zona is not None:
                Cuantos_Null = 0
                Datos.append({
                    "Zone":Shell["A"+str(Rows)].value,
                    "Last Scan":Shell["B"+str(Rows)].value,
                    "Packing Slip":Shell["C"+str(Rows)].value,
                    "Part":Shell["D"+str(Rows)].value,
                    "Serial":Shell["E"+str(Rows)].value,
                    "Tare Serial":Shell["F"+str(Rows)].value,
                    "Loaded":Shell["G"+str(Rows)].value,
                    "Pack Qty":Shell["H"+str(Rows)].value,
                    "Route Date":Shell["I"+str(Rows)].value,
                    "Route":Shell["J"+str(Rows)].value,
                    "Plant":Shell["K"+str(Rows)].value,
                    "Weight":Shell["L"+str(Rows)].value,
                    "Supplier Code":Shell["M"+str(Rows)].value,
                    "Supplier P/S":Shell["N"+str(Rows)].value,
                    "Cut Time":Shell["O"+str(Rows)].value
                })
            else:
                Cuantos_Null += 1
            Rows += 1
        DB.Instruccion("UPDATE ilc_saltillo.cwms_zonas SET cwz_actual = 0 WHERE cwz_actual = 1")
        Qu = []
        for D in Datos:
            Qu.append("""('"""+str(D["Zone"])+"""','"""+str(D["Last Scan"].strftime("%Y-%m-%d %H:%M:%S"))+"""','"""+str(D["Packing Slip"])+"""','"""+str(D["Part"])+"""','"""+str(D["Serial"])+"""','"""+str(D["Tare Serial"])+"""','"""+str(D["Loaded"])+"""','"""+str(D["Pack Qty"]).replace(" ","").replace(",","")+"""','"""+str(D["Route Date"])+"""','"""+str(D["Route"])+"""','"""+str(D["Plant"])+"""','"""+str(D["Weight"]).replace(" ","").replace(",","")+"""','"""+str(D["Supplier Code"])+"""','"""+str(D["Supplier P/S"])+"""','"""+str(D["Cut Time"].strftime("%Y-%m-%d %H:%M:%S"))+"""',NOW())""")
        Error = DB.Instruccion("""INSERT INTO ilc_saltillo.cwms_zonas (cwz_zona,cwz_last_scan,cwz_paking_slip,cwz_part,cwz_serial,cwz_tare_serial,cwz_loaded,cwz_pack_qty,cwz_rout_date,cwz_route,cwz_plant,cwz_weight,cwz_supplier_code,cwz_supplier_p_s,cwz_cut_time,cwz_fecha_hora)  VALUES  """ + ",".join(Qu))
        if Error == "":
            DB.Instruccion("UPDATE servicios.cvariables_general SET ultima_zona_wms_ocs = NOW()")
            os.remove("I:/Portal_File/"+str(Info_Datos["Excel"][0]))
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] = str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)


form = cgi.FieldStorage()
if "Fun" in form:
    Encriptar = LibDM_2023.Encriptar()
    Funcion = form.getvalue("Fun")
    Par = {}
    for P in form.keys():
        if str(P) != "Fun":
            Par[str(P)] = form.getvalue(str(P))
    ID_User = Encriptar.Get_IDUsuario(cookievalue['K_Portal_U'].value)
    Par["ID_User"] = ID_User
    Par["Key_U"] = str(cookievalue['K_Portal_U'].value)
    locals()[Funcion](Par)
else:
    try:
        if cookievalue is not None:
            Encriptar = LibDM_2023.Encriptar()
            Datos = {}
            Datos["ID_User"] = Encriptar.Get_IDUsuario(cookievalue['K_Portal_U'].value)
            Datos["Key_U"] = str(cookievalue['K_Portal_U'].value)
            Inicio(Datos)            
        else:
            Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
            Cur += """
            <script>
                window.location= "../";
            </script>
            """
            print (Cur)
    except:
        Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
        Cur += """
        <script>
            window.location= "../";
        </script>
        """
        print (Cur)