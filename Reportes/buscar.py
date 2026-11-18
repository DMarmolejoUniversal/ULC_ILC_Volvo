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
import random
import string

sys.path.insert(1,r'C:\inetpub\wwwroot\Portal2\Componentes_2023')
import LibDM_2023
locale.setlocale( locale.LC_ALL, '' )
if "HTTP_COOKIE" in os.environ:
    cookievalue = cookies.SimpleCookie(os.environ["HTTP_COOKIE"])
else:
    cookievalue = None
Dir_Raiz = "./Reportes/buscar.py"

BD_Nombre_Yard = "ulc_saltillo"
DB_Nombre_Negocio = "ulc_saltillo_volvo"
Negocio = "VOLVO"
Bandera_Dock = "VOLVO"
    
def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:

        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi-feature-search'></i> Buscar Contenedor</div>"
        Formulario = {"Col":"12", "Campos": [],"Clase": "Buscar" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Container","titulo":"Contenedor","Requerido":1,"min":1,"max":150,"valor":"","Col":4})
        Formulario["Campos"].append({"tipo":"fecha","campo":"Fecha de Rura","titulo":"Fecha de Ingreso","Requerido":0,"Col":4,"valor":"","editable":True})
        Formulario["Campos"].append({"tipo":"seleccion","campo":"Carrier","titulo":"Carrier","Requerido":0,"Tipo_Opciones":"Query","Opciones":"SELECT cca_nombre as Valor, cca_nombre as Texto FROM "+str(BD_Nombre_Yard)+".ccarrier ","valor":"","Col":4})
        
        Resultado["Contenido"] += "<div class='container'>"
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <div class='w-100 text-center mt-2'><button class='btn btn-primary w-75' onclick='Buscar_Caja()'><i class='mdi mdi-cloud-download'></i> Buscar</button></div>
        <hr>
        <div id='Res'></div>
        """
        Resultado["Contenido"] += """
        </div>
        <script>
            function Buscar_Caja(){
                var Info = Dame_Formulario(".Buscar",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Buscar_Caja","Info":JSON.stringify(Info)};
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
        </script>
        """
    except:
         Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Buscar_Caja(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        #Resultado["Contenido"] += str(Datos)
        Datos_Info = json.loads(str(Datos["Info"]))
        Cajas = []
        if str(Datos_Info["Fecha de Rura"]) != "":
            if Datos_Info["Carrier"] == "-1":
                Cajas_DB = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_negocio = '"+str(Negocio)+"' AND cc_contenedor like '%"+str(Datos_Info["Container"])+"%' AND DATE(cc_fecha_hora) = '"+str(Datos_Info["Fecha de Rura"])+"'")
            else:
                Cajas_DB = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_negocio = '"+str(Negocio)+"' AND cc_contenedor like '%"+str(Datos_Info["Container"])+"%' AND DATE(cc_fecha_hora) = '"+str(Datos_Info["Fecha de Rura"])+"' and cc_informacion_actual like '%"+str(Datos_Info["Carrier"])+"%'")
        else:
            if Datos_Info["Carrier"] == "-1":
                Cajas_DB = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_negocio = '"+str(Negocio)+"' AND cc_contenedor like '%"+str(Datos_Info["Container"])+"%'")
            else:
                Cajas_DB = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_negocio = '"+str(Negocio)+"' AND cc_contenedor like '%"+str(Datos_Info["Container"])+"%' and cc_informacion_actual like '%"+str(Datos_Info["Carrier"])+"%'")
        H_Str = []
        for C in Cajas_DB:
            H_Str.append("'"+str(C["cc_id"])+"'")
        Historial = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento WHERE cch_master IN ("+str(','.join(H_Str))+")")
        
        for Caja in Cajas_DB:
            Historial_Aqui = []
            for H in Historial:
                if int(H["cch_master"]) == int(Caja["cc_id"]):
                    Historial_Aqui.append(H)
            Info_Actual = json.loads(str(Caja["cc_informacion_actual"]))


            Sello = ""
            if "Sello Proveedor" in Info_Actual.keys():
                Sello = "<span style='color:blue'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Proveedor"])
            if "Sello Temporal" in Info_Actual.keys():
                Sello = "<span style='color:#7a7a7a'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Temporal"])
            if "Sello Blanco" in Info_Actual.keys():
                Sello = "<span><i class='mdi mdi-label-outline'></i></span> "+str(Info_Actual["Sello Temporal"])
            if "Sello Rojo" in Info_Actual.keys():
                Sello = "<span style='color:#ff0000'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Temporal"])

            Ruta = ""
            Ruta += "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(Caja["cc_tipo_actual"])+");Color:var(--Color_"+str(Caja["cc_tipo_actual"])+");'>__</span>"
            try:
                Tipo = Info_Actual[Caja["cc_tipo_actual"]]
                Ruta += str(Tipo)
                Tipo_1 = Info_Actual[Tipo]
                Ruta += "/"+str(Tipo_1)
                Tipo_2 = Info_Actual[Tipo_1]
                Ruta += "/"+str(Tipo_2)
            except:
                pass
            
            if "Fecha de Salida" in Info_Actual.keys():
                Ruta += " <"+str(Info_Actual["Fecha de Salida"])+">"

            Estado_IN = "Not assigned"
            Carrier_IN = Info_Actual["Carrier"]
            Fecha_Out = ""
            Carrier_Out = ""
            Estado_Out = ""
            Sello_de_Salida = ""
            Opciones = "<div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Ver_Historico("+str(Caja["cc_id"])+",\""+str(Caja["cc_contenedor"])+"\")'><i class='mdi mdi-history'></i></button>"
            Opciones += "</div>"
            
            for HQ in Historial_Aqui:
                if str(HQ["cch_movimiento"]) == "ASIGNAR":
                    Info_H = json.loads(str(HQ["cch_informacion_actual"]))
                    Carrier_IN = Info_H["Carrier"]
                    Estado_IN = "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(HQ["cch_tipo_actual"])+");Color:var(--Color_"+str(HQ["cch_tipo_actual"])+");'>__</span>"
                    try:
                        Tipo = Info_H[HQ["cch_tipo_actual"]]
                        Estado_IN += str(Tipo)
                        Tipo_1 = Info_H[Tipo]
                        Estado_IN += "/"+str(Tipo_1)
                        Tipo_2 = Info_H[Tipo_1]
                        Estado_IN += "/"+str(Tipo_2)
                    except:
                        pass
                if str(HQ["cch_movimiento"]) == "SALIDA":
                    Info_H = json.loads(str(HQ["cch_informacion_actual"]))
                    Fecha_Out = str(HQ["cch_fecha_hora"])
                    Carrier_Out = Info_H["Carrier"]
                    Sello_de_Salida = ""
                    if "Sello Proveedor" in Info_H.keys():
                        Sello_de_Salida = "<span style='color:blue'><i class='mdi mdi-label'></i></span> "+str(Info_H["Sello Proveedor"])
                    if "Sello Temporal" in Info_H.keys():
                        Sello_de_Salida = "<span style='color:#7a7a7a'><i class='mdi mdi-label'></i></span> "+str(Info_H["Sello Temporal"])
                    if "Sello Blanco" in Info_H.keys():
                        Sello_de_Salida = "<span><i class='mdi mdi-label-outline'></i></span> "+str(Info_H["Sello Temporal"])
                    if "Sello Rojo" in Info_H.keys():
                        Sello_de_Salida = "<span style='color:#ff0000'><i class='mdi mdi-label'></i></span> "+str(Info_H["Sello Temporal"])
                if str(HQ["cch_movimiento"]) == "ELIMINADO":
                    Ruta = "<span class='text-danger'>ELIMINADO</span>"



            Cajas.append({"Contenedor":str(Caja["cc_contenedor"]),"In Date":Caja["cc_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Ultimo Estado":Ruta,"In Carrier":Carrier_IN,"Out Date":Fecha_Out,"Out Carrier":Carrier_Out,"Sello de Salida":Sello_de_Salida,'Opciones':Opciones})

            pass
        Resultado["Contenido"] += """
        <div id='Tabla_Reporte' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Reporte;
            var Tabla_Reporte = new Tabulator("#Tabla_Reporte", {
                minHeight:500,
                layout:"fitColumns",
                data:"""+str(Cajas)+""",
                columns:[
                    {field:"Contenedor","title":"Contenedor"},
                    {field:"In Date","title":"Fecha de Ingreso"},
                    {field:"In Carrier","title":"Carrier de Ingreso"},
                    {field:"Ultimo Estado","title":"Ultimo Estado",formatter:"html"},
                    {field:"Out Date","title":"Fecha de Salida"},
                    {field:"Out Carrier","title":"Carrier de Salida"},
                    {field:"Sello de Salida","title":"Sello de Salida",formatter:"html"},
                    {field:"Opciones","title":"Op",formatter:"html",width:20}
                ]
            });
            Tabla_Reporte.on("tableBuilt", function(){ 
                setTimeout(() => {
                    Tabla_Reporte.clearFilter();
                    Tabla_Reporte.redraw();
                }, 100);
            });
        </script>
        """
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

