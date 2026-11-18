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
Dir_Raiz = "./Reportes/daily.py"

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

        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi mdi-book'></i> Reporte Diario</div>"
        Formulario = {"Col":"12", "Campos": [],"Clase": "Buscar" }
        #Formulario["Campos"].append({"tipo":"texto","campo":"Container","titulo":"Container","Requerido":1,"min":1,"max":150,"valor":"","Col":6})
        Formulario["Campos"].append({"tipo":"fecha","campo":"Fecha de Rura","titulo":"Fecha de Ingreso","Requerido":1,"Col":12,"valor":"","editable":True})
        
        Resultado["Contenido"] += "<div class='container'>"
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <div class='w-100 text-center mt-2'><button class='btn btn-primary w-75' onclick='Cargar_Reporte()'><i class='mdi mdi-cloud-download'></i> Descargar</button></div>
        <hr>
        <div id='Res'></div>
        """
        Resultado["Contenido"] += """
        </div>
        <script>
            function Cargar_Reporte(){
                var Info = Dame_Formulario(".Buscar",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Cargar_Reporte","Info":JSON.stringify(Info)};
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
def Cargar_Reporte(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Datos_Info = json.loads(str(Datos["Info"]))
        Cajas = []
        Cajas_2 = []
        Outbound = []
        Cajas_DB = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_negocio = '"+str(Negocio)+"'  AND DATE(cc_fecha_hora) = '"+str(Datos_Info["Fecha de Rura"])+"'")
        H_Str = []
        for C in Cajas_DB:
            H_Str.append("'"+str(C["cc_id"])+"'")
        Historial = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento WHERE cch_master IN ("+str(','.join(H_Str))+")")
        
        for Inbound in DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento where cch_negocio = '"+str(Negocio)+"' AND cch_tipo_actual = 'Inbound' and cch_movimiento  = 'ASIGNAR' and DATE(cch_fecha_hora) = '"+str(Datos_Info["Fecha de Rura"])+"'"):
            Tipo_Route = ""
            Route = ""
            Dock = ""
            Opciones = "<div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Ver_Historico("+str(Inbound["cch_master"])+",\""+str(Inbound["cch_contenedor"])+"\")'><i class='mdi mdi-history'></i></button>"
            Opciones += "</div>"
            Info_H = json.loads(str(Inbound["cch_informacion_actual"]))
            Carrier = Info_H["Carrier"]
            Estado = "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(Inbound["cch_tipo_actual"])+");Color:var(--Color_"+str(Inbound["cch_tipo_actual"])+");'>__</span>"
            try:
                Tipo = Info_H[Inbound["cch_tipo_actual"]]
                Tipo_Route = Tipo
                Estado += str(Tipo)
                Tipo_1 = Info_H[Tipo]
                Route = Tipo_1
                Estado += "/"+str(Tipo_1)
                Tipo_2 = Info_H[Tipo_1]
                Dock = Tipo_2
                Estado += "/"+str(Tipo_2)
            except:
                pass
            Cajas.append({"Contenedor":str(Inbound["cch_contenedor"]),"In Date":Inbound["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Route Type":Tipo_Route,"Route":Route,"Dock":Dock,"In Carrier":Carrier,'Opciones':Opciones})
        

        for Outbound in DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento where cch_negocio = '"+str(Negocio)+"' AND cch_tipo_actual = 'Outbound' and cch_movimiento  = 'SALIDA' and DATE(cch_fecha_hora) = '"+str(Datos_Info["Fecha de Rura"])+"'"):
            Tipo_Route = ""
            Route = ""
            Dock = ""
            Opciones = "<div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Ver_Historico("+str(Outbound["cch_master"])+",\""+str(Outbound["cch_contenedor"])+"\")'><i class='mdi mdi-history'></i></button>"
            Opciones += "</div>"
            Info_H = json.loads(str(Outbound["cch_informacion_actual"]))
            Carrier = Info_H["Carrier"]
            Estado = "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(Outbound["cch_tipo_actual"])+");Color:var(--Color_"+str(Outbound["cch_tipo_actual"])+");'>__</span>"
            try:
                Tipo = Info_H[Outbound["cch_tipo_actual"]]
                Tipo_Route = Tipo
                Estado += str(Tipo)
                Tipo_1 = Info_H[Tipo]
                Route = Tipo_1
                Estado += "/"+str(Tipo_1)
                Tipo_2 = Info_H[Tipo_1]
                Dock = Tipo_2
                Estado += "/"+str(Tipo_2)
            except:
                pass
            Cajas_2.append({"Contenedor":str(Outbound["cch_contenedor"]),"Out Date":Outbound["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Route Type":Tipo_Route,"Route":Route,"Dock":Dock,"Out Carrier":Carrier,'Opciones':Opciones})

            

                    
        Resultado["Contenido"] += """
        <h1>Inbound</h1>
        <button class='btn btn-primary mb-1' id="download-xlsx-Inbound"><i class='mdi mdi-microsoft-excel'></i> Descargar Excel</button>
        <div id='Tabla_Reporte' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Reporte;
            var Tabla_Reporte = new Tabulator("#Tabla_Reporte", {
                minHeight:500,
                layout:"fitColumns",
                data:"""+str(Cajas)+""",
                columns:[
                    {field:"Contenedor","title":"Contenedor"},
                    {field:"In Date","title":"Fecha de Entrada"},
                    {field:"Route Type","title":"Tipo de Ruta",formatter:"html"},
                    {field:"Route","title":"Ruta",formatter:"html"},
                    {field:"Dock","title":"Dock",formatter:"html"},
                    {field:"In Carrier","title":"Carrier de Entrada"},
                    {field:"Opciones","title":"Op",formatter:"html",download: false}
                ]
            });
            Tabla_Reporte.on("tableBuilt", function(){ 
                setTimeout(() => {
                    Tabla_Reporte.clearFilter();
                    Tabla_Reporte.redraw();
                }, 100);
            });
            document.getElementById("download-xlsx-Inbound").addEventListener("click", function(){
                Tabla_Reporte.download("xlsx", "Inbound.xlsx", {sheetName:"My Data"});
            });
        </script>
        """
        Resultado["Contenido"] += """
        <h1>Outbound</h1>
        <button class='btn btn-primary mb-1' id="download-xlsx-Outbound"><i class='mdi mdi-microsoft-excel'></i> Descargar Excel</button>
        <div id='Tabla_Reporte_2' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Reporte_2;
            var Tabla_Reporte_2 = new Tabulator("#Tabla_Reporte_2", {
                minHeight:500,
                layout:"fitColumns",
                data:"""+str(Cajas_2)+""",
                columns:[
                    {field:"Contenedor","title":"Contenedor"},
                    {field:"Out Date","title":"Fecha de Salida"},
                    {field:"Route Type","title":"Tipo de Ruta",formatter:"html"},
                    {field:"Route","title":"Ruta",formatter:"html"},
                    {field:"Dock","title":"Dock",formatter:"html"},
                    {field:"Out Carrier","title":"Carrier de Salida"},
                    {field:"Opciones","title":"Op",formatter:"html",download: false}
                ]
            });
            Tabla_Reporte_2.on("tableBuilt", function(){ 
                setTimeout(() => {
                    Tabla_Reporte_2.clearFilter();
                    Tabla_Reporte_2.redraw();
                }, 100);
            });
            document.getElementById("download-xlsx-Outbound").addEventListener("click", function(){
                Tabla_Reporte_2.download("xlsx", "Outbound.xlsx", {sheetName:"My Data"});
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

