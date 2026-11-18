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
Dir_Raiz = "./EDI/edi.py"

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

        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi-feature-search'></i> EDI</div>"
        Resultado["Contenido"] += """
        <hr>
        <div class='container'>
            <div class='text-end pe-1 pt-1'><small class='link-primary' style='cursor:pointer' onclick='Cargar();'>Actualizar <i class='mdi mdi-refresh'></i></small></div>
        """
        Fechas = str((datetime.now()-timedelta(days=3)).strftime("%Y-%m-%d"))+" to "+str(datetime.now().strftime("%Y-%m-%d"))
        Formulario = {"Col":"12", "Campos": [],"Clase": "Buscar" }
        Formulario["Campos"].append({"tipo":"fecha-rango","campo":"Fecha","titulo":"Fechas","Requerido":0,"Col":12,"valor":str(Fechas),"editable":True})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))

        Resultado["Contenido"] += """
            <div class='w-100 text-center mt-2'><button class='btn btn-primary w-75' onclick='Cargar()'><i class='mdi mdi-cloud-download'></i> Buscar</button></div>
            <hr>
            <div id='Res'></div>
        </div>
        """
        

        Resultado["Contenido"] += """
        </div>
        <script>
            function Cargar(){
                var Info = Dame_Formulario(".Buscar",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Cargar","Info":JSON.stringify(Info)};
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
            $(document).ready(function() {
                Cargar();
            });
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
def Cargar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info = json.loads(Datos["Info"])
        EDIs = []
        for EDI in DB.Get_Dato("select EDI.GIR,EDI.cfecha_hora_procesamiento,EDI.cimpreso  from ulc_saltillo_volvo.edi EDI where EDI.cfecha_hora_procesamiento >= '"+str(Info["Fecha"][0])+" 00:00:00' and EDI.cfecha_hora_procesamiento <= '"+str(Info["Fecha"][1])+" 23:59:59'  group by EDI.GIR "):
            Op = "<i style='cursor:pointer' onclick='Ver_Info(\""+str(EDI["GIR"])+"\")' class='text-dark mdi mdi-eye'></i>"
            EDIs.append({"GIR":str(EDI["GIR"]),"Procesado":str(EDI["cfecha_hora_procesamiento"]),"Impreso":str(EDI["cimpreso"]),"Op":str(Op)})
        Resultado["Contenido"] += """
        <div id='EDI_Tabla' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete EDI_Tabla;
            var EDI_Tabla = new Tabulator("#EDI_Tabla", {
                height:800,
                layout:"fitColumns",
                data:"""+str(EDIs)+""",
                columns:[
                    {title: 'GIR', field: 'GIR',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Procesado', field: 'Procesado',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Impreso', field: 'Impreso',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: '', field: 'Op',formatter:"html",hozAlign:"center",headerVertical:true,width:1,headerSort:false,download:false},
                ],
                initialSort:[
                    {column:"ID_Contenedor", dir:"asc"},
                ],
                selectable:false
            });
            document.getElementById("download-xlsx-detalle-clips_error").addEventListener("click", function(){
                Clips_Error.download("xlsx", 'Tabla.xlsx', {sheetName:"My Data",compress:false});
            });

            function Ver_Info(GIR){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-eyer'></i> "+GIR);
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-lg');
                var parametros = {"Fun":"Ver_Info","GIR":GIR};
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
        </script>
        """
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)

def Ver_Info(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        EDIs_Detalles = []
        for EDI in DB.Get_Dato("SELECT * FROM  ulc_saltillo_volvo.edi WHERE GIR = '"+str(Datos["GIR"])+"'"):
            EDIs_Detalles.append({"SEQ":str(EDI["SEQ"]).zfill(8),"GIR":str(EDI["GIR"]),"Parte":str(EDI["Parte"]),"Descripcion":str(EDI["Descripcion"]),"Cantidad":str(EDI["Cantidad"]),"Posicion":str(EDI["Posicion"])})
        Resultado["Contenido"] += """
        <div class='text-end mb-1'> <button class='btn btn-sm btn-dark' onclick='Imprimir(\""""+str(Datos["GIR"])+"""\")'><i class='mdi mdi-printer'></i> Imprimir</button></div>
        <div id='EDIs_Detalles' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete EDIs_Detalles;
            var EDIs_Detalles = new Tabulator("#EDIs_Detalles", {
                height:300,
                layout:"fitColumns",
                data:"""+str(EDIs_Detalles)+""",
                columns:[
                    {title: 'ID', field: 'ID',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html",download:false},
                    {title: 'SEQ', field: 'SEQ',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'GIR', field: 'GIR',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Parte', field: 'Parte',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Descripcion', field: 'Descripcion',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Cantidad', field: 'Cantidad',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"},
                    {title: 'Posicion', field: 'Posicion',headerFilter: 'input',vertAlign:"middle",hozAlign:"center",formatter:"html"}
                ],
                initialSort:[
                    {column:"ID_Contenedor", dir:"asc"},
                ],
                selectable:false
            });
            function Imprimir(GIR){
                Mostrar_Ventana_Cargando(false);
                var parametros = {"Fun":"Imprimir","GIR":GIR};
                $.ajax({data:  parametros,url:\""""+str(Dir_Raiz)+"""\",type:  "post",
                    success:  function (response)
                    {
                        var Resultado = JSON.parse(response);
                        if(Resultado["Estado"] == 1){
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

def Imprimir(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        DB.Instruccion("UPDATE ulc_saltillo_volvo.edi SET cimpreso = NULL WHERE GIR = '"+str(Datos["GIR"])+"'")
        Resultado["Estado"] = 1
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

