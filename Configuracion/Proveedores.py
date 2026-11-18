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
Dir_Raiz = "./Configuracion/Proveedores.py"

BD_Nombre_Yard = "ulc_saltillo"
DB_Nombre_Negocio = "ulc_saltillo_volvo"
Negocio = "VOLVO"
Bandera_Dock = "VOLVO"

def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Proveedores = DB.Get_Dato("SELECT * FROM "+str(DB_Nombre_Negocio)+".cproveedores WHERE crilcpr_activo = 1")
        Patio = []
        for Proveedor in Proveedores:
            Opciones = "<div class='text-center'><div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Opciones += "<button class='btn btn-sm btn-warning p-0 ps-1 pe-1' onclick='Modificar("+str(Proveedor["crilcpr_id"])+",\""+str(Proveedor["crilcpr_codigo"])+"\")'><i class='mdi mdi-pencil'></i></button>"
            Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Eliminar("+str(Proveedor["crilcpr_id"])+",\""+str(Proveedor["crilcpr_codigo"])+"\")'><i class='mdi mdi-trash-can'></i></button>"
            Opciones += "</div></div>"
            Patio.append({"Codigo":str(Proveedor["crilcpr_codigo"]),"Nombre":str(Proveedor["crilcpr_nombre"]),"Opciones":Opciones})


        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi-card-bulleted'></i> Administrador de Proveedores</div>"
        Resultado["Contenido"] += """
        <div class='container'>
        <div class='row'>
            <div class='col-12'>
                <div class='text-end mb-1'> <button class='btn btn-success' onclick='Nueva_Proveedor()'><i class='mdi mdi-plus'></i> Nuevo Proveedor</button> </div>
                <div id='Tabla_Proveedores' class='border border-dark bg-dark-subtle'></div>
                <script>
                    delete Tabla_Proveedores;
                    var Tabla_Proveedores = new Tabulator("#Tabla_Proveedores", {
                        minHeight:800,
                        layout:"fitColumns",
                        data:"""+str(Patio)+""",
                        columns:[
                            {field:"Codigo","title":"Código",widthGrow:1,headerFilter:"input"},
                            {field:"Nombre","title":"Nombre",headerFilter:"input",widthGrow:10},
                            {field:"Opciones","title":"Op",formatter:"html",widthGrow:1}
                        ]
                    });
                    Tabla_Proveedores.on("tableBuilt", function(){ 
                        setTimeout(() => {
                            Tabla_Proveedores.clearFilter();
                            Tabla_Proveedores.redraw();
                        }, 100);
                    });
                </script>
            </div>
        </div>
        </div>
        <script>
            function Nueva_Proveedor(){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-plus'></i> Nuevo Proveedor");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Nueva_Proveedor"};
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
            function Modificar(ID,Codigo){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-pencil'></i>" + Codigo);
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Modificar","ID":ID};
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
            function Eliminar(ID,Codigo){
                Swal.fire({
                    title: '¿Estás seguro de eliminar el proveedor ['+Codigo+']?',
                    buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Si",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                    customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                    preConfirm: () => {
                        
                        Mostrar_Ventana_Cargando(false);
                            var parametros = {"Fun":"Eliminar","ID":ID};
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
                })
            }
        </script>
        """
    except:
         Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Nueva_Proveedor(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Formulario = {"Col":"12", "Campos": [],"Clase": "Alta_Proveedor" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Codigo","titulo":"Código","Requerido":1,"min":1,"max":30,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Name","titulo":"Nombre","Requerido":1,"min":1,"max":150,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Direccion","titulo":"Dirección","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Ciudad","titulo":"Ciudad","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Estado","titulo":"Estado","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"CP","titulo":"CP","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"multitexto","campo":"Correos","titulo":"E-mail(s)","Requerido":0,"min":0,"max":2000,"valor":""})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Nueva_Proveedor_Guardar()'><i class='mdi mdi-floppy'></i> Save</button></div>
        """
        Resultado["Contenido"] += """
        <script>
            function Nueva_Proveedor_Guardar(){
                var Info = Dame_Formulario(".Alta_Proveedor",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Nueva_Proveedor_Guardar","Info":JSON.stringify(Info)};
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
def Nueva_Proveedor_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Datos = json.loads(str(Datos["Info"]))
        Existe = DB.Get_Dato("SELECT * FROM "+str(DB_Nombre_Negocio)+".cproveedores WHERE crilcpr_codigo = '"+str(Info_Datos["Codigo"])+"'  ")
        if len(Existe) == 0:
            Error = DB.Instruccion("""
            INSERT INTO """+str(DB_Nombre_Negocio)+""".cproveedores 
            (crilcpr_codigo,crilcpr_nombre,crilcpr_direccion,crilcpr_ciudad,crilcpr_estado,crilcpr_codigopostal,crilcpr_email_s)
            VALUES
            ('"""+str(Info_Datos["Codigo"])+"""','"""+str(Info_Datos["Name"])+"""','"""+str(Info_Datos["Direccion"])+"""','"""+str(Info_Datos["Ciudad"])+"""','"""+str(Info_Datos["Estado"])+"""','"""+str(Info_Datos["CP"])+"""','"""+str(Info_Datos["Correos"])+"""')
            """)
            if Error == "":
                Resultado["Estado"] = 1
            else:
                Resultado["Contenido"] += str(Error)
        else:
            Resultado["Contenido"] += "Este proveedor ya existe!"
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Modificar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Actual = DB.Get_Dato("SELECT * FROM "+str(DB_Nombre_Negocio)+".cproveedores WHERE crilcpr_id = '"+str(Datos["ID"])+"'")[0]
        for k in Info_Actual.keys():
            if Info_Actual[k] is None:
                Info_Actual[k] = ""
        Formulario = {"Col":"12", "Campos": [],"Clase": "Alta_Proveedor" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Codigo","titulo":"Código","Requerido":1,"min":1,"max":30,"valor":str(Info_Actual["crilcpr_codigo"]),"editable":False})
        Formulario["Campos"].append({"tipo":"texto","campo":"Name","titulo":"Nombre","Requerido":1,"min":1,"max":150,"valor":str(Info_Actual["crilcpr_nombre"]),"editable":False})
        Formulario["Campos"].append({"tipo":"texto","campo":"Direccion","titulo":"Dirección","Requerido":0,"min":0,"max":500,"valor":str(Info_Actual["crilcpr_direccion"])})
        Formulario["Campos"].append({"tipo":"texto","campo":"Ciudad","titulo":"Ciudad","Requerido":0,"min":0,"max":500,"valor":str(Info_Actual["crilcpr_ciudad"])})
        Formulario["Campos"].append({"tipo":"texto","campo":"Estado","titulo":"Estado","Requerido":0,"min":0,"max":500,"valor":str(Info_Actual["crilcpr_estado"])})
        Formulario["Campos"].append({"tipo":"texto","campo":"CP","titulo":"CP","Requerido":0,"min":0,"max":500,"valor":str(Info_Actual["crilcpr_codigopostal"])})
        Formulario["Campos"].append({"tipo":"multitexto","campo":"Correos","titulo":"E-mail(s)","Requerido":0,"min":0,"max":2000,"valor":str(Info_Actual["crilcpr_email_s"])})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Modificar_Guardar("""+str(Datos["ID"])+""")'><i class='mdi mdi-floppy'></i> Save</button></div>
        """
        Resultado["Contenido"] += """
        <script>
            function Modificar_Guardar(ID){
                var Info = Dame_Formulario(".Alta_Proveedor",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Modificar_Guardar","Info":JSON.stringify(Info),"ID":ID};
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
def Modificar_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Info_Datos = json.loads(str(Datos["Info"]))
        Error = DB.Instruccion("""
        UPDATE """+str(DB_Nombre_Negocio)+""".cproveedores SET
        crilcpr_codigo = '"""+str(Info_Datos["Codigo"])+"""',
        crilcpr_nombre = '"""+str(Info_Datos["Name"])+"""',
        crilcpr_direccion = '"""+str(Info_Datos["Direccion"])+"""',
        crilcpr_ciudad = '"""+str(Info_Datos["Ciudad"])+"""',
        crilcpr_estado = '"""+str(Info_Datos["Estado"])+"""',
        crilcpr_codigopostal = '"""+str(Info_Datos["CP"])+"""',
        crilcpr_email_s = '"""+str(Info_Datos["Correos"])+"""'
        WHERE crilcpr_id = '"""+str(Datos["ID"])+"""'
        """)
        if Error == "":
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Eliminar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    Error = ""
    try:
        Error = DB.Instruccion("UPDATE "+str(DB_Nombre_Negocio)+".cproveedores SET crilcpr_activo = 0 WHERE crilcpr_id = '"+str(Datos["ID"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
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

