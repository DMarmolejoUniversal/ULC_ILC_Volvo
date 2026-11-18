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
Dir_Raiz = "./Configuracion/Docks.py"

BD_Nombre_Yard = "ulc_saltillo"
DB_Nombre_Negocio = "ulc_saltillo_volvo"
Negocio = "VOLVO"
Bandera_Dock = "VOLVO"

def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Docks = []
        for Dock in DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".cdock_asignacion"):
            if Dock["cd_etiqueta"] is None:
                Dock["cd_etiqueta"] = ""
            if int(Dock["cd_activo"]) == 1:
                if str(Dock["cd_ub"]) == str(Bandera_Dock):
                    Opciones = "<div class='text-center'><div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
                    Opciones += "<button class='btn btn-sm btn-warning p-0 ps-1 pe-1' onclick='Modificar("+str(Dock["cd_dock"])+")'><i class='mdi mdi-pencil'></i></button>"
                    Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Eliminar("+str(Dock["cd_dock"])+")'><i class='mdi mdi-trash-can'></i></button>"
                    Opciones += "</div></div>"
                elif Dock["cd_ub"] is None:
                    Opciones = "<div class='text-center'><div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
                    Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Asignar("+str(Dock["cd_dock"])+")'><i class='mdi mdi-hand-front-right'></i></button>"
                    Opciones += "</div></div>"
                else:
                    Opciones = ""
            else:
                Opciones = "<span class='ms-1 me-1'><i class='mdi mdi-lock'></i></span>"
            if Dock["cd_ub"] is None:
                Dock["cd_ub"] = ""
            Docks.append({"Dock":str(Dock["cd_dock"]),"Label":str(Dock["cd_etiqueta"]),"Opciones":Opciones,"UB":str(Dock["cd_ub"]),"Activo":int(Dock["cd_activo"])})


        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi mdi-sign-caution'></i> Administrador de Docks</div>"
        Resultado["Contenido"] += """
        <div class='container'>
        <div class='row'>
            <div class='col-12'>
                <div id='Tabla_Proveedores' class='border border-dark bg-dark-subtle'></div>
                <script>
                    delete Tabla_Proveedores;
                    var Tabla_Proveedores = new Tabulator("#Tabla_Proveedores", {
                        minHeight:800,
                        layout:"fitColumns",
                        data:"""+str(Docks)+""",
                        rowFormatter:function(row){
                            if(row.getData().Activo == 0){
                                 if(row.getPosition()%2)
                                        row.getElement().style.backgroundColor = "#bfbfbf";
                                    else
                                        row.getElement().style.backgroundColor = "#aeaeae";
                            }
                            else{
                                if(row.getData().UB == "XDOCK NISSAN"){
                                    if(row.getPosition()%2)
                                        row.getElement().style.backgroundColor = "#dcf8d7";
                                    else
                                        row.getElement().style.backgroundColor = "#bbf7af";
                                }
                                if(row.getData().UB != "" && row.getData().UB != "XDOCK NISSAN"){
                                    if(row.getPosition()%2)
                                        row.getElement().style.backgroundColor = "#f8d7d7";
                                    else
                                        row.getElement().style.backgroundColor = "#f7afaf";
                                }
                            }
                            
                            
                        },
                        columns:[
                            {field:"Dock","title":"DOCK",widthGrow:1,headerFilter:"input"},
                            {field:"Label","title":"Etiqueta",headerFilter:"input",widthGrow:10},
                            {field:"UB","title":"UB",headerFilter:"input",widthGrow:10},
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
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-plus'></i> New Suppliers");
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
            function Modificar(Dock){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-pencil'></i>" + Dock);
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm');
                var parametros = {"Fun":"Modificar","Dock":Dock};
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
            function Eliminar(Dock){
                Swal.fire({
                    title: '¿Estás seguro de liberar el dock ['+Dock+']?',
                    buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Yes",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                    customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                    preConfirm: () => {
                        
                        Mostrar_Ventana_Cargando(false);
                            var parametros = {"Fun":"Eliminar","Dock":Dock};
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
            function Asignar(Dock){
                Swal.fire({
                    title: '¿Estás seguro de asignar el dock ['+Dock+'] al """+str(Bandera_Dock)+"""?',
                    buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Si",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                    customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                    preConfirm: () => {
                        
                        Mostrar_Ventana_Cargando(false);
                            var parametros = {"Fun":"Asignar","Dock":Dock};
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
def Ver_Historico(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Historico = []
        Usuarios = DB.Get_Dato("SELECT USR.cusrid,EMP.Nombre FROM portal.cuser USR inner join linc.empleados EMP on USR.cusidempleado = EMP.IDEmpleado ")
        for H in DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento WHERE cch_master = '"+str(Datos["ID"])+"' ORDER BY cch_fecha_hora"):
            Info_Actual = json.loads(str(H["cch_informacion_actual"]))
            Usr_Aqui = None
            for U in Usuarios:
                if str(U["cusrid"]) == str(H["cch_usuario"]):
                    Usr_Aqui = U
                    break
            Tipo_Gen = "Not assigned" if H["cch_tipo_actual"] is None else H["cch_tipo_actual"]
            Usuario = str(H["cch_usuario"]) if Usr_Aqui is None else str(Usr_Aqui["Nombre"])
            Ubicacion = str(H["cch_ubicacion"]) if H["cch_dock"] is None else str(H["cch_ubicacion"]) + "["+str(H["cch_dock"])+"]"
            Archivos = ""
            Estado = ""
            Tipo = ""
            try:
                Tipo = Info_Actual[H["cch_tipo_actual"]]
                Estado += str(Tipo)
                Tipo_1 = Info_Actual[Tipo]
                Estado += "/"+str(Tipo_1)
                Tipo_2 = Info_Actual[Tipo_1]
                Estado += "/"+str(Tipo_2)
            except:
                pass

            if "Fotos" in Info_Actual.keys():
                for F in Info_Actual["Fotos"]:
                    Archivos += "<a href='http://10.4.7.219:8080/Portal_File/"+str(F)+"' target='_blank'><i class='mdi mdi-file-image'></i></a>"
            if "Archivos" in Info_Actual.keys():
                for F in Info_Actual["Archivos"]:
                    Archivos += "<a href='http://10.4.7.219:8080/Portal_File/"+str(F)+"' target='_blank'><i class='mdi mdi-file'></i></a>"
            
            Sello = ""
            if "Sello Proveedor" in Info_Actual.keys():
                Sello = "<span style='color:blue'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Proveedor"])
            if "Sello Temporal" in Info_Actual.keys():
                Sello = "<span style='color:#7a7a7a'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Temporal"])
            if "Sello Blanco" in Info_Actual.keys():
                Sello = "<span><i class='mdi mdi-label-outline'></i></span> "+str(Info_Actual["Sello Temporal"])
            if "Sello Rojo" in Info_Actual.keys():
                Sello = "<span style='color:#ff0000'><i class='mdi mdi-label'></i></span> "+str(Info_Actual["Sello Temporal"])

            Historico.append({"Fecha":H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Usuario":str(Usuario),"Ubicacion":str(Ubicacion),"Carrier":Info_Actual["Carrier"],"Tipo":Tipo_Gen,"Estado":Estado,"Movimiento":str(H["cch_movimiento"]),"Archivos":Archivos,"Sello":Sello})
        Resultado["Contenido"] += """
        <div id='Tabla_Historico' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Historico;
            var Tabla_Historico = new Tabulator("#Tabla_Historico", {
                minHeight:500,
                layout:"fitColumns",
                data:"""+str(Historico)+""",
                columns:[
                    {field:"Fecha","title":"Date"},
                    {field:"Usuario","title":"User"},
                    {field:"Ubicacion","title":"Location"},
                    {field:"Carrier","title":"Carrier"},
                    {field:"Tipo","title":"Type"},
                    {field:"Estado","title":"Status"},
                    {field:"Sello","title":"Seal",formatter:"html"},
                    {field:"Movimiento","title":"Step"},
                    {field:"Archivos","title":"Attached",formatter:"html"}
                ]
            });
            Tabla_Historico.on("tableBuilt", function(){ 
                setTimeout(() => {
                    Tabla_Historico.clearFilter();
                    Tabla_Historico.redraw();
                }, 100);
            });
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
        Formulario["Campos"].append({"tipo":"texto","campo":"Codigo","titulo":"Code","Requerido":1,"min":1,"max":30,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Name","titulo":"Name","Requerido":1,"min":1,"max":150,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Direccion","titulo":"Address","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Ciudad","titulo":"City","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"Estado","titulo":"State","Requerido":0,"min":0,"max":500,"valor":""})
        Formulario["Campos"].append({"tipo":"texto","campo":"CP","titulo":"Zip","Requerido":0,"min":0,"max":500,"valor":""})
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
        Info_Actual = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".cdock_asignacion WHERE cd_dock = '"+str(Datos["Dock"])+"'")[0]
        if Info_Actual["cd_etiqueta"] is None:
            Info_Actual["cd_etiqueta"] = ""
        Formulario = {"Col":"12", "Campos": [],"Clase": "Dock" }
        Formulario["Campos"].append({"tipo":"texto","campo":"Etiqueta","titulo":"Etiqueta","Requerido":0,"min":0,"max":30,"valor":str(Info_Actual["cd_etiqueta"])})
        Resultado["Contenido"] += str(Compartido_2023.Formulario(Formulario))
        Resultado["Contenido"] += """
        <hr>
        <div class='w-100 text-center'><button class='btn btn-success w-75' onclick='Modificar_Guardar("""+str(Datos["Dock"])+""")'><i class='mdi mdi-floppy'></i> Guardar</button></div>
        """
        Resultado["Contenido"] += """
        <script>
            function Modificar_Guardar(Dock){
                var Info = Dame_Formulario(".Dock",true);
                if(Info != null)
                {
                    Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Modificar_Guardar","Info":JSON.stringify(Info),"Dock":Dock};
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
        UPDATE """+str(BD_Nombre_Yard)+""".cdock_asignacion SET
        cd_etiqueta = '"""+str(Info_Datos["Etiqueta"])+"""'
        WHERE cd_dock = '"""+str(Datos["Dock"])+"""'
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
        Error = DB.Instruccion("UPDATE "+str(BD_Nombre_Yard)+".cdock_asignacion SET cd_etiqueta = null, cd_ub = null WHERE cd_dock = '"+str(Datos["Dock"])+"'")
        if Error == "":
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Asignar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    Error = ""
    try:
        Error = DB.Instruccion("UPDATE "+str(BD_Nombre_Yard)+".cdock_asignacion SET cd_etiqueta = null, cd_ub = '"+str(Bandera_Dock)+"' WHERE cd_dock = '"+str(Datos["Dock"])+"'")
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