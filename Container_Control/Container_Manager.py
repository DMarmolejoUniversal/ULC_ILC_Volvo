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
Dir_Raiz = "./Container_Control/Container_Manager.py"

BD_Nombre_Yard = "ulc_saltillo"
DB_Nombre_Negocio = "ulc_saltillo_volvo"
Bandera_Dock = "VOLVO"
Negocio = "VOLVO"


def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Contenedores_Patio = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas where cc_negocio = '"+str(Negocio)+"' and cc_activo = 1 AND cc_caja_danada = 0 AND cc_bloqueo_rg = 0 AND cc_ubicacion = 'Patio'")
        Contenedores_Dock = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas where cc_negocio = '"+str(Negocio)+"' and cc_activo = 1 AND cc_caja_danada = 0 AND cc_bloqueo_rg = 0 and cc_ubicacion = 'Dock'")
        Patio = []
        for Contenedor in Contenedores_Patio:
            Opciones = "<div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Estado = ""
            Info_Actual = json.loads(str(Contenedor["cc_informacion_actual"]))


            Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Ver_Historico("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-history'></i></button>"
            Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Eliminar_Caja("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-trash-can'></i></button>"
            if Contenedor["cc_bloquear"] == 0:
                Opciones += "<button class='btn btn-sm btn-warning p-0 ps-1 pe-1' onclick='Modificar("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-pencil'></i></button>"
                #Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Eliminar_Caja("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-trash-can'></i></button>"
            #if Contenedor["cc_qr_salida"] is None:
            Opciones += "<button class='btn btn-sm btn-info p-0 ps-1 pe-1' onclick='Regresar("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-cursor-move'></i></button>"
            if Contenedor["cc_bloquear"] == 1:
                Opciones += "<span class='ms-1 me-1'><i class='mdi mdi-lock'></i></span>"

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
            Ruta += "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(Contenedor["cc_tipo_actual"])+");Color:var(--Color_"+str(Contenedor["cc_tipo_actual"])+");'>__</span>"
            try:
                Tipo = Info_Actual[Contenedor["cc_tipo_actual"]]
                Ruta += str(Tipo)
                Tipo_1 = Info_Actual[Tipo]
                Ruta += "/"+str(Tipo_1)
                Tipo_2 = Info_Actual[Tipo_1]
                Ruta += "/"+str(Tipo_2)
            except:
                pass

            if "Fecha de Salida" in Info_Actual.keys():
                Ruta += " <"+str(Info_Actual["Fecha de Salida"])+">"

            Estado = ""
            if "Etapa" not in Info_Actual.keys():
                if Contenedor["cc_tipo_actual"] is None:
                    Estado = "<span class='fw-bold text-danger'>Not assigned</span>"
                else:
                    Estado = "Waiting"
            else:
                Estado = Info_Actual["Etapa"]
                if Info_Actual["Etapa"] == "Pending Exit":
                    if "Return Empty" in str(Info_Actual):
                        pass
                        #Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Salida("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-exit-to-app'></i></button>"
                    pass

            Opciones += "</div>"
            for K in Contenedor.keys():
                if Contenedor[K] is None:
                    Contenedor[K] = ""
            Patio.append({"Tiempo":"<span class='tiempo' tiempo='"+str(Contenedor["cc_ultimo_mov"])+"'></span>","Caja":str(Contenedor["cc_contenedor"]),"Carrier":Info_Actual["Carrier"],"Tipo":Contenedor["cc_tipo_actual"],"Estado":str(Estado),"Sello":Sello,"Ruta":Ruta,"Opciones":"<div class='text-center'>"+str(Opciones)+"</div>"})

        Dock = []
        for Contenedor in Contenedores_Dock:
            Opciones = "<div class='btn-group' role='group' aria-label='Basic mixed styles example'>"
            Estado = ""
            Info_Actual = json.loads(str(Contenedor["cc_informacion_actual"]))

            Opciones += "<button class='btn btn-sm btn-primary p-0 ps-1 pe-1' onclick='Ver_Historico("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-history'></i></button>"
            if Contenedor["cc_bloquear"] == 0:
                #Opciones += "<button class='btn btn-sm btn-warning p-0 ps-1 pe-1' onclick='Modificar("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-pencil'></i></button>"
                Opciones += "<button class='btn btn-sm btn-danger p-0 ps-1 pe-1' onclick='Eliminar_Caja("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-trash-can'></i></button>"
            if Contenedor["cc_qr_salida"] is None:
                Opciones += "<button class='btn btn-sm btn-info p-0 ps-1 pe-1' onclick='Regresar("+str(Contenedor["cc_id"])+",\""+str(Contenedor["cc_contenedor"])+"\")'><i class='mdi mdi-cursor-move'></i></button>"
            if Contenedor["cc_bloquear"] == 1:
                Opciones += "<span class='ms-1 me-1'><i class='mdi mdi-lock'></i></span>"
                
        
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
            Ruta += "<span class='h-100 border' style='width:20px; background:var(--Color_"+str(Contenedor["cc_tipo_actual"])+");Color:var(--Color_"+str(Contenedor["cc_tipo_actual"])+");'>__</span>"
            try:
                Tipo = Info_Actual[Contenedor["cc_tipo_actual"]]
                Ruta += str(Tipo)
                Tipo_1 = Info_Actual[Tipo]
                Ruta += "/"+str(Tipo_1)
                Tipo_2 = Info_Actual[Tipo_1]
                Ruta += "/"+str(Tipo_2)
            except:
                pass

            Estado = ""
            if "Etapa" not in Info_Actual.keys():
                Estado = "Waiting"
            else:
                Estado = Info_Actual["Etapa"]

            for K in Contenedor.keys():
                if Contenedor[K] is None:
                    Contenedor[K] = ""
                    
            Dock.append({"Dock":str(Contenedor["cc_dock"]),"Tiempo":"<span class='tiempo' tiempo='"+str(Contenedor["cc_ultimo_mov"])+"'></span>","Caja":str(Contenedor["cc_contenedor"]),"Carrier":Info_Actual["Carrier"],"Tipo":Contenedor["cc_tipo_actual"],"Estado":Estado,"Sello":Sello,"Ruta":Ruta,"Opciones":"<div class='text-center'>"+str(Opciones)+"</div>"})

        Resultado["Contenido"] += """<div class='text-end pe-1 pt-1'><small class='link-primary' style='cursor:pointer' onclick='Llamar_Funcion(\""""+str(Dir_Raiz)+"""\");'>Actualizar <i class='mdi mdi-refresh'></i></small></div>"""
        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi-truck'></i>Administrador de Cajas</div>"
        Resultado["Contenido"] += """
        <div class='row'>
            <div class='col-6'>
                <div class='h3 text-center'><i class='mdi mdi-car-brake-parking'></i> Yard</div>
                <div id='Tabla_Patio' class='border border-dark bg-dark-subtle'></div>
                <script>
                    delete Tabla_Patio;
                    var Tabla_Patio = new Tabulator("#Tabla_Patio", {
                        minHeight:800,
                        layout:"fitColumns",
                        data:"""+str(Patio)+""",
                        columns:[
                            {field:"Tiempo","title":"Ultimo Movement",formatter:"html"},
                            {field:"Caja","title":"Caja",headerFilter:"input"},
                            {field:"Carrier","title":"Carrier"},
                            {field:"Tipo","title":"Tipo"},
                            {field:"Ruta","title":"Ruta",formatter:"html"},
                            {field:"Sello","title":"Sello",formatter:"html"},
                            {field:"Estado","title":"Estado",formatter:"html"},
                            {field:"Opciones","title":"Op",formatter:"html"}
                        ]
                    });
                    Tabla_Patio.on("tableBuilt", function(){ 
                        setTimeout(() => {
                            Tabla_Patio.clearFilter();
                            Tabla_Patio.redraw();
                        }, 100);
                    });
                </script>
            </div>
            <div class='col-6'>
                <div class='h3 text-center'><i class='mdi mdi-sign-caution'></i> Docks</div>
                <div id='Tabla_Docks' class='border border-dark bg-dark-subtle'></div>
                <script>
                    delete Tabla_Docks;
                    var Tabla_Docks = new Tabulator("#Tabla_Docks", {
                        minHeight:800,
                        layout:"fitColumns",
                        data:"""+str(Dock)+""",
                        columns:[
                            {field:"Dock","title":"Dock"},
                            {field:"Tiempo","title":"Ultimo Movement",formatter:"html"},
                            {field:"Caja","title":"Caja",headerFilter:"input"},
                            {field:"Carrier","title":"Carrier"},
                            {field:"Tipo","title":"Tipo"},
                            {field:"Ruta","title":"Ruta",formatter:"html"},
                            {field:"Opciones","title":"Op",formatter:"html"}
                        ]
                    });
                    Tabla_Docks.on("tableBuilt", function(){ 
                        setTimeout(() => {
                            Tabla_Docks.clearFilter();
                            Tabla_Docks.redraw();
                        }, 100);
                    });
                </script>
            </div>
        </div>
        <script>
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
            function Modificar(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-pencil'></i> Modify ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm')
                var parametros = {"Fun":"Modificar","ID":ID};
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
            function Opciones(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-menu'></i> Options ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm')
                var parametros = {"Fun":"Opciones","ID":ID};
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
            function Regresar(ID,Contenedor){
                Mostrar_Ventana_Cargando(false);
                $("#Vent_1").find(".modal-title").html("<i class='mdi mdi-history'></i> Return previous ["+Contenedor+"]");
                $("#Vent_1").removeClass('modal-xl modal-lg modal-sm').addClass('modal-xl');
                var parametros = {"Fun":"Regresar","ID":ID};
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
            function Eliminar_Caja(ID,Contenedor){
                Swal.fire({
                title: '¿Estás seguro de eliminar el contenedor ['+Contenedor+']?',
                input: 'text',buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Si",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                preConfirm: (Comentario) => {
                    
                    if(Comentario.trim() == ""){
                        Mensaje(1,"Agrega comentario");
                    }else{
                        Mostrar_Ventana_Cargando(false);
                        var parametros = {"Fun":"Eliminar_Caja","Comentario":Comentario,"ID":ID};
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
                })
            }
        </script>
        """
    except:
         Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Regresar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    try:
        Historico = []
        Usuarios = DB.Get_Dato("SELECT USR.cusrid,EMP.Nombre FROM portal.cuser USR inner join linc.empleados EMP on USR.cusidempleado = EMP.IDEmpleado ")

        Dock = []
        DOCK_DISPONIBLES = DB.Get_Dato("""
        SELECT * FROM """+str(BD_Nombre_Yard)+""".cdock_asignacion DOCK 
        left join """+str(BD_Nombre_Yard)+""".ccajas CAJAS on CAJAS.cc_activo = 1 and CAJAS.cc_dock = DOCK.cd_dock and CAJAS.cc_id != '"""+str(Datos["ID"])+"""'
        where DOCK.cd_ub = '"""+str(Bandera_Dock)+"""' and cd_activo = 1 and CAJAS.cc_id is null
        """)
        for D in DOCK_DISPONIBLES:
            Dock.append(str(D["cd_dock"]))

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
            Dock_Aqui = H["cch_dock"]
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
                    Archivos += "<a href='http://10.4.7.219:8080/Portal_File/"+str(P)+"' target='_blank'><i class='mdi mdi-file-image'></i></a>"
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
            if Dock_Aqui is None:
                Historico.append({"Fecha":H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Usuario":str(Usuario),"Ubicacion":str(Ubicacion),"Carrier":Info_Actual["Carrier"],"Tipo":Tipo_Gen,"Estado":Estado,"Movimiento":str(H["cch_movimiento"]),"Archivos":Archivos,"Sello":Sello,'Move':"<div class='text-center'><button class='btn btn-primary btn-sm p-0 ps-1 pe-1' onclick='Regresa_Guardar("+str(Datos["ID"])+",\""+str(H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"))+"\")'><i class='mdi mdi-forward'></i></button</div>"})
            else:
                if str(Dock_Aqui) in Dock:
                    Historico.append({"Fecha":H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Usuario":str(Usuario),"Ubicacion":str(Ubicacion),"Carrier":Info_Actual["Carrier"],"Tipo":Tipo_Gen,"Estado":Estado,"Movimiento":str(H["cch_movimiento"]),"Archivos":Archivos,"Sello":Sello,'Move':"<div class='text-center'><button class='btn btn-primary btn-sm p-0 ps-1 pe-1' onclick='Regresa_Guardar("+str(Datos["ID"])+",\""+str(H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"))+"\")'><i class='mdi mdi-forward'></i></button</div>"})
                else:
                    Historico.append({"Fecha":H["cch_fecha_hora"].strftime("%Y-%m-%d %H:%M:%S"),"Usuario":str(Usuario),"Ubicacion":str(Ubicacion),"Carrier":Info_Actual["Carrier"],"Tipo":Tipo_Gen,"Estado":Estado,"Movimiento":str(H["cch_movimiento"]),"Archivos":Archivos,"Sello":Sello,'Move':""})
        Resultado["Contenido"] += """
        <div id='Tabla_Historico' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Historico;
            var Tabla_Historico = new Tabulator("#Tabla_Historico", {
                minHeight:500,
                layout:"fitColumns",
                data:"""+str(Historico)+""",
                columns:[
                    {field:"Move","title":"Move",formatter:"html"},
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

            function Regresa_Guardar(ID,Fecha){
                
                Swal.fire({
                title: '¿Estás seguro de regresar el contenedor a la fecha ['+Fecha+']?',
                input: 'text',buttonsStyling: false,showCancelButton: true,confirmButtonText: "<i class='mdi mdi-check'></i> Yes",cancelButtonText: "<i class='mdi mdi-close'></i> No",showLoaderOnConfirm: true,
                customClass: {confirmButton: 'btn btn-success ms-1 me-1',cancelButton: 'btn btn-danger ms-1 me-1'},
                preConfirm: (Comentario) => {
                   Mostrar_Ventana_Cargando(false);
                    var parametros = {"Fun":"Regresa_Guardar","Comentario":Comentario,"ID":ID,"Fecha":Fecha};
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
def Regresa_Guardar(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    Error = ""
    try:
        Info_Historico = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas_moviemiento WHERE cch_master = '"+str(Datos["ID"])+"' AND cch_fecha_hora = '"+str(Datos["Fecha"])+"'")[0]
        Dock = "null"
        if Info_Historico["cch_dock"] is not None:
            Dock = "'"+str(Info_Historico["cch_dock"])+"'"
        Tipo = "null"
        if Info_Historico["cch_tipo_actual"] is not None:
            Tipo = "'"+str(Info_Historico["cch_tipo_actual"])+"'"
        
        Negocio = "null"
        if Info_Historico["cch_negocio"] is not None:
            Negocio = "'"+str(Info_Historico["cch_negocio"])+"'"

        Zona = "null"
        if Info_Historico["cch_zona"] is not None:
            Zona = "'"+str(Info_Historico["cch_zona"])+"'"

        if "Comentario" not in Datos.keys():
            Datos["Comentario"] = ""
        
        Info_Aqui = json.loads(str(Info_Historico["cch_informacion_actual"]))
        
        Error += DB.Instruccion("UPDATE "+str(BD_Nombre_Yard)+".ccajas SET cc_ubicacion = '"+str(Info_Historico["cch_ubicacion"])+"',cc_informacion_actual = '"+str(Info_Historico["cch_informacion_actual"])+"',cc_tipo_actual = "+str(Tipo)+", cc_dock= "+str(Dock)+", cc_negocio = "+str(Negocio)+", cc_zona ="+str(Zona)+",  cc_ultimo_mov = NOW() WHERE cc_id = '"+str(Datos["ID"])+"' ")
        if Error == "":
            Info_Ahora = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
            Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
            Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
            Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
            Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
            Error = DB.Instruccion(""" 
            INSERT INTO """+str(BD_Nombre_Yard)+""".ccajas_moviemiento
            (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_comentario)
            VALUES
            ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(Info_Ahora["cc_informacion_actual"])+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','REGRESO A ["""+str(Datos["Fecha"])+"""]','"""+str(Datos["Comentario"])+"""')
            """)
        if Error == "":
            Resultado["Estado"] = 1
        else:
            Resultado["Contenido"] += str(Error)
    except:
        Resultado["Contenido"] = str(sys.exc_info())
    Cur += json.dumps(Resultado)
    print(Cur)
def Eliminar_Caja(Datos):
    DB = LibDM_2023.DataBase()
    Compartido_2023 = LibDM_2023.Compartido()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
    Resultado = {"Contenido":"","Estado":0}
    Error = ""
    try:
        if "Comentario" not in Datos.keys():
            Datos["Comentario"] = ""
        Error += DB.Instruccion("UPDATE "+str(BD_Nombre_Yard)+".ccajas SET cc_activo = 0, cc_ultimo_mov = NOW() WHERE cc_id = '"+str(Datos["ID"])+"' ")
        if Error == "":
            Info_Ahora = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas WHERE cc_id = '"+str(Datos["ID"])+"'")[0]
            Info_Ahora["cc_dock"] = 'null' if Info_Ahora["cc_dock"] is None else "'"+str(Info_Ahora["cc_dock"])+"'"
            Info_Ahora["cc_tipo_actual"] = 'null' if Info_Ahora["cc_tipo_actual"] is None else "'"+str(Info_Ahora["cc_tipo_actual"])+"'"
            Info_Ahora["cc_zona"] = 'null' if Info_Ahora["cc_zona"] is None else "'"+str(Info_Ahora["cc_zona"])+"'"
            Info_Ahora["cc_negocio"] = 'null' if Info_Ahora["cc_negocio"] is None else "'"+str(Info_Ahora["cc_negocio"])+"'"
            Error = DB.Instruccion(""" 
            INSERT INTO """+str(BD_Nombre_Yard)+""".ccajas_moviemiento
            (cch_master,cch_fecha_hora,cch_contenedor,cch_ubicacion,cch_informacion_actual,cch_dock,cch_tipo_actual,cch_zona,cch_negocio,cch_usuario,cch_movimiento,cch_comentario)
            VALUES
            ('"""+str(Info_Ahora["cc_id"])+"""',NOW(),'"""+str(Info_Ahora["cc_contenedor"])+"""','"""+str(Info_Ahora["cc_ubicacion"])+"""','"""+str(Info_Ahora["cc_informacion_actual"])+"""',"""+str(Info_Ahora["cc_dock"])+""","""+str(Info_Ahora["cc_tipo_actual"])+""","""+str(Info_Ahora["cc_zona"])+""","""+str(Info_Ahora["cc_negocio"])+""",'"""+str(Datos["ID_User"])+"""','ELIMINADO','"""+str(Datos["Comentario"])+"""')
            """)
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

