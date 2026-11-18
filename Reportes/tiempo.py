import sys
import cgi
import os
from http import cookies
import json
import locale
import hashlib
from datetime import datetime
import random
import string
sys.path.insert(1,r'C:\inetpub\wwwroot\Portal2\Componentes_2023')
import LibDM_2023
locale.setlocale( locale.LC_ALL, '' )
if "HTTP_COOKIE" in os.environ:
    cookievalue = cookies.SimpleCookie(os.environ["HTTP_COOKIE"])
else:
    cookievalue = None
Dir_Raiz = "./Reportes/tiempo.py"

BD_Nombre_Yard = "ulc_saltillo"
DB_Nombre_Negocio = "ulc_saltillo_volvo"
Negocio = "VOLVO"
Bandera_Dock = "VOLVO"

    
def Inicio(Datos):
    DB = LibDM_2023.DataBase()
    Cur = "content-type: text/html;charset=ISO-8859-1\n\n"
    Resultado = {"Contenido":"","Estado":0}
    try:
        Contenedores_Patio = DB.Get_Dato("SELECT * FROM "+str(BD_Nombre_Yard)+".ccajas where cc_negocio = '"+str(Negocio)+"' AND cc_activo = 1")
        Patio = []
        for Contenedor in Contenedores_Patio:
            Info_Actual = json.loads(str(Contenedor["cc_informacion_actual"]))
            Opciones = ""
            UB = ""
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

            
            Patio.append({"Tiempo":"<span class='tiempo' tiempo='"+str(Contenedor["cc_ultimo_mov"])+"'></span>","Caja":str(Contenedor["cc_contenedor"]),"Carrier":Info_Actual["Carrier"],"Tipo":Contenedor["cc_tipo_actual"],"Estado":Ruta,"Opciones":Opciones,"UB":UB})
        Resultado["Contenido"] += "<div class='h2 fw-lighter mt-1 mb-1 text-center border-bottom'><i class='mdi mdi-car-brake-parking'></i> Reporte de Tiempo</div>"
        Resultado["Contenido"] += """
        <div class='container'>
        <div id='Tabla_Patio' class='border border-dark bg-dark-subtle'></div>
        <script>
            delete Tabla_Patio;
            var Tabla_Patio = new Tabulator("#Tabla_Patio", {
                minHeight:800,
                layout:"fitColumns",
                data:"""+str(Patio)+""",
                rowFormatter:function(row){
                    if(row.getData().Estado == "Not assigned"){
                        $(row.getElement()).find('div').eq(5).addClass('text-danger fw-bold');
                        if(row.getPosition()%2)
                            row.getElement().style.backgroundColor = "#F8D7DA";
                        else
                            row.getElement().style.backgroundColor = "#f7afb5";
                    }
                },
                columns:[
                    {field:"Tiempo","title":"Ultimo Movimiento",formatter:"html"},
                    {field:"Caja","title":"Contenedor",headerFilter:"input"},
                    {field:"Carrier","title":"Carrier"},
                    {field:"Tipo","title":"Tipo"},
                    {field:"Estado","title":"Ruta",formatter:"html"}
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
        """
    except:
        Resultado["Contenido"] += str(sys.exc_info())
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