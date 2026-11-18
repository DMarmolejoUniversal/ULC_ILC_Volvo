import sys
import cgi
import os
from http import cookies
import json
import locale
import hashlib
sys.path.insert(1,r'C:\inetpub\wwwroot\Portal2\Componentes_2023')
import LibDM_2023
locale.setlocale( locale.LC_ALL, '' )
if "HTTP_COOKIE" in os.environ:
    cookievalue = cookies.SimpleCookie(os.environ["HTTP_COOKIE"])
else:
    cookievalue = None
Dir_Raiz = "../"
Menu_Activo = "ULC Saltillo / VOLVO"
BD_Nombre = "ulc_saltillo_volvo"
Bandera_Dock = "VOLVO"


    
def Inicio(Datos):
    try:
        Menu = LibDM_2023.Menu()
        Compartido = LibDM_2023.Compartido()
        Cur = "content-type: text/html;charset=ISO-8859-1\n\n"
        Cur += """
        <!DOCTYPE html>
        <html class='h-100 w-100'>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1">
        """
        Cur += str(Compartido.Complementos(Dir_Raiz,["07","15","04"]))
        Cur += """
            </head>
            <body style='height:90%;'>
        """
        Cur += str(Menu.Menu(Menu_Activo,Dir_Raiz,Datos["Key_U"]))
        Cur += """
                <div class='w-100 h-100 position-relative'>
                    <div style='color: #ebebeb; text-shadow: -1px 0 #8f8f8f, 0 1px #bbbbbb, 1px 0 #8f8f8f, 0 -1px #bbbbbb;' class="fw-bold position-absolute top-50 start-50 translate-middle display-1 z-n1" id='Titulo'>"""+str(Menu.Get_Titulo(Menu_Activo))+"""</div>
                    <div id='Pag' class='container-fluid pb-5'></div>
                </div>
                <script>
                    $( document ).ready(function() {
                        setInterval(Actualizar_Hora_Final, 500)
                    });
                    function Actualizar_Hora_Final()
                    {
                        $(".tiempo").each(function() {
                            var Limite = moment($(this).attr('tiempo'));
                            var Ahora = moment(new Date()); //todays date
                            var Faltan = moment.duration(Ahora.diff(Limite));
                            var Dias = Faltan.asDays().toString().split(".")[0];
                            var Horas = Faltan.hours();
                            if(Horas.toString().length == 1)
                                Horas = "0"+Horas;
                            var Minutos = Faltan.minutes();
                            if(Minutos.toString().length == 1)
                                Minutos = "0"+Minutos;
                            var Segundos = Faltan.seconds();
                            if(Segundos.toString().length == 1)
                                Segundos = "0"+Segundos;
                            if (Dias == 0)
                                $(this).html("<b>" + Horas +":" + Minutos +":"+ Segundos +"</b>")
                            else
                                $(this).html("<b>" + Dias + " Days, " + Horas +":" + Minutos +":"+ Segundos +"</b>")

                        });

                    }
                </script>
            </body>
        </html>
        """
    except:
        Cur = "content-type: text/html;charset=ISO-8859-1\n\n "
        Cur += str(sys.exc_info())
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