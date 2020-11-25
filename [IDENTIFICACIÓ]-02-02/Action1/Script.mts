'**************************************************************************
'Nom de l'script: [IDENTIFICACIÓ]-02-02
'Cas de prova: Identificació a Portal mitjançant certificat digital EC-Ciutadania per a si mateix
'Fitxer de dades necessari: N/A
'Nom de la biblioteca: -
'Nom del repositori d'objectes: Repositorio de objetos - [ALM] Subject\PT\IDENTIFICACIÓ\[IDENTIFICACIÓ]-01-01
'Autor OQ ATC
'Creat el: 30/09/2020
'Versió 1,0
'Data de revisió: 30/09/2020
'Revisat per: OQ ATC
'Canvis: -
'**************************************************************************

'Accedim a mapa web
' Cal afegir selecció d'enllaç en funció de l'entorn a testejar
'A integració: https://integracio.seu.atc.intranet.gencat.cat/mapaWeb.aspx
'A preproducció: https://preproduccio.seu.atc.intranet.gencat.cat/mapaWeb.aspx
OpenPortalMapaWeb
'Inici prova
Dim DataAntiga
DataAntiga = DateAdd("y",-10,Date)
'Msgbox (DataAntiga)

wait 5
With Browser("page. ATC - Seu electronica")
'Obrim qualsevol impost, per accedir mitjançant certificat digital
.Page("page. ATC - Seu electronica").Link("Taxa fiscal sobre el joc").Click
wait 5
'accedim mitjançant certificat digital (EC-Ciutadania)
.Page("Inici de sessió").WebButton("Certificat digital:  idCAT,").Click
.Dialog("Seguridad de Windows").WinObject("Seguridad de Windows").Static("Emisor: EC-Ciutadania").Click 118,6
.Dialog("Seguridad de Windows").WinButton("Aceptar").Click
'Esperem que carregui el formulari
wait 20
'Comprovem que es mostra El nom del Titular i que es correspon amb el del certificat

.Page("Model 040. Formulari telemàtic").Check CheckPoint("Model 040. Formulari telemàtic. ATC - Seu electronica") @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Model 040. Formulari telemàtic")_;_script infofile_;_ZIP::ssf3.xml_;_
.Page("Model 040. Formulari telemàtic").WebElement("Copiar les dades del presentad").Click
.Page("Model 040. Formulari telemàtic").WebElement("Emplenar autoliquidació").Click
'.Page("Model 040. Formulari telemàtic").SAPUIButton("Emplenar autoliquidació").Click
'Seleccionem la data. 
'.Page("Model 040. Formulari telemàtic").WebEdit("ctl00$MainContent$MainControl$_4").Set DataAntiga
.Page("Model 040. Formulari telemàtic").WebEdit("ctl00$MainContent$MainControl$_2").Set DataAntiga

wait 5
.Page("Model 040. Formulari telemàtic").WebList("sTarifa").Select "JCA"
.Page("Model 040. Formulari telemàtic").WebEdit("ctl00$MainContent$MainControl$").Set "100,00"
.Page("Model 040. Formulari telemàtic").WebElement("Validar").Click
.Page("Model 040. Formulari telemàtic").WebElement("D'acord").Click
wait 49
.Page("Carpeta de tramitació.").Check CheckPoint("Carpeta de tramitació. ATC - Seu electronica") @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Carpeta de tramitació.")_;_script infofile_;_ZIP::ssf4.xml_;_
'Tanquem sessió i navegador
.Page("Carpeta de tramitació.").Link("LogOut").Click
wait 5
Browser("page. ATC - Seu electronica").CloseAllTabs	
End With

