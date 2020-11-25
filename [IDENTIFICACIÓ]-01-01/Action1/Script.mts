'**************************************************************************
'Nom de l'script: [IDENTIFICACIÓ]-01-01
'Cas de prova: Identificació a Portal mitjançant certificat digital EC-SectorPublic 
'Fitxer de dades necessari: N/A
'Nom de la biblioteca: -
'Nom del repositori d'objectes: Repositorio de objetos - [ALM] Subject\PT\IDENTIFICACIÓ\[IDENTIFICACIÓ]-01-01
'Autor OQ ATC
'Creat el: 29/09/2020
'Versió 1,0
'Data de revisió: 29/09/2020
'Revisat per: OQ ATC
'Canvis -
'**************************************************************************
'test
'Accedim a mapa web

OpenPortalMapaWeb
'Inici prova
With Browser("page. ATC - Seu electronica")
'Obrim qualsevol impost, per accedir mitjançant certificat digital
.Page("page. ATC - Seu electronica").Link("Taxa fiscal sobre el joc").Click @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("page. ATC - Seu electronica").Link("Taxa fiscal sobre el joc")_;_script infofile_;_ZIP::ssf1.xml_;_
'accedim mitjançant certificat digital (EC-SectorPublic)
.Page("Inici de sessió").WebButton("Certificat digital:  idCAT,").Click @@ hightlight id_;_132682_;_script infofile_;_ZIP::ssf1.xml_;_
.Dialog("Seguridad de Windows").WinObject("Seguridad de Windows").Static("00000000T Persona de la").Click 101,19 @@ hightlight id_;_2067342232_;_script infofile_;_ZIP::ssf15.xml_;_
.Dialog("Seguridad de Windows").WinButton("Aceptar").Click @@ hightlight id_;_658764_;_script infofile_;_ZIP::ssf12.xml_;_
.Page("Inici de sessió").WebButton("Continua").Click @@ hightlight id_;_527552_;_script infofile_;_ZIP::ssf6.xml_;_
'Esperem que carregui el formulari
wait 20
.Page("Model 040. Formulari telemàtic").Sync @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Model 040. Formulari telemàtic")_;_script infofile_;_ZIP::ssf7.xml_;_
'Comprovem que es mostra El nom del presentador i titular i que es correspon amb el del certificat
.Page("Model 040. Formulari telemàtic").Check CheckPoint("Model 040. Formulari telemàtic. ATC - Seu electronica_4") @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Model 040. Formulari telemàtic")_;_script infofile_;_ZIP::ssf14.xml_;_
'Tanquem sessió i navegador
.Page("Model 040. Formulari telemàtic").Link("LogOut").Click @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Model 040. Formulari telemàtic").Link("LogOut")_;_script infofile_;_ZIP::ssf10.xml_;_
.Page("Formulari de login GICAR").Sync @@ hightlight id_;_Browser("page. ATC - Seu electronica").Page("Formulari de login GICAR")_;_script infofile_;_ZIP::ssf11.xml_;_
Browser("page. ATC - Seu electronica").CloseAllTabs	
End With

