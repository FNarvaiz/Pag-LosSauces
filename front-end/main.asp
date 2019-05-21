<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% option explicit %>

<!--#include file="const.asp"-->
<!--#include file="usr.asp"-->
<!--#include file="login.asp"-->
<!--#include file="classifieds.asp"-->
<!--#include file="surveys.asp"-->
<!--#include file="suppliers.asp"-->
<!--#include file="timeLine.asp"-->
<!--#include file="downloads.asp"-->
<!--#include file="bookings.asp"-->
<!--#include file="utils/db.asp"-->
<!--#include file="forms/formsData.asp"-->
<!--#include file="forms/formsUtils.asp"-->
<%

dim serverApp: serverApp = "front-end/main.asp"

function renderMainMenu
  if getUsrData then
    %>
    <div id="espacio"></div>
    <div onclick="menu()" class="menu_bar">
      <span>Vecinos de los Sauces</span>
    </div>
    <ul id="navegador">
      <%
      dim labels
      labels = eval("mainMenuLabels" & lang)
      dim i
      for i = 0 to uBound(labels)
'        if mainMenuContents(i) <> "classifieds" or usrName = "gbd" then
          %>
          <li id="mainMenuOption<%= i %>" class="li" onclick="load(this, '<%= mainMenuContents(i) %>')"><%= labels(i) %></li>
          
          <%
      next
      %>
    </ul>
    <%
  else
    %>
    <div id="espacio"></div>
    <div onclick="menu()" class="menu_bar">
      <span>Vecinos de los Sauces</span>
    </div>
    <ul id="navegadorLogin">
      <li class="btn anchor" onclick="showLoginDialog()" id="loginMenuItem"><%= eval("loginMenuItem" & lang) %></li>
      <li class="btn anchor" onclick="load(this, 'registrationDialog')"><%= eval("registrationMenuItem" & lang) %></li>
      <li class="btn anchor" onclick="load(this, 'faq')"><%= eval("faqMenuItem" & lang) %></li>
    </ul>
    <%
  end if
end function

function renderUserMenu
  if not getUsrData then exit function
  dim logoutLabel, changePasswordLabel
  logoutLabel = eval("logoutLabel" & lang)
  changePasswordLabel = eval("changePasswordLabel" & lang)
  %>
  <span id="userName">Familia <%= usrName %></span>
  <span>|</span>
  <span class="anchor" onclick="showChangePasswordDialog()"><%= changePasswordLabel %></span>
  <span>|</span>
  <span class="anchor" onclick="logout()"><%= logoutLabel %></span>
  <%
end function

const newsCategoryId = 4

function renderHome
  if not getUsrData then exit function
  %>
  <div id="homeBg"></div>
  <div id="homePanel">
    <table cellpadding="0" cellspacing="3" width="100%" id="homeNews">
      <tr><th colspan="2" align="left" valign="middle">NOVEDADES</th></tr>
      <%
      if dbGetData("SELECT TOP 10 ID, NOMBRE, dbo.DATE_TO_STR(REC_FECHA, " & sQuotes(lang) & ") AS FECHA, CONTENIDO_TEXTO, ARCHIVO_FILENAME, ARCHIVO_FILESIZE " & _
          "FROM LINEA_TIEMPO WHERE APROBADO=1 AND ID_CATEGORIA=" & newsCategoryId & " AND CONTENIDO_TEXTO IS NOT NULL" & _
          " ORDER BY ANIO DESC, MES DESC, REC_FECHA DESC, ID DESC") then
        do while not rs.EOF
          %>
          <tr>
            <td width="80" align="left" valign="middle"><%= replace(rs("FECHA"), "/", "-") %></td>
            <td valign="middle"><span class="anchor"
              onclick="track('Novedades', '<%= rs("NOMBRE") %>', 'Familia ' + usrName); showDataViewer(<%= rs("ID") %>)"><%= rs("NOMBRE") %>
            </td>
          </tr>
          <%
          rs.MoveNext
        loop
      else
        %>
        <tr><td align="center" height="25">(No hay novedades en este momento.)</td></tr>
        <%
      end if
      dbReleaseData
      %>
    </table>
  </div>
  <div id="schedule" >
    <section class="days">
      <h2>LUNES</h2>
      <article class="activityFirst">
        <h3 class="hour">8:30 - 10:30 </h3>
        <p >
          Yoga jovita- house
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">10:30 - 12:00</h3>
        <p >
           GYM LOCAL
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">16:45 - 18:00</h3>
        <p >
          Iniciación Danza 3 a 6
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">17:30 - 18:30</h3>
        <p >
           Hip Hop Infantil 7 a 8
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">18:15 - 19:30</h3>
        <p >
           Hip Hop Juvenil 9 a 12
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">18:30 - 20:00</h3>
        <p >
           Acrobacia Tela -SUM
        </p>
      </article>
    </section>
    <section class="days">
      <h2>MARTES</h2>
      <article class="activityFirst">
        <h3 class="hour">9:00 - 10:30 </h3>
        <p >
         Yoga Raquel
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">13:30 - 15:00 </h3>
        <p >
         GYM LOCAL
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">18:15 - 20:30 </h3>
        <p >
         GYM LOCAL
        </p>
      </article>
    </section>
     <section class="days">
      <h2>MIERCOLES</h2>
      <article class="activityFirst">
        <h3 class="hour">9:00 - 10:30 </h3>
        <p >
         Yoga Cristina
        </p>
      </article>
       <article class="activitySecond">
        <h3 class="hour">10:30 - 12:00 </h3>
        <p >
         GYM LOCAL
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">17:00 - 18:15 </h3>
        <p >
         Acrobacia Tela
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">18:15 - 19:30 </h3>
        <p >
         GYM Aero Box
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">19:15 - 20:30 </h3>
        <p >
         Yoga Cristina
        </p>
      </article>
    </section>
     <section class="days">
      <h2>JUEVES</h2>
      <article class="activityFirst">
        <h3 class="hour">13:30 - 15:00 </h3>
        <p >
         GYM Aero Box
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">18:15 - 19:30 </h3>
        <p >
         GYM Local
        </p>
      </article>
      <article class="activitySecond">
        <h3 class="hour">19:15 - 20:30</h3>
        <p >
         GYM Aero Box
        </p>
      </article>
    </section>
     <section class="days">
      <h2>VIERNES</h2>
      <article class="activityFirst">
        <h3 class="hour">8:30 - 10:30 </h3>
        <p >
         Yoga jovita - House
        </p>
      </article>
            <article class="activitySecond">
        <h3 class="hour">10:30 - 12:00 </h3>
        <p >
        GYM Local
        </p>
      </article>
    </section>
    <section class="days">
      <h2>SABADO</h2>
      <article class="activityFirst">
        <h3 class="hour">10:00 - 11:30 </h3>
        <p >
         Yoga Raquel
        </p>
      </article>
    </section>
     
  </div>
  <%
  
  'if dbGetData("SELECT HTML FROM CONTENIDOS WHERE ID=10") then
   '' response.write(rs(0))
  'end if
  'dbReleaseData
  logActivity "Inicio", ""
end function

function renderFAQ
  dim url: url = "contenidos/faq/faq.html"
  %>
  <div id="dynPanelBg"></div>
  <div id="contentsPanel">
    <div id="wrapper">
      <div id="qs">
        <h1><img src="contenidos/faq/signo.png" /> PREGUNTAS FRECUENTES</h1>
        <table width="800" cellpadding="10" cellspacing="5" >
          <tr>
            <td width="50%" class="preg" valign="top">
              <h2>Olvidé mi contraseña. ¿Cómo la recupero?</h2>
              <p>Es común no recordar este tipo de cosas.<br>Para recuperarla, en la pantalla de ingreso al sistema 
              hacé clic en “Olvidé mi contraseña”. Allí escribí el e-mail con el que te registraste, y el sistema te enviará la contraseña.</p>
            </td>
            <td class="preg" width="50%" valign="top">
              <h2>Si poseo más de un lote, ¿puedo visualizar los datos de todos ellos a la vez?</h2>
              <p>No. Cada usuario representa a un lote. Por ello, al poseer más de uno deberás ingresar de manera individual
              para visualizar la información de cada uno.</p>
            </td>
          </tr>
          
          <tr>
            <td colspan="1" class="preg" width="50%" valign="top">
              <h2>No puedo ingresar al sistema. ¿Qué hago?</h2>
              <p>Antes que nada asegúrate de ingresar correctamente tu e-mail y tu contraseña.
              Si continúas experimentando problemas, por favor, ponete en contacto con nosotros.</p>
            </td>
          </tr>
          <tr>
            <td colspan="1" class="preg" width="50%" valign="top">
              <h2>¿Cómo cambio mi contraseña?</h2>
              <p>Una vez que hayas ingresado al sistema, al pie de la pantalla hacé clic en “Cambiar contraseña”.
              Luego ingresá tu contraseña actual y la nueva que hayas elegido.
              Recordá que es aconsejable cambiar periódicamente la contraseña.</p>
            </td>
          </tr>
          
        </table>
      </div>
    </div>
  </div>
  <div id="contactPanel">
    Envianos tu consulta a <a href="mailto:info@vecinosdelossauces.com.ar">info@vecinosdelossauces.com.ar</a><br>
    o mediante el siguiente formulario.
    <form name="contactForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendContactMessage">
      <input type="hidden" name="sessionId" value="<%= request("sessionId") %>">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="mensaje de contacto">
      <table cellpadding="0" cellspacing="6" class="contact">
        <tr style="display: none">
          <td align="right">Comisión</td>
          <td>
            <select name="commissionId">
              <option selected="selected" value="">(seleccionar)</option>
              <%
              dbGetData("SELECT ID, NOMBRE FROM COMISIONES ORDER BY NOMBRE")
              do while not rs.EOF
                %>
                <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
                <%
                rs.moveNext
              loop
              dbReleaseData
              %>
            </select>
          </td>
        </tr>
        <tr>
          <td align="right">Asunto</td>
          <td><input name="subject" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td align="right" valign="top">Mensaje</td>
          <td><textarea name="message" rows="6" cols="31"></textarea></td>
        </tr>
        <tr>
          <td colspan="2" align="right"><input type="button" value="Enviar" style="width: 60px"
            onclick="sendFormData('contactForm')"></td>
        </tr>
      </table>
    </form>
    </center>
  </div>
  <%
  logActivity "FAQ", ""
end function

function renderAbout
  if not getUsrData then exit function
  dim url: url = "contenidos/about/about.html"
  %>
  <div id="dynPanelBg"></div>
  <iframe id="aboutContentPanel" scrolling="no" src="<%= url %>"></iframe>
  <div id="aboutResponsive">
    <section id="tabla">
      <article class="artAbout">
        <img src="contenidos/about/iconos/13.png"/>
        <title>Indendente</title>
        <p>Oscar Vega</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/3.png"/>
        <title>Espacios Verdes y Medio Ambiente</title>
        <p>Sandra Addiego, Andrea Regnicoli</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/1.png"/>
        <title>Deportes</title>
        <p>Alejandro Ramallo, Guillermo Balagna (Futbol), Juan Hormaechea (tenis)</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/2.png"/>
        <title>Arquitectura</title>
        <p>Martin Kordich, Patricio Jones, Gustavo Rozenman</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/2.png"/>
        <title>Seguridad</title>
        <p>Diego Rennison, Daniel Peuchot, 
Gabriel Fiore, Daniel De Andres, Gaspar Farina, Alex Borchardt</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/6.png"/>
        <title>Presupuesto</title>
    <p>Alejandro Ramallo, Claudia Moffa,  Mariano Serra</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/3.png"/>
        <title>Mesa de enlace Municipal</title>
        <p>Miguel Fliit</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/7.png"/>
        <title>Disciplina</title>
        <p>Tete Brizio, Sofia Dolega</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/4.png"/>
        <title>Club House</title>
        <p>Elizabeth Levi, Yanina Pausic, Jesica Suide, Mónica Alvarez</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/4.png"/>
        <title>Directores por AVN</title>
        <p>Manuel Kosoy y Carlos Fernandez Lobbe</p>
      </article>
      <article class="artAbout">
        <img src="contenidos/about/iconos/4.png"/>
        <title>Directores por los vecinos</title>
        <p>Miguel Flitt (Director) Claudia Moffa (Directora Suplente)</p>
      </article>
        
    </section>  
  </div>
  <div id="aboutContactPanel">
    <span style="font-size: 14px">Mensajes a la Administración</span><br><br>
    &nbsp;&nbsp;&nbsp;&nbsp;Por favor, completá el siguiente formulario<br>
    &nbsp;&nbsp;&nbsp;&nbsp;a la brevedad recibirás nuestra respuesta.<br><br>
    <form name="contactForm" action="<%= serverApp %>">
      <input type="hidden" name="content" value="sendContactMessage">
      <input type="hidden" name="sessionId" value="<%= request("sessionId") %>">
      <input type="hidden" name="lang" value="<%= lang %>">
      <input type="hidden" name="trackingLabel" value="mensaje de contacto">
      <table cellpadding="0" cellspacing="6" class="contact">
        <tr>
          <td align="right">Comisión</td>
          <td>
            <select name="commissionId">
              <option selected="selected" value="">(seleccionar)</option>
              <%
              dbGetData("SELECT ID, NOMBRE FROM COMISIONES ORDER BY NOMBRE")
              do while not rs.EOF
                %>
                <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
                <%
                rs.moveNext
              loop
              dbReleaseData
              %>
            </select>
          </td>
        </tr>
        <tr>
          <td align="right">Asunto</td>
          <td><input class="editbox" name="subject" type="text" size="40" maxlength="100"></td>
        </tr>
        <tr>
          <td align="right" valign="top">Mensaje</td>
          <td><textarea name="message" rows="9" cols="31"></textarea></td>
        </tr>
        <tr>
          <td colspan="2" align="right"><input class="button" type="button" value="Enviar" style="width: 60px"
            onclick="sendFormData('contactForm')"></td>
        </tr>
      </table>
    </form>
    </center>
  </div>
  <%
  logActivity "Quienes somos", ""
end function

function renderHelp
  if not getUsrData then exit function
  dim url: url = "contenidos/phones/phones.html"
  %>
  <div id="dynPanelBg"></div>
  <div id="helpBg"></div>
  <div id="helpPanel">
    <iframe id="contentsPanel" src="<%= url %>"></iframe>
  </div>
  <%
  logActivity "Ayuda", ""
end function

function servicesAllowed
  servicesAllowed = true
  if usrServicesPemission then exit function
  servicesAllowed = false
  %>
  <div id="servicesDisabled">Tu acceso a este servicio se encuentra deshabilitado por la Administración.</div>
  <%
end function

function sendContactMessage(commissionId, subject, message)
  if not getUsrData then exit function
  dbGetData("SELECT UNIDAD, EMAIL FROM VECINOS WHERE ID=" & usrId)
  dim usrUnit: usrUnit = rs("UNIDAD")
  dim usrEmail: usrEmail = rs("EMAIL")
  dbReleaseData
  message = "<div style='margin:3px 1px;border:1px solid #B39DDB; background-color:#EDE7F6;color:#000;padding: 5px; font-weight: normal;' ><h2 style='margin:6px 0px;color: #000;font-weight: normal;font-size: 17pt;' >Unidad/Lote: " & usrUnit & "<br>" & _
    "Familia: " & usrName & "</h2>" & _
    "<h4 style='margin:3px 0px;color:#000; font-weight: normal;font-size: 12pt;' >Mensaje: " & replace(message, vbLf, "<br>") &"</h4></div>"
  if dbGetData("SELECT MAILS, NOMBRE FROM COMISIONES WHERE ID=" & commissionId) then
    sendMail uCase(rs("NOMBRE")) & " - Vecinos de los Sauces", rs("MAILS"), subject, message, "Familia " & usrName, usrEmail' lCase(rs("SIGLA")) & "@vecinosdelossauces.com.ar", subject, message, "Familia " & usrName, usrEmail
    logActivity "Mensaje de contacto", rs("NOMBRE")
  else
    JSONAddOpFailed
    JSONAddMessage "La comisión seleccionada no existe; el mensaje no se envió."
    JSONSend
  end if
  dbReleaseData
end function

function MailPrueba
dim objMessage,objConfig,Flds
  set objMessage = createobject("cdo.message")
set objConfig = createobject("cdo.configuration")
Set Flds = objConfig.Fields
 
Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="localhost"
 
' ' Passing SMTP authentication
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="info@vecinosdelossauces.com.ar"
Flds.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") =PassMails
 
Flds.update
Set objMessage.Configuration = objConfig
objMessage.To = "facu.232@gmail.com"
objMessage.From = "info@vecinosdelossauces.com.ar"
objMessage.Subject = "ASUNTO"
objMessage.fields.update
objMessage.HTMLBody = "Mensaje"
objMessage.Send
end function

function sendMail(toName, toEmail, subject, message, replyToName, replyToEmail)
  on error resume next
  dbLog("sendMail: toName=" & toName & ", toEmail=" & toName & ", subject=" & subject)
  dbLog("sendMail: message=" & message)
  dbLog("sendMail: replyToName=" & replyToName & ", replyToEmail=" & replyToEmail)
  dim cdoMail: set cdoMail = server.CreateObject("CDO.Message")
  dim cdoConf: set cdoConf = server.createObject("CDO.Configuration") 
  dim cdoFields: set cdoFields = cdoConf.fields
  with cdoFields
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@vecinosdelossauces.com.ar"
    .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassMails
    .update
  end with
  with cdoMail
    set .configuration = cdoConf
    .from = replyToName & " <info@vecinosdelossauces.com.ar>"
    .subject = subject
    .HTMLBody = message
  end with
  if InStr(toEmail,",")>0 then
      cdoMail.to = toEmail
  else
      cdoMail.to = toName & " <" & toEmail & ">"
  end if
  if not isNull(replyToEmail) and not isNull(replyToName) then 
    cdoMail.Cc = replyToName & " <" & replyToEmail & ">"
  end if
  cdoMail.Send
  sendMail = err.number '0 = OK
   set cdoFields = nothing
  set cdoConf = nothing
  set cdoMail = nothing
  if err.number = 0 then
    JSONAddOpOK
    JSONAddMessage "El mensaje ha sido enviado."
  else
    JSONAddOpFailed
    JSONAddMessage "Ha ocurrido un error; el mensaje no se envió."
    JSONAdd "description", left(err.description, len(err.description) - 2)
  end if
  JSONSend
end function

dbConnect
select case getStringParam("content", 30)
	case "loginDialog": renderLoginDialog
	case "faq": renderFAQ
	case "login": login
	case "logout": logout
  case "sendPasswordRemainder": sendPasswordRemainder
	case "changePasswordDialog": renderChangePasswordDialog
  case "changePassword": doChangePassword
  case "passwordRecoveryDialog": renderPasswordRecoveryDialog
  case "sendPasswordRecoveryMessage": sendPasswordRecoveryMessage getStringParam("usr", 100)
  case "registrationDialog" renderRegistrationDialog
  case "sendRegistrationMessage": sendRegistrationMessage getIntegerParam("unit"), getStringParam("name", 100), getStringParam("email", 100)

  case "mainMenu": renderMainMenu
  case "userMenu": renderUserMenu
  case "home": renderHome
  case "about": renderAbout

  case "timeLine": renderTimeLine
  case "timeLineDetail": renderTimeLineDetail getIntegerParam("timeLineRow"), getStringParam("detailKeyValues", 50)

  case "timeLineData": sendTimeLineData getIntegerParam("recordId"), getIntegerParam("forceDownload") = 1
  case "timeLineThumbnail": sendTimeLineData getIntegerParam("recordId"), false ' sendTimeLineThumbnail getIntegerParam("recordId"), getIntegerParam("height")

  case "downloads": renderDownloads getIntegerParam("categoryId")
  case "downloadListing": renderDownloadListing getIntegerParam("categoryId")
  
  case "textContent": getTextContent getIntegerParam("contentId")

  case "services": renderServices
  case "bookings": renderBookings
  case "bookingsListing": renderBookingsListing
  case "jBookingStartOptions": jBookingStartOptions getIntegerParam("resourceId"), getDateTimeParam("bookingDate")
  case "bookingsStatus": renderBookingsStatus getIntegerParam("resourceId")
  case "bookingsTurnTypes": renderBookingsTurnTypes getIntegerParam("resourceId"), getDateTimeParam("bookingDate"), getIntegerParam("turnStart")
  case "bookingsTake": bookingsTake getIntegerParam("resourceId"), getDateTimeParam("bookingDate"), getIntegerParam("turnStart"), _ 
          getIntegerParam("turnType"), getStringParam("extraNeighborUnit", 20)
  case "bookingCancel": bookingCancel getIntegerParam("bookingId")
  case "bookingsSendRequest": bookingsSendRequest getDateTimeParam("bookingDate"),getIntegerParam("placeId"), getIntegerParam("turnId")
  
  case "surveys": renderSurveys
  case "saveSurveyVote": saveSurveyVote getIntegerParam("surveyId"), getIntegerParam("voteValue"), getIntegerParam("itemNumber")
  case "saveSurveyVoteComment": saveSurveyVoteComment getIntegerParam("surveyId"), getStringParam("voteComment", 200), getIntegerParam("itemNumber")
  case "surveyItem": renderSingleSurveyItem getIntegerParam(surveyItemId)
  case "suppliers": renderSuppliers getIntegerParam("categoryId")
  case "suppliersListing": renderSuppliersListing getIntegerParam("categoryId")
  case "saveSupplierVote": saveSupplierVote getIntegerParam("supplierId"), getIntegerParam("voteValue"), getIntegerParam("itemNumber")
  case "classifieds": renderClassifieds getIntegerParam("categoryId")
  case "classifiedsListing": renderClassifiedsListing getIntegerParam("categoryId")
  case "classifiedImage": sendClassifiedImage getIntegerParam("recordId")
  case "profile": response.write("")
  case "help": renderHelp

  case "sendContactMessage": sendContactMessage getIntegerParam("commissionId"), getStringParam("subject", 200), getStringParam("message", 2000)
  case else
    JSONAddOpFailed
    JSONAddMessage "Error interno. " & getStringParam("content", 30)
    JSONSend
end select
dbDisconnect

%>

