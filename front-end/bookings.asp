<%

const tennisCourt1BookingResourceId = 10
const tennisCourt2BookingResourceId = 20
const clubHouseBookingResourceId = 30

const footballClassBookingTypeId = 25
const tennisClassBookingTypeId = 30


function renderBookings
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  dim i
  %>
  <div id="dynPanelBg"></div>
  <div id="bookingsPanel">
    <div id="bookingsTitle">MIS RESERVAS</div>
    <div id="bookingsListing">
      <% renderBookingsListing %>
    </div>
  </div>
  <div id="bookingsFormPanel">
    <div id="bookingsFormTitle">CLUB HOUSE</div>
    <table id="bookingsformControls">
      <tr>
        <td valign="middle" width="125" align="right">Fecha del evento</td>
        <td valign="middle" width="100"><input type="text" maxlength="10" value="" id="bookingsFormDate" readonly="readonly"></td>
        <td valign="middle" width="50"><img src="back-end/forms/resource/btnDatePicker.png" class="anchor" title="Ver calendario"
          onclick="datePickerToggle(document.getElementById('bookingsFormDate')); event.cancelBubble=true;"></td>
        <td valign="middle" width="120" rowspan="2">
          <div class="bookingsButtons anchor" onclick="bookingsSendRequest()">Solicitar reserva</div>
        </td>
      </tr>
      <tr>
        <td valign="middle" width="125" align="right">Turno</td>
        <td valign="middle" width="100" colspan="2">
          <select id="bookingsFormTurn">
            <%
            dbGetData("SELECT ID, NOMBRE FROM RESERVAS_TURNOS_ESPECIALES WHERE ID_RECURSO=" & clubHouseBookingResourceId & " ORDER BY ID")
            i = 0
            do while not rs.EOF
              if i = 0 then
                %>
                <option value="<%= rs("ID") %>" selected="selected"><%= rs("NOMBRE") %></option>
                <%
              else
                %>
                <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
                <%
              end if
              i = i + 1
              rs.moveNext
            loop
            dbReleaseData
          %>
          </select>
        </td>
      </tr>
      <tr><td align="right" colspan="4" id="bookingsformNote"><br>La Administración se pondrá en contacto para confirmar la reserva.<br>
        Opción Horario / Valores:<br>
 Martes a viernes: 16 a 20 hs.  $1.500.- Sin costo adicional.<br>
 20 a 24 hs.  $2.000.- Mas $ 800.- del servicio de limpieza y control<br> 
 Viernes:  20 a 03 hs. $2.800.-  Mas $    1.000.- del servicio de limpieza y control  <br>
 Sábado:  13 a 17 hs. $2.500 sin costo adicional,<br>
 17 a 21 hs. $2.500.-  Mas $     450.-  del servicio de limpieza y control<br>
 21 a 01 hs. $2.800.-  Mas $     800.-  del servicio de limpieza y control<br>
 21 a 03 hs. $3.200.-  Mas $   1.000.- del servicio de limpieza y control</td></tr> 
    </table>
  </div>
  <div id="bookingsStatusPanel">
    <div id="bookingsStatusTitle">ESTADO DE LAS RESERVAS
      <div id="bookingsStatusControls">
        <select id="bookingsResourceSelector" onchange="bookingResourceSelected(this)">
          <%
          dbGetData("SELECT ID, NOMBRE FROM RESERVAS_RECURSOS ORDER BY ID")
          i = null
          do while not rs.EOF
            if isNull(i) then
              i = rs("ID")
              %>
              <option value="<%= i %>" selected="selected"><%= rs("NOMBRE") %></option>
              <%
            else
              %>
              <option value="<%= rs("ID") %>"><%= rs("NOMBRE") %></option>
              <%
            end if
            rs.moveNext
          loop
          dbReleaseData
        %>
        </select>
      </div>
    </div>
    <div id="bookingsStatus">
      <% renderBookingsStatus(i) %>
    </div>
  </div>
	<%
	logActivity "Reservas", ""
end function

function renderBookingsListing
  if not getUsrData then exit function
  %>
  <table cellpadding="4" cellspacing="0" width="100%">
    <tr>
      <th align="center">FECHA</th>
      <th align="center">RECURSO</th>
      <th align="center">TURNO</th>
      <th width="60">&nbsp;</th>
    </tr>
    <%
    if dbGetData("SELECT ID, FECHA, ID_RECURSO, dbo.NOMBRE_RECURSO_RESERVA(ID_RECURSO) AS RECURSO, " & _
        "dbo.NOMBRE_TIPO_RESERVA(ID_RECURSO, ID_TIPO) + ': ' + dbo.NOMBRE_TURNO(INICIO, DURACION) AS TURNO, " & _
        "CAST(CASE WHEN DATEDIFF(MINUTE, GETDATE(), dbo.FECHA_INICIO_TURNO(FECHA, INICIO)) >= 0 THEN 1 ELSE 0 END AS BIT) AS CANCELABLE " & _
        "FROM RESERVAS WHERE (ID_VECINO = " & usrId & " OR ID_VECINO_2 = " & usrId & ") AND FECHA >= (SELECT TOP 1 FECHA FROM HOY) AND FECHA_FIN > GETDATE() " & _ 
        "ORDER BY FECHA, RECURSO, INICIO DESC") then
      dim rowClass
      dim oddRow: oddRow = false
      dim cancelBtnVisibility
      do while not rs.EOF
        if rs("CANCELABLE") then
          cancelBtnVisibility = "visible"
        else
          cancelBtnVisibility = "hidden"
        end if
        if oddRow then rowClass = "oddRow" else rowClass = ""
        %>
        <tr class="<%= rowClass %>">
          <td width="75" align="center"><%= rs("FECHA") %></td>
          <td width="150" align="center"><%= rs("RECURSO") %></td>
          <td width="140" align="center"><%= rs("TURNO") %></td>
          <td>
            <div class="bookingsButtons anchor" style="visibility: <%= cancelBtnVisibility %>" 
              onclick="bookingsCancel(<%= rs("ID") %>,<%= rs("ID_RECURSO") %>)">Cancelar</div>
          </td>
        </tr>
        <%
        oddRow = not oddRow
        rs.moveNext
      loop
    else
      %>
      <tr><td colspan="4" align="left" height="25">(No hay ninguna reserva a utilizar.)</td></tr>
      <%
    end if
    dbReleaseData
    %>
  </table>
  <%
end function

function renderBookingsStatus(resourceId)
  if resourceId = clubHouseBookingResourceId then
    renderClubHouseBookingsStatus(resourceId)
  else
    renderTennisBookingsStatus(resourceId)
  end if
end function

function renderTennisBookingsStatus(resourceId)
  if not getUsrData then exit function
  
    dbGetData("SELECT E.FECHA, dbo.HOUR_TO_STR(dbo.MINUTOS_A_HORA(E.INICIO)) AS TURNO, E.INICIO, T.DURACION, E.ID_VECINO, E.ID_VECINO_2, E.ESTADO, " & _
      "CAST(CASE WHEN dbo.FECHA_FIN_TURNO(E.FECHA, E.INICIO, T.DURACION) < GETDATE() THEN 1 ELSE 0 END AS BIT) AS FINALIZADO, " & _
      "CAST(CASE WHEN GETDATE() < dbo.FECHA_INICIO_TURNO(FECHA, INICIO) AND UPPER(E.ESTADO)='DISPONIBLE' THEN 1 ELSE 0 END AS BIT) AS RESERVABLE, " & _
      "dbo.ID_TIPO_RESERVA(dbo.ID_RESERVA_TURNO_RESERVA(" & resourceId & ", E.FECHA, E.INICIO, E.DURACION)) AS TIPO_RESERVA, " & _
      "CAST(CAST(T.DURACION AS FLOAT) / 60 AS NVARCHAR) + ' hs (' + T.NOMBRE + ')' AS TIPO, T.NOMBRE AS ABR, T.DURACION, " & _
      "dbo.ESTADO_TURNO_RESERVA(" & resourceId & ", E.FECHA, E.INICIO, T.DURACION) AS ESTADO_TURNO, " & _
      "dbo.TURNO_RESERVA_VALIDO(" & resourceId & ", E.INICIO, T.DURACION) AS VALIDO, T.REQUIERE_VECINO_2, " & _
      "dbo.NOMBRE_TIPO_RESERVA(" & resourceId & ", E.ID_TIPO) AS TIPO_RESERVA_HECHA, T.ID AS ID_TIPO_RESERVA_DISPONIBLE, " & _
      "dbo.UNIDAD_VECINO(E.ID_VECINO) AS UNIDAD, dbo.UNIDAD_VECINO(E.ID_VECINO) AS UNIDAD_PRIMARIA, dbo.UNIDAD_VECINO(E.ID_VECINO_2) AS UNIDAD_EXTRA " & _
      "FROM dbo.ESTADOS_RESERVAS(" & resourceId & ") E " & _
      "JOIN RESERVAS_TIPOS T ON T.ID_RECURSO = " & resourceId & " AND T.ID>1 " & _
      "GROUP BY E.INICIO, E.FECHA, T.ID, T.DURACION, T.NOMBRE, E.ID_VECINO, E.ID_VECINO_2, E.ESTADO, E.DURACION, T.REQUIERE_VECINO_2, E.ID_TIPO " & _
      "ORDER BY  E.FECHA,E.INICIO, T.ID")
    dim rowClass
    dim dateClass: dateClass="date"
    dim oddRow: oddRow = false
    dim colTurn: colTurn = -1
    dim fecha: fecha = -1
    dim turnAnterior: turnAnterior = -2
    dim d: d = dateSerial(2000, 1, 1)
    dim emptyCell: emptyCell = true
    do while not rs.EOF
      if fecha<> rs("FECHA") then
        if fecha<> -1 then 
          %></section><%
          dateClass="date secondDate"
        end if
        %><section id="bookingRows" class="<%= dateClass%>"><h4><%= weekdayname(weekday(rs("FECHA"))) %>&nbsp;<%= zeroPad(day(rs("FECHA")), 2) %>/<%= zeroPad(month(rs("FECHA")), 2) %></h4><%
        fecha = rs("FECHA")
        d = rs("FECHA")
        colTurn = -1
      end if 
      if colTurn <> rs("INICIO") then 
        if oddRow then rowClass = "fila" else rowClass = "fila2"
        oddRow = not oddRow
        if colTurn <> -1 then 
          %> </div> <% 
        end if
        %><div class="<%= rowClass %>"><%
        colTurn = rs("INICIO")
        %><h5 class="bookingStatusTurn"><%= rs("TURNO") %></h5><%
      end if
      if rs("FINALIZADO") then
        if d <> rs("FECHA") then
          d = rs("FECHA")
        end if
      else
        if rs("RESERVABLE") then
          if resourceId < 26  or not(( Weekday(d,1) < 6 and rs("INICIO") = 1320 ) or (Weekday(d,1) >5 and rs("INICIO") = 480 )) then 
            if d <> rs("FECHA") then
              d = rs("FECHA")
              %><td width="190" class="bookingsButtons turnAvailable" align="center"><%
            end if
            if uCase(rs("ESTADO_TURNO")) = "DISPONIBLE" and rs("VALIDO") then
              %>
              <span class="bookingsTypeButton anchor" title="Reservar <%= rs("TIPO") %>"
                onclick="bookingsTake(this, <%= resourceId %>, '<%= day(d) %>/<%= month(d) %>/<%= year(d) %>', <%= rs("INICIO") %>, <%= rs("ID_TIPO_RESERVA_DISPONIBLE") %>, <%= abs(cInt(rs("REQUIERE_VECINO_2"))) %>)"><%= rs("ABR") %></span>
              <%
              emptyCell = false
            end if
          else
            d = rs("FECHA")
            %><div title="Bloqueado" class="bookingsButtons turnNotAvailable" align="center">No disponible</div><%
          end if 
        elseif uCase(rs("ESTADO")) = "VACANTE" then
          if d <> rs("FECHA") then
            d = rs("FECHA")
            %><div  title="<%= rs("ESTADO") %>" class="bookingsButtons turnNotUsed" ></div><%
          end if
        elseif rs("ID_VECINO") = usrId then
          if turnAnterior <> rs("INICIO") then
            if len(rs("UNIDAD_EXTRA")) > 0  then
              %><div class="bookingsButtons turnOwnReservation" ><%= rs("TIPO_RESERVA_HECHA") %> con lote <%= rs("UNIDAD_EXTRA") %></div><%
            else
              %><div  class="bookingsButtons turnOwnReservation" ><%= rs("TIPO_RESERVA_HECHA") %> (reserva propia)</div><%
            end if
            d = rs("FECHA")
          end if 
        elseif rs("ID_VECINO_2") = usrId then
          if turnAnterior <> rs("INICIO") then
            %><div class="bookingsButtons turnOwnReservation"><%= rs("TIPO_RESERVA_HECHA") %> con lote <%= rs("UNIDAD_PRIMARIA") %></div><%
          end if
        else
           if turnAnterior <> rs("INICIO") then
            if len(rs("UNIDAD_EXTRA")) > 0 then
              %><div title="<%= rs("ESTADO") %>" class="bookingsButtons turnNotAvailable"><%= rs("TIPO_RESERVA_HECHA") %>: <%= rs("UNIDAD") %> y <%= rs("UNIDAD_EXTRA") %></div><%
            else
              %><div title="<%= rs("ESTADO") %>" class="bookingsButtons turnNotAvailable"><%= rs("TIPO_RESERVA_HECHA") %>: <%= rs("UNIDAD") %></div><%
            end if
          end if
        end if
      end if
      turnAnterior= colTurn
      rs.moveNext
    loop
    %></div></section><%
    dbReleaseData
    
end function

function renderClubHouseBookingsStatus(resourceId)
  if not getUsrData then exit function
  %>
 <section class='headerTable'>
      <h4>FECHA</h4>
      <h4>TURNO</h4>
  </section>
    <%
    if dbGetData("SELECT FECHA, dbo.NOMBRE_TURNO(INICIO, DURACION) AS TURNO " & _
        "FROM dbo.ESTADOS_RESERVAS(" & resourceId & ") ORDER BY FECHA, INICIO") then
      dim rowClass
      dim oddRow: oddRow = false
      do while not rs.EOF
        if oddRow then rowClass = "fila2" else rowClass = "fila"
        oddRow = not oddRow
        %>
        <div class="<%= rowClass %>">
            <div class="rowText left"><%= rs("FECHA") %></div>
            <div class="rowText"><%= rs("TURNO") %></div>
          </div>
        <%
        rs.moveNext
      loop
    else
      %>
      <div class="fila"><div class="rowText">(No se han realizado reservas)</div></div>
      <%
    end if
    dbReleaseData
    %>
  <%
end function

function bookingCancel(bookingId)
  if not getUsrData then exit function
  dbGetData("SELECT ID_RECURSO, FECHA, dbo.NOMBRE_RECURSO_RESERVA(ID_RECURSO) AS RECURSO, dbo.HOUR_TO_STR(INICIO) AS TURNO, " & _
    "CAST(CASE WHEN DATEDIFF(MINUTE, GETDATE(), dbo.FECHA_INICIO_TURNO(FECHA, INICIO)) >= 120 THEN 1 ELSE 0 END AS BIT) AS OK " & _
    "FROM RESERVAS WHERE ID=" & bookingId)
  dim OK: OK = rs("OK")
  dim resourceId: resourceId = rs("ID_RECURSO")
  dim resource: resource = rs("RECURSO")
  dim turn: turn = rs("TURNO")
  dim bookingDate: bookingDate = rs("FECHA")
  dbReleaseData
  if resourceId = clubHouseBookingResourceId then
    JSONAddOpFailed
    JSONAddMessage "Para cancelar la reserva del Club House deberá comunicarse con la Administración."
    logActivity "Cancela reserva Club House", "Denegada: debe comunicarse con Administración."
  else
    if OK then
      dbExecute("DELETE FROM RESERVAS WHERE ID=" & bookingId & " AND (ID_VECINO=" & usrId & " OR ID_VECINO_2=" & usrId & ")")
      JSONAddOpOK
      logActivity "Cancela reserva " & resource, bookingDate & "&nbsp;" & turn
    else
      JSONAddOpFailed
      JSONAddMessage resource & ": el turno " & bookingDate & " " & turn & "\n\nYa no es posible cancelar la reserva porque la anticipación mínima es de 2 horas."
      logActivity "Cancela reserva " & resource, bookingDate & " " & turn & ", denegada: anticipación insuficiente."
    end if
  end if
  JSONSend
end function

function bookingsTake(resourceId, bookingDate, turnStart, turnType, extraNeighborUnit)
  if isNumeric(resourceId) and isNumeric(turnStart) and isNumeric(turnType)  then 
    if not getUsrData then exit function
    dbGetData("SELECT DURACION, REQUIERE_VECINO_2 FROM RESERVAS_TIPOS WHERE ID_RECURSO=" & resourceId & " AND ID=" & turnType)
    dim turnDuration: turnDuration = rs(0)
    dim extraNeighborRequired: extraNeighborRequired = rs(1)
    dbReleaseData
    dbGetData("SELECT dbo.NOMBRE_RECURSO_RESERVA(" & resourceId & ") AS RECURSO, dbo.NOMBRE_TURNO(" & turnStart & ", " & turnDuration & ") AS TURNO")
    dim resource: resource = rs("RECURSO")
    dim turn: turn = rs("TURNO")
    dbReleaseData
    dim OK
    dim sqlDate: sqlDate = sqlValue(dbDatePrefix & bookingDate)
    dbGetData("SELECT COUNT(*) FROM RESERVAS WHERE (" & _
      resourceId & " IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") AND " & _
      "ID_RECURSO IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") OR ID_RECURSO=" & resourceId & ") AND " & _
      "FECHA_FIN > GETDATE() AND (ID_VECINO=" & usrId & " OR ID_VECINO_2=" & usrId & ")")
    OK = (rs(0) = 0)
    dbReleaseData
    if OK then
      dbGetData("SELECT CAST(CASE WHEN GETDATE() <= dbo.FECHA_INICIO_TURNO(" & sqlDate & "," & turnStart & ") THEN 1 ELSE 0 END AS BIT)")
      OK = rs(0)
      dbReleaseData
      if OK then
        dbGetData("SELECT dbo.ESTADO_TURNO_RESERVA(" & resourceId & ", " & sqlDate & ", " & turnStart & ", " & turnDuration & ")")
        dim status: status = rs(0)
        dbReleaseData
        if uCase(status) = "DISPONIBLE" then
          dim extraNeighborId: extraNeighborId = null
  			  if resourceId = tennisCourt1BookingResourceId or resourceId = tennisCourt2BookingResourceId then
  					dbGetData("SELECT CAST(CASE WHEN DATEDIFF(DAY, GETDATE(), " & sqlDate & ") = 0 OR DATEPART(HOUR, GETDATE()) >= 8 THEN 1 ELSE 0 END AS BIT)")
              OK = rs(0)
              dbReleaseData
          end if
    	    if OK then
            if extraNeighborRequired then
              if not isNull(extraNeighborUnit) then 
                dbGetData("SELECT ID FROM VECINOS WHERE UNIDAD=" & sqlValue(extraNeighborUnit) & "")
                extraNeighborId = rs(0)
                dbReleaseData
                OK = not isNull(extraNeighborId)
                if OK then
                  OK = extraNeighborId <> usrId
                  if OK then
                    dbGetData("SELECT COUNT(*) FROM RESERVAS WHERE (" & _
                      resourceId & " IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") AND " & _
                      "ID_RECURSO IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") OR ID_RECURSO=" & resourceId & ") AND " & _
                      "FECHA_FIN > GETDATE() AND (ID_VECINO=" & extraNeighborId & " OR ID_VECINO_2=" & extraNeighborId & ")")
                    OK = (rs(0) = 0)
                    dbReleaseData
                    if not OK then
                      JSONAddOpFailed
                      JSONAddMessage resource & ": ya existe una reserva a utilizar para el Número de Lote " & extraNeighborUnit & "."
                      logActivity "Reserva compartida " & resource, bookingDate & " " & turn & ", denegada: el otro vecino ya tiene una reserva a utilizar "
                    end if
                  else
                    JSONAddOpFailed
                    JSONAddMessage resource & ": el Número de Lote para reserva compartida no puede ser el propio."
                    logActivity "Reserva compartida " & resource, bookingDate & " " & turn & ", denegada: Número de Lote no puede ser el propio"
                  end if
                else
                  JSONAddOpFailed
                  JSONAddMessage resource & ": el Número de Lote para reserva compartida es incorrecto."
                  logActivity "Reserva compartida " & resource, bookingDate & " " & turn & ", denegada: Número de Lote incorrecto"
                end if
              else
                JSONAddOpFailed
                JSONAddMessage resource & ": falta el Número de Lote para reserva compartida."
                logActivity "Reserva compartida " & resource, bookingDate & " " & turn & ", denegada: falta Número de Lote"
              end if
            else
              if turnType = tennisClassBookingTypeId then
                dbGetData("SELECT CAST(CASE WHEN (SELECT COUNT(*) FROM RESERVAS WHERE " & _
                  resourceId & " IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") AND " & _
                  "ID_RECURSO IN (" & tennisCourt1BookingResourceId & ", " & tennisCourt2BookingResourceId & ") AND " & _
                  "ID_TIPO = " & tennisClassBookingTypeId & " AND FECHA=" & sqlDate & " AND INICIO=" & turnStart & ")=1 AND " & _ 
                  "(DATEPART(WEEKDAY, " & sqlDate & ") IN (1, 7) OR DATEPART(HOUR, " & sqlDate & ") BETWEEN 19 AND 21) " & _
                  "THEN 1 ELSE 0 END AS BIT)")
                OK = not rs(0)
                dbReleaseData
              end if
              if not OK then
                JSONAddOpFailed
                JSONAddMessage resource & ": de lunes a viernes de 19 a 21hs, y fines de semana, se permite una clase por turno."
                logActivity "Reserva de clase " & resource, bookingDate & " " & turn & ", denegada: se permite una clase por turno."
              end if
            end if
  			  else
      			JSONAddOpFailed
      			JSONAddMessage resource & ": Las reservas para mañana se pueden efectuar recién a partir de las 8:00hs."
      			logActivity "Reserva día siguiente " & resource, bookingDate & " " & turn & ", denegada: se permite a partir de las 8.00 "
    			end if
          if OK then
            dbExecute("INSERT INTO RESERVAS (REC_ID_USUARIO, ID_RECURSO, ID_VECINO, FECHA, INICIO, DURACION, ID_TIPO, ID_VECINO_2) VALUES (" & _
              "1, " & resourceId & ", " & usrId & ", " & sqlDate & ", " & turnStart & ", " & turnDuration & ", " & turnType & ", " & sqlValue(extraNeighborId) & ")")
            if failed then
              JSONAddOpFailed
              JSONAddMessage "Error interno al grabar la reserva."
              logActivity "Reserva " & resource, bookingDate & " " & turn & ", denegada: error interno"
            else
              JSONAddOpOK
              
            end if
          end if
        else
          JSONAddOpFailed
          JSONAddMessage resource & ": el turno " & bookingDate & " " & turn & " no está disponible."
          logActivity "Reserva " & resource, bookingDate & " " & turn & ", denegada: turno no disponible" 
        end if
      else
        JSONAddOpFailed
        JSONAddMessage resource & ": el turno " & bookingDate & " " & turn & " ya no se puede reservar porque ha pasado el horario de comienzo."
        logActivity "Reserva " & resource, bookingDate & " " & turn & ", denegada: turno pasado o iniciado"
      end if
    else
      JSONAddOpFailed
      JSONAddMessage "Ya hay una reserva a utilizar.\nUna vez que la utilices podrás realizar otra."
      logActivity "Reserva " & resource, bookingDate & " " & turn & ", denegada: ya tiene una reserva."
    end if
  else 
    JSONAddOpFailed
    JSONAddMessage "Envio de parametros erroneo.\nElija una reserva del tablero."
    logActivity "Reserva " & resource, bookingDate & " " & turn & ", Datos erroneos: TurTyp:"&turnType &" ExVec:"&extraNeighborUnit& " ."
  end if 
  JSONSend
end function

function bookingsSendRequest(bookingDate, turnId)
  if not getUsrData then exit function
  dim sqlDate: sqlDate = sqlValue(dbDatePrefix & bookingDate)
  dbGetData("SELECT CAST(CASE WHEN DATEDIFF(D, (SELECT TOP 1 FECHA FROM HOY), " & sqlDate & ") <= 90 THEN 1 ELSE 0 END AS BIT)")
  dim dateOK: dateOK = rs(0)
  dbReleaseData
  if dateOK then
    dim turnStart, turnDuration
    dbGetData("SELECT INICIO, DURACION FROM RESERVAS_TURNOS_ESPECIALES WHERE ID_RECURSO=" & clubHouseBookingResourceId & " AND ID=" & turnId)
    turnStart = rs("INICIO")
    turnDuration = rs("DURACION")
    dbReleaseData
    dbGetData("SELECT dbo.ESTADO_TURNO_RESERVA(" & clubHouseBookingResourceId & ", " & sqlDate & ", " & turnStart & ", " & turnDuration & ")")
    dim status: status = rs(0)
    dbReleaseData
    if uCase(status) = "DISPONIBLE" then
      dbGetData("SELECT UNIDAD, EMAIL FROM VECINOS WHERE ID=" & usrId)
      dim usrUnit: usrUnit = rs("UNIDAD")
      dim usrEmail: usrEmail = rs("EMAIL")
      dbReleaseData
      dbGetData("SELECT dbo.NOMBRE_TURNO(" & turnStart & ", " & turnDuration & ")")
      dim turnName: turnName = rs(0)
      dbReleaseData
      dim message: message = "<table border=1>" & _
        "<tr><td>Unidad/Lote:</td><td>" & usrUnit & "</td></tr>" & _
        "<tr><td>Familia:</td><td>" & usrName & "</td></tr>" & _
        "<tr><td>Fecha del evento:</td><td>" & bookingDate & "</td></tr>" & _
        "<tr><td>Turno:</td><td>" & turnName & "</td></tr>" & _
        "</table>"
      sendMail "Vecinos de los Sauces", "info@vecinosdelossauces.com.ar", "Solicitud de reserva del Club House", message, "Familia " & usrName, usrEmail
      logActivity "Solicitud reserva Club House", "Enviada"
      exit function
    else
      JSONAddOpFailed
      JSONAddMessage "El turno indicado no se encuentra disponible."
      logActivity "Solicitud reserva Club House", "Denegada, turno no disponible."
    end if
  else
    JSONAddOpFailed
    JSONAddMessage "Las reservas del Club House deben solicitarse con una anticipación máxima de 90 días."
    logActivity "Solicitud reserva Club House", "Denegada, anticipación excesiva."
  end if
  JSONSend
end function
function jBookingStartOptions(resourceId,bookingDate)
  if isNull(resourceId) then
    JSONAddStr "bookingTurnOptions", "[]"
  else
    dim dia
    if isnull(bookingDate) then 
      dia = 1
    else
      dia = Weekday(bookingDate)
    end if
    dim b: b = dbConnect
    dbGetData("SELECT ID, NOMBRE FROM RESERVAS_TURNOS_ESPECIALES where ID_RECURSO="& resourceId & " AND dias like '%"&dia&"%' ORDER BY ID")
    
    JSONAddArray "bookingStartOptions", rs.getRows, array("id", "name")
    dbReleaseData
    if b then dbDisconnect
  end if
  JSONSend
end function
%>
