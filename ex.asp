<system.web>
   <compilation debug="true" targetFramework="4.5" />
   <httpRuntime targetFramework="4.5" requestPathInvalidCharacters="" />
</system.web>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Respuesta de la consulta</TITLE>

 
</HEAD>



<%@ LANGUAGE=VBScript%>
<%


barcode=Request.form("T1")
set conexion = Server.CreateObject("ADODB.Connection")

'Server=tcp:supermode.database.windows.net,1433;Initial Catalog=SUPERMODE;Persist Security Info=False;User ID=supermode;Password={your_password};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;
conexion.ConnectionString = "Driver='SQL Server Native Client 11.0';" & _
"Server='tcp:supermode.database.windows.net,1433';" & _
"Database='supermode';" & _
"Uid='supermode@supermode';" & _
"Pwd='bAdCAj7Bbnncz57';" & _
"Encrypt='yes';Connection Timeout='30';"



'conexion.ConnectionString = "Server=tcp:supermode.database.windows.net,1433;Initial Catalog=SUPERMODE;Persist Security Info=False;User ID=supermode;Password=bAdCAj7Bbnncz57;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;
conexion.Open

	
	dim oShell
	dim booOferta		'Selector
	dim booPromo	    'Selector	
	dim strDesc_c		'Descripcion corta
	dim strUnidad		'Unidad de venta
	dim strArticulo		'Código R3
	dim dblP_normal		'Precio normal
	dim dblP_cdcto		'Precio c/descuento
	dim intPr_dsc		'Porcentaje de descuento
	dim dblAhorro		'Ahorro en pesos
	dim datIniOferta	'Fecha de inicio de oferta
	dim datFinOferta	'Fecha de fin de oferta
	dim strCategoria	'Categoria del articulo
	dim datFecha		'La fecha actual
	dim datTest
	dim typFecha
	dim sql,sql2
	dim strMensaje
	dim tsql
	dim lv_vProm,lv_ODiario,lv_Objetivo,lv_OMensual,lv_ObjetivoMensual
'	function ExisteArticulo() 
		datFecha =  Mid(date(), 7,4) &"-" & Mid(date(), 4,2) &"-"& Left(date(), 2)'date()
		sql = "SELECT * from temp_ventas where fecha ='"& datFecha &"' order by sucursal"
		
		set rs=conexion.Execute(sql)
		'if ( not rs.eof) then		
		'strMensaje = sql
		strMensaje = strMensaje & "<Table border = 1><TR><TD>Sucursal<TD>Venta<TD>Clientes<TD> Ticket Promedio<TD>Objetivo Diario<TD>Cubrimiento Objetivo Diario<TD>Objetivo Mensual<TD>Cubrimiento Objetivo Mensual"
		while not rs.eof
			sql2 = "SELECT * from presupuestos where ano ='"& Mid(date(), 7,4) &"' and mes ='"&Mid(date(), 4,2)&"' and Sucursal ='"&rs.fields("sucursal")&"'" 	
			set rs2=conexion.Execute(sql2)	
			lv_vProm = round(rs.fields("ventadiaria") /  rs.fields("clientes") ,2)	
			lv_ODiario = round(rs2.fields("Presupuesto") / 30,2)
			lv_Objetivo = round((rs.fields("ventadiaria")/lv_ODiario) * 100,2)
			lv_OMensual = round(rs2.fields("Presupuesto"),2)
			lv_ObjetivoMensual = round((rs.fields("ventamensual")/lv_OMensual) * 100,2)
			
				strMensaje = strMensaje  & "<TR><TD>" &rs.fields("sucursal") & "<TD>" & FormatCurrency(rs.fields("ventadiaria")) & "<TD>" & rs.fields("clientes") & "<TD>$" & lv_vProm & "<TD>" & lv_ODiario & "<TD>" & lv_Objetivo&"%"&"<TD>"&FormatCurrency(lv_OMensual)&"<TD>"&lv_ObjetivoMensual&"%"
		rs.movenext
		wend
		strMensaje = strMensaje & "</Table>"
			ExisteArticulo=0
		'end if
'	end function

%>



<BODY BACKGROUND="MarcaAguadorada.JPG">

		<CENTER><FONT SIZE="2" COLOR="#000066"><B><h1>Ventas Al momento <%=datFecha%></h1> <%=strMensaje%></B></FONT></CENTER><BR>	

</BODY>
</HTML>