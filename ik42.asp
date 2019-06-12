<%response.cookies("ikmenu") = "ik42.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>khbp4002</title>
<script language="javascript">
	function invalid_data(b)
	{
		alert(b+" Kod telah wujud");
		return(true);
	}
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">

<table border="0" width="100%" height="15">
<tr><td width="100%" height="15"></td></tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolorlight="#003366" >
<tr valign="top"> 
          
<td width="100%"> 
<form method ="Post" action="ik42.asp">

<% 
	Set objConn = Server.CreateObject ("ADODB.Connection")
  	ObjConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

	proses = Request.Form("B1")
	response.cookies("amenu") = "ik42.asp"
	
	kod1 = Request.Form("kod1")
	kod11 = mid(kod1,6,3)
	njab1 = mid(kod1,9,60)

	kod2 = Request.Form("kod2")
	kod22 = mid(kod2,6,60)
			q = "select substr('"& kod22 &"',1,instr('"& kod22 &"','-',1,1)-1) as dkod, "
			q = q & " substr('"& kod22 &"',instr('"& kod22 &"','-',1,1)+1,50) as djab from dual "
   			Set objRsq = Server.CreateObject("ADODB.Recordset")
   			Set objRsq = objConn.Execute(q)

	dkod1 = objRsq("dkod")
	djab = objRsq("djab")


 '************************************** PROSES RESET ************************************************
 

	if proses = "reset" then
		response.redirect"ik42.asp"
	end if
	
	

 '************************************** PROSES RES ************************************************
	
	
	if proses = "res" then
	vkod = ucase(Request.Form("fkod"))
	vketer = ucase(Request.Form("fketer"))
	vrowid = request.form("rowid")

		vkod = ""
		vketer = ""
		vrowid = ""
	end if
	
	
	
'****************************************************************************************************
	
	
	if proses <> "pilih" or proses <> "+" then
		bilrec = request.form("bilrec")
	else
		bilrec = ""
	end if
	
	if proses = "+" then
	vkod = Request.Form("fkod")
	vketer = Request.Form("fketer")
	vrowid = request.form("rowid")
	
		if vrowid = "" then			
			if vkod <> "" then
				if vketer <> "" then
				
				a11 = " select kod,keterangan from kompaun.jenis_kesalahan where kod like '"&vkod&"' and perkara like '"&kod11&"' "
				
				Set objRs11a = Server.CreateObject("ADODB.Recordset")
				Set objRs11a = objConn.Execute(a11)
				if not objRs11a.eof then
				
				response.write "<script language=""javascript"">"
				response.write "var timeID = setTimeout('invalid_data("" "");',1)"
				response.write "</script>"

				else 
				
				a1 = " insert into kompaun.jenis_kesalahan(perkara,kod,keterangan) values ('"&kod11&"','"&vkod&"','"&vketer&"') "
	
				Set objRs1a = Server.CreateObject("ADODB.Recordset")
				Set objRs1a = objConn.Execute(a1)
				end if
				end if
			end if
		else
		
		b1 = " update kompaun.jenis_kesalahan "
		b1 = b1 & " set kod = '"&vkod&"', "
		b1 = b1 & "     keterangan = '"&vketer&"' "
		b1 = b1 & " where rowid = '"&vrowid&"' "
		
		Set objRs1b = Server.CreateObject("ADODB.Recordset")
		Set objRs1b = objConn.Execute(b1)	
		end if
		proses1 = "KESALAHAN"
		proses2 = "pilih"
		
	end if	
	
	
	if bilrec <> "" then
     for i = 1 to bilrec
         rowid = "rowid" + cstr(i)
         kod = "kodz" + cstr(i)
         keter = "keter" + cstr(i)
         b1 = "b1" + cstr(i)
         b2 = "b2" + cstr(i)
        
         rowid = request.form(""& rowid &"")
         vkod = request.form(""& kod &"")
         vketer = request.form(""& keter &"")
         b1 = request.form(""& b1 &"")
         b2 = request.form(""& b2 &"")
       

         if b1 = "-" then
     
            db = " delete kompaun.jenis_kesalahan where rowid = '"& rowid &"' "
            Set objRsdb = Server.CreateObject("ADODB.Recordset")
            Set objRsdb = objConn.Execute(db)
         elseif b2 = "e" then 
         		fkod = vkod
         		fkara = vkara
         		fketer = vketer
         		frowid = rowid
          		proses4 = "+"
          		
				
         end if
         proses1 = "KESALAHAN"
         proses2 = "pilih"
     next
  end if
  
	
	
	if proses = "pilih" then
		proses2 = "pilih"
	end if
		
%>
	
<%		if proses = "" then		%>
<table width="98%" align="center">
<tr>  
<td width="10%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>Akta/UUK</b></font></td>
<td width="88%" bgcolor="lightgrey"> 
<select size="1" name="kod1">

<%  if kod11 <> "" then 	%>  
 
<option value="kod1=<%=kod11%><%=njab1%>"><%=kod11%> - <%=njab1%></option>
<%	end if
 
  	zz = "select kod kod1,keterangan from kompaun.perkara order by kod "
  	Set objRszz = objConn.Execute(zz)
  	Do While Not objRszz.EOF 
%>
<option value="kod1=<%=objRszz("kod1")%><%=objRszz("keterangan")%>"><%=objRszz("kod1")%>
- <%=objRszz("keterangan")%></option>
<% 
  objRszz.MoveNext
  loop 
%>
</select>
<input type="submit" value="pilih" name="B1" ><input type="submit" value="reset" name="B1" ></td></tr> 
</table>

<%	else	%>

<% if kod11 <> "" then 		%> 
<table width="98%" align="center">
<tr >  
<td width="10%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>Akta/UUK</b></font></td>  
<td width="88%" bgcolor="lightgrey"> 
<select size="1" name="kod1">
<option value="kod1=<%=kod11%><%=njab1%>"><%=kod11%> - <%=njab1%></option>
<%	 zz = "select kod kod1,keterangan from kompaun.perkara order by kod "
  	 Set objRszz = objConn.Execute(zz)
  	 
  	 Do While Not objRszz.EOF 
%>
<option value="kod1=<%=objRszz("kod1")%><%=objRszz("keterangan")%>"><%=objRszz("kod1")%>
- <%=objRszz("keterangan")%></option>
<%	objRszz.MoveNext
  	loop 	%>
</select>
<input type="submit" value="pilih" name="B1" ><input type="submit" value="reset" name="B1" ></td></tr> 
</table>
<%	end if
	end if		%>

<%
	if proses2 = "pilih" then
	
		
	b = " select rowid, kod, initcap(keterangan) keterangan, perkara "
	b = b & " from kompaun.jenis_kesalahan where perkara like '"&kod11&"' order by kod"
	Set objRs2 = Server.CreateObject("ADODB.Recordset")
	Set objRs2 = objConn.Execute(b)
%>
	
<table width="80%" align="center">
<tr > 
<td width="10%" bgcolor="#FFFFFF" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Kod</b></font></td>
<td width="65%" bgcolor="#FFFFFF" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Keterangan</b></font></td>
<td width="5%" bgcolor="#FFFFFF" align="center"></td>
</tr>

<%if proses4 <> "" then%>

<tr > 
<td width="10%" bgcolor="#FFFFFF" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2">
<input type="text" name="fkod" size="8" value="<%=fkod%>" maxlength="10"></font></td>
<td width="65%" bgcolor="#FFFFFF" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2">
<input type="text" name="fketer" size="70" value="<%=fketer%>" maxlength="70"></font></td>
<td width="5%" bgcolor="#FFFFFF" align="center">
<input type="submit" value="+" name="B1"><input type="submit" value="res" name="B1">
<input type="hidden" name="rowid" value="<%=frowid%>" ></td>
</tr>

<%else%>

<tr > 
<td width="10%" bgcolor="#FFFFFF" align="center"><input type="text" name="fkod" size="8" maxlength="10"></td>
<td width="65%" bgcolor="#FFFFFF" align="center"><textarea rows="4" cols="50" name="fketer"></textarea></td>
<td width="5%" bgcolor="#FFFFFF" align="center">
<input type="submit" value="+" name="B1"><input type="submit" value="res" name="B1">
</td>
</tr>
</table>

<%end if
	  
  if not objRs2.eof then	%>

<table width="80%" align="center" border="0"  bgcolor="#330000" cellspacing="1">
<tr  bgcolor="#99FF99"> 
<td width="10%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Kod</b></font></td>
<td width="58%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Keterangan</b></font></td>
<td width="5%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Edit</b></font></td>
<td width="7%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Hapus</b></font></td>
</tr>
<%
	bil = 0
	do while not objRs2.eof
	bil = bil + 1
%>
<tr> 
<td width="10%" bgcolor="lightgrey" align="center">
<font face="Verdana" size="2"><%=objRs2("kod")%></font></td>
<td width="58%" bgcolor="lightgrey" align="left">
<font face="Verdana" size="2"><%=objRs2("keterangan")%></font></td>
<td width="5%" bgcolor="lightgrey" align="center">
<input type="submit"  name="b2<%=bil%>" value="e" >
<td width="7%" bgcolor="lightgrey" align="center">
<input type="submit" onClick="return confirm('Hapus Satu Data ?')" name="b1<%=bil%>" value="-" >
<input type="hidden" name="rowid<%=bil%>" value="<%=objRs2("rowid")%>" >
<input type="hidden" name="kodz<%=bil%>" value="<%=objRs2("kod")%>" >
<input type="hidden" name="keter<%=bil%>" value="<%=objRs2("keterangan")%>" ></td>
</tr>
<%	objRs2.movenext
	loop		%>
</table>

<input type="hidden" name="bilrec" value="<%=bil%>" >
<%	end if
	end if		%>
</form>

</td>
</tr>
</table>
</body>
</html>

