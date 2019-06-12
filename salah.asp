<% Response.buffer = "True"	%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Jenis Kesalahan</title>
<script language="javascript">
function invalid_fakt(a)
    {  
       alert (a+" Sila Pilih Akta/UUK Dahulu !!! ");
		return(true);
    }
</script>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFCCCC">

<form method="POST" action = "salah.asp" >

<%
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
    
    
	menu = Request.cookies("kmenu")
	proses = Request.Form("B2")
	nkod = Request.querystring("nkod")
	fakt = ucase(Request.cookies("fakt"))
	kodx = Request.form("kodx")
	response.cookies("fasp") = ""
	
	if fakt = "" then
			response.write "<script language=""javascript"">"
   			response.write "var timeID = setTimeout('invalid_fakt(""  "");',1) "
       	response.write "</script>"
   else	
%>

<p align="center">
<font face="Verdana" size="2"><b>Kod Kesalahan :</b></font>
<input type="text" name="kodx" value="<%=kodx%>" size=10><input type="submit" name="B2" value="Hantar" style="font-family: Verdana">
<%  
  '************************************	PROSES PILIH JENIS KESALAHAN	***********************************	
	

	fjenis = Request.form("jenis")
	biljenis = Request.form("biljenis")
	
	
	if biljenis <> "" then
  
     for j = 1 to biljenis
     
     
     	fjenis1 = "jenis" + cstr(j)
       B2 = "B2" + cstr(j)
       
   		fjenis2 = request.form(""& fjenis1 &"")
       B2x = request.form(""& B2 &"")


       if B2x = "p" then
       	response.cookies("fjenis") = fjenis2
       	response.cookies("fasp2") = "salah.asp"
       	if menu = "ik11.asp" then
	         	response.redirect "ik11.asp"
	         	
	       elseif menu = "ik11b.asp" then
   	       	response.cookies("pilihakta") = "akta"
	       	response.cookies("pilihsalah") = "salah"
	       	response.redirect "ik11b.asp"		
  			end if	       	  	
       else 
       	fjenis = ""
       	biljenis = ""
       end if
     next
  end if
	
 
  
 if proses = "Hantar" then
 
    a = " select rowid, perkara, kod, initcap(keterangan||' '||keterangan2) keterangan from kompaun.jenis_kesalahan "
    a = a & " where kod like '"& kodx &"'||'%'  "
    a = a & " and perkara = '"&fakt&"' order by kod "
    Set objRsa = Server.CreateObject("ADODB.Recordset")
    Set objRsa = objConn.Execute(a)
%>

<table border="0" width="616" cellspacing="1" align="center">
  <tr bgcolor="#9D2024"> 
   <td width="84"  align="center"><b><font face="Verdana" size="2" color="#FFFFFF">Perkara</font></b></td>
   <td width="84" align="center"><b><font face="Verdana" size="2" color="#FFFFFF">Kod</font></b></td>
   <td width="471" align="center"><b><font face="Verdana" size="2" color="#FFFFFF">Keterangan</font></b></td>
   <td width="41"><b><font face="Verdana" size="2" color="#FFFFFF">Pilih</font></b></td>
</tr>
<%	ctrz = 0
	Do while not objRsa.EOF
	
	   ctrz = cdbl(ctrz)  + 1  
%>
<tr>
<td width="84" bgcolor="#C0C0C0" align="center"><font face="Verdana" size="2"><%=objRsa("perkara")%></font></td>
<td width="84" bgcolor="#C0C0C0" align="center"><font face="Verdana" size="2"><%=objRsa("kod")%></font></td>
<td width="471" bgcolor="#C0C0C0"><font face="Verdana" size="2"><%=objRsa("keterangan")%></font>
	<input type="hidden" name="jenis<%=ctrz%>" value="<%=objrsa("kod")%>" >
	<input type="hidden" name="djenis<%=ctrz%>" value="<%=objrsa("keterangan")%>" >
	<input type="hidden" name="drowid<%=ctrz%>" value="<%=objrsa("rowid")%>" >
</td>
<td width="41" bgcolor="#C0C0C0" align="center">
<input type="submit" onClick="this.form.action='<%=nkod%>';" name="B2<%=ctrz%>" value="p"></td>
</tr>
<%  objRsa.MoveNext
    Loop
%>
<input type="hidden" name="biljenis" value="<%=ctrz%>" >
<input type="hidden" name="vrujuk" value="<%=rujuk1%>" >
<input type="hidden" name="aktas" value="<%=dakt%>">

<%	
	end if
	end if
%>

</table>
</form>
</body>
</html>

