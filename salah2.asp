<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Jenis Kesalahan</title>
</head>

<body>
<form method="POST" action = "salah2.asp" >

<%
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
	aktx = Request.QueryString("rujuk")
	jenisx = Request.QueryString("rujuk1")    
  
    a = " select perkara, kod, initcap(keterangan) keterangan from kompaun.jenis_kesalahan "
    a = a & " where perkara = '"&aktx&"' order by kod "
    Set objRsa = Server.CreateObject("ADODB.Recordset")
    Set objRsa = objConn.Execute(a)
    
    b = " select initcap(keterangan) keterangan from kompaun.perkara where kod = '"&aktx&"' "
    Set objRsb = Server.CreateObject("ADODB.Recordset")
    Set objRsb = objConn.Execute(b)
    keter = objRsb("keterangan")
 %>

<p align="center"><b><font color="#003300" face="MS Serif" size="3">
Senarai Jenis Kesalahan Bagi&nbsp;<%=keter%>(<%=aktx%>)</font></b></p>


<table align = "center" border="0" width="90%" bordercolor="#FFFFFF" cellspacing="1" bgcolor="#333300">
<tr bgcolor="#006633"> 
<td width="5%" align="center" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF"> 
<b><font face="MS Serif" color="#FFFFFF">Bil</font></b></td>
<td width="10%" align="center" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF"> 
<b><font face="MS Serif" color="#FFFFFF">Perkara</font></b></td>
<td width="10%" align="center" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF"> 
<b><font face="MS Serif" color="#FFFFFF">Kod</font></b></td>
<td width="65%" align="center" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF"> 
<b><font face="MS Serif" color="#FFFFFF">Keterangan</font></b></td>
</tr>
<%	
	bil = 0

	Do while not objRsa.EOF
	bil = bil + 1
	
	kod = objRsa("kod")
	if kod = jenisx then
%> 
<tr> 
<td width="5%" align="center" bgcolor="#CCFFCC" height="25"><font face="Century Gothic" size="2"><%=bil%></font></td>
<td width = "10%" align="center" bgcolor="#CCFFCC" height="25"><font face="Century Gothic" size="2"><%=objRsa("perkara")%></font></td>
<td width = "10%" align="center" bgcolor="#CCFFCC" height="25"><font face="Century Gothic" size="2"><%=objRsa("kod")%></font></td>
<td width="65%" bgcolor="#CCFFCC"><font face="Century Gothic" size="2">&nbsp;&nbsp;<%=objRsa("keterangan")%></font></td>
</tr>
<%	else	%>
<tr> 
<td width="5%" align="center" bgcolor="lightgrey" height="25"><font face="Century Gothic" size="2"><%=bil%></font></td>
<td width="10%" align="center" bgcolor="lightgrey" height="25"><font face="Century Gothic" size="2"><%=objRsa("perkara")%></font></td>
<td width="10%" bgcolor="lightgrey" align="center" height="25"><font face="Century Gothic" size="2"><%=objRsa("kod")%></font></td>
<td width="65%" bgcolor="lightgrey"><font face="Century Gothic" size="2">&nbsp;&nbsp;<%=objRsa("keterangan")%></font></td>
</tr>
<%	end if

    objRsa.Movenext  
    Loop
%> 
</table>
</form>

<form action="javascript:history.back(-1);"><p align="center"><input type="submit" value="BACK" name="B2"></form>

</body>
</html>

