<% Response.Buffer = True %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Sistem Perkhidmatan Kaunter</title>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">

<table border="0" width="100%" height="30">
<tr><td width="100%" height="30"></td></tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top">                     
<td width="100%"> 

<form method = "Post" action= "ik213b.asp">

<%
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open = "dsn=12c;uid=majlis;pwd=majlis;"
	
	nakaun = Request.QueryString("kod")
	
	
	b = "select no_akaun,jabatan,kategori,initcap(nama) nama,initcap(alamat1) alamat1, "
	b = b & " initcap(alamat2) alamat2,initcap(alamat3) alamat3,initcap(alamat4) alamat4, "
	b = b & " initcap(perkara1)||lower(perkara2) perkara1,initcap(perkara3) perkara3,no_rujukan,"
	b = b & " no_rujukan2, nvl(amaun,0) as amaun, nvl(amaun_bayar,0) as abayar, no_resit from hasil.bil"
	b = b & " where rowid = '"&nakaun&"' "
	b = b & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
	b = b & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
	b = b & " or no_akaun like '76101'||'%') "
		'18042014 : jun tambah boleh papar kod 76413 jugak, sebelum ni 76410 & 76420 saja
	Set objRs2 = Server.CreateObject("ADODB.Recordset")
	Set objRs2 = objConn.Execute(b) 
	jab = objRs2("jabatan")
	
	
	c = " select initcap(keterangan) terang from payroll.ptj where kod = '"&jab&"' "
	Set objRsc = Server.CreateObject("ADODB.Recordset")
	Set objRsc = objConn.Execute(c)
	
	if not objRsc.EOF then
		terang = objRsc("terang")
	else
		terang = ""
	end if		
%>

      <table align=center width="75%" cellspacing="1" bgcolor="#330000">
        <tr >
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;Jabatan</font></b></td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("jabatan")%> - <%=terang%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;No 
            Akaun</font></b></td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("no_akaun")%></font></td>
</tr>

<tr>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;Nama 
            Pembayar</font></b></td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("nama")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;Alamat</font></b></td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("alamat1")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333">&nbsp;</td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("alamat2")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333">&nbsp;</td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("alamat3")%>&nbsp;<%=objRs2("alamat4")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;Perkara</font></b></td>
<td width="57%" colspan="3" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("perkara1")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333">&nbsp;</td>
<td width="22%" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%'=objRs2("perkara2")%>&nbsp;&nbsp;<%=objRs2("perkara3")%></font></td>
          <td width="18%" bgcolor="#993333"><b><font size="2" face="Verdana" color="#FFFFFF">&nbsp;Kategori</font></b></td>
<td width="17%" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("kategori")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;No 
            Rujukan</font></b></td>
<td width="22%" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("no_rujukan")%></font></td>
          <td width="18%" bgcolor="#993333"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;No 
            Resit</font></b></td>
<td width="17%" bgcolor="lightgrey"><font face="Verdana" size="2">&nbsp;<%=objRs2("no_resit")%></font></td>
</tr>
<tr>
          <td width="18%" bgcolor="#993333"><font face="Verdana" size="2" color="#FFFFFF"><b>&nbsp;Amaun</b></font></td>
<td width="22%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>&nbsp;RM</b>&nbsp;<%=FormatNumber(objRs2("amaun"),2)%></font></td>
          <td width="18%" bgcolor="#993333"><font face="Verdana" size="2" color="#FFFFFF"><b>&nbsp;Amaun 
            Bayar</b></font></td>
<td width="17%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>&nbsp;RM</b>&nbsp;<%=FormatNumber(objRs2("abayar"),2)%></font></td>
</tr>
</table>

</form>
<% 
	s = "select no_akaun from hasil.bil where rowid = '"&nakaun&"' "
	set objRss = Server.CreateObject("ADODB.Recordset")
	set objRss = objConn.Execute(s)
	acc = objRss("no_akaun")
%>
<table width ="100%" align="center">
<tr>
<td width="50%" align="right"><form method = "Post" action= "ha2111c.asp">
<input type="hidden" name="akaunz" value="<%=nakaun%>">
<input type="submit" value="Cetak Salinan" name="B2" > </form></td>
<td width="50%" align="left">
<form action="javascript:history.back(-1);"><input type="submit" value="BACK" name="B2"></form>
</td></tr></table>

</body>




























