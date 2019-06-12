<!-- '#INCLUDE FILE="ik.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Kompaun Sudah Bayar</title>
</head>
<script>
function invalid_rekod(a)
{ alert (a+" Tiada Rekod");
  return(true); }
</script>
<body>

<%

'Set objConn = Server.CreateObject("ADODB.Connection")
'  objconn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  'tkh1 = request.cookies("tkh1")
  'tkh2 = request.cookies("tkh2")
  
  'response.write "<p>"&tkh1&"</p>"
  'response.write "<p>"&tkh2&"</p>"
  tkh1 = Request.form("ftkh1")
  tkh2= Request.form("ftkh2")
  tkh1a = cstr(mid(tkh1,1,2)) + "/" + cstr(mid(tkh1,3,2)) + "/" + cstr(mid(tkh1,5,4))
  tkh2a = cstr(mid(tkh2,1,2)) + "/" + cstr(mid(tkh2,3,2)) + "/" + cstr(mid(tkh2,5,4))
  
  z = "select to_char(sysdate,'dd/mm/yyyy hh24:mi:ss') tkh from dual"
  Set oz = objConn.Execute(z)
   
  		b = " select to_date('"&tkh1&"','ddmmyyyy') as tkha,"
		b = b & " to_date('"&tkh2&"','ddmmyyyy') as tkhb from dual "
		Set objRsb = Server.CreateObject ("ADODB.Recordset")
   		Set objRsb = objConn.Execute(b)
   		
   		if objRsb.eof then
   		response.write "<script language=""javascript"">"
       response.write "var timeID = setTimeout('invalid_tarikh(""  "");',1) "
       response.write "</script>"
       
       else
       tkha = objRsb("tkha")
       tkhb = objRsb("tkhb")
       
       if tkha > tkhb then
       response.write "<script language=""javascript"">"
       response.write "var timeID = setTimeout('invalid_tarikh(""  "");',1) "
       response.write "</script>"      
   		
   		else  		
		
		d = " select no_akaun,perkara, no_rujukan2, nama, to_char(tkh_masuk,'ddmmyyyy') tkh_masuk,jabatan, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'ddmmyyyy') tkh_bayar, no_resit "
		d = d & " from hasil.bil "
		d = d & " where tkh_bayar between to_date('"& tkh1 &"', 'ddmmyyyy') "
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (perkara <> 'P01' or perkara is null) "
		'd = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%') "
		'18042014 : jun tambah boleh papar kod 76413 jugak, sebelum ni 76410 & 76420 saja
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%') "
		d = d & " and amaun_bayar > 0 "
	 	'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		
		
		pekz = request.cookies("gnop")
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		'response.Write(admin)
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		d = d & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
		d = d & " order by tkh_bayar "
		Set objRs2 = objConn.Execute(d)
		
%>

<table align = "center" border="0" width="100%" cellspacing="1" >
<tr>
  <td width="26%"><font face="Trebuchet MS" size="1"><%=oz("tkh")%></font></td>
  <td width="74%" align="right"><font face="Trebuchet MS" size="1">Muka : 1<%'=muka%></font></td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      
    <td align="center"><b><font size="3" face="Trebuchet MS">MAJLIS PERBANDARAN 
      SEBERANG PERAI</font></b></td>
    </tr>
    <tr>
      
    <td align="center"><b><font size="4" face="Trebuchet MS"> <font size="3">SENARAI 
      KOMPAUN SUDAH BAYAR DARI TARIKH<br>
      <%=tkh1%>&nbsp;HINGGA&nbsp;<%=tkh2%></font></font></b></td>
    </tr>
  </table>
  <br>

<hr>

<table border="0" width="100%">
  <tr bgcolor="#FFFFFF" > 
    <td width="4%" align="center"><b><font size="2" face="Trebuchet MS">Bil</font></b></td>
    <td width="13%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">No 
      Kompaun</font></b></td>
    <td width="7%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Akta</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Jenis&nbsp;</font></b></td>
    <td width="30%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Nama</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Tarikh&nbsp;</font></b></td>
    <td width="11%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Tkh 
      Bayar&nbsp;</font></b></td>
    <td width="9%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">No 
      Resit</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Amaun</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Jabatan</font></b></td>
  </tr>
  <%	
	bil = 0
	ctrz = 0
	
	do while not objRs2.eof
	bil = bil + 1
	ctrz = cdbl(ctrz) + 1
	abayar = objRs2("amaun_bayar")
			
	jamaun = cdbl(jamaun) + cdbl(abayar)
	
				kodJbtn = objRs2("jabatan")
		
		q1="select keterangan from payroll.ptj where kod='"&kodJbtn&"'"
		set rq1 = objConn.execute(q1)
		
		if not rq1.eof then namaJabatan = rq1("keterangan")
		   %>
  <tr bgcolor="#FFFFFF" > 
    <td><font color="#000000" size="2" face="Trebuchet MS">&nbsp;<%=bil%></font></td>
    <td width="10%" align="center"><font color="#000000" size="2" face="Verdana"><%=objRs2("no_akaun")%></font></td>
    <%	kara = objRs2("perkara") 
  		if kara <> "" then
  %>
    <td width="7%" align="center" > <font size="2" face="Verdana" color="#000000"><b><%=objRs2("perkara")%></b></font></td>
    <%	else	%>
    <td width="7%" align="center"> <font size="2" face="Verdana" color="#000000"><b><%=objRs2("perkara")%></b></font></td>
    <%	end if		
  		rujuk2 = objRs2("no_rujukan2")
  		if rujuk2 <> "" then
  %>
    <td width="10%" align="center" > 
      <font size="2" face="Verdana" color="#000000"><b><%=objRs2("no_rujukan2")%></b></font></td>
    <%	else		%>
    <td width="10%" align="center">
      <font size="2" face="Verdana" color="#000000"><b><%=objRs2("no_rujukan2")%></b></font></td>
    <%	end if		%>
    <td width="28%" align="center"><font size="2" face="Verdana">
      <p align="left"><font color="#000000"><%=objRs2("nama")%></font></font></td>
    <td width="11%" align="center"><font color="#000000" size="2" face="Verdana"><%=objRs2("tkh_masuk")%> 
      </font></td>
    <td width="10%" align="center"><font color="#000000" size="2" face="Verdana"><%=objRs2("tkh_bayar")%> 
      </font></td>
    <td width="10%" align="center"><font color="#000000" size="2" face="Verdana"><%=objRs2("no_resit")%></font></td>
    <td width="9%" align="center"><font color="#000000" size="2" face="Verdana"><%=FormatNumber(objRs2("amaun_bayar"),2)%></font></td>
      <td><%=namaJabatan%></td>
  </tr>
  <%	objRs2.MoveNext			
  	Loop
%>
  <tr bgcolor="#FFFFFF" > 
    <td colspan="2" align="center">&nbsp;</td>
    <td colspan="3" align="right"><b><font size="2" face="Trebuchet MS">&nbsp; 
      Jumlah :&nbsp;&nbsp;</font></b></td>
    <td align="right"><b><font size="2" face="Trebuchet MS"><%=formatnumber(jamaun,2)%>&nbsp;&nbsp;&nbsp;</font></b></td>
  </tr>
</table>
<%
end if
end if%>
<hr>
</body>

</html>