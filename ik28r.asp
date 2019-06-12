<!-- '#INCLUDE FILE="ik.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Kompaun Mengikut Tred</title>
</head>

<body onload='self.print()' topmargin="0" leftmargin="0">

 <%
	'Set objConn = Server.CreateObject("ADODB.Connection")
'  	objconn.Open "dsn=12c;uid=majlis;pwd=majlis;"

	tkh1 = request.cookies("tkh1") 
	tkh2 = request.cookies("tkh2") 
	tred = request.cookies("tred") 

	muka = 0
	
	if tred <> "" then
	   papar
	end if

sub papar 

 	z = "select to_char(sysdate,'dd/mm/yyyy hh24:mi:ss') tkh from dual"
  	Set oz = objConn.Execute(z)
	
			rr = " select initcap(keterangan) terangz from lesen.tred "
   		rr = rr & " where kod = '"& tred &"' "
   		Set objRsrr = Server.CreateObject ("ADODB.Recordset")
   		Set objRsrr = objConn.Execute(rr)
		
  	muka = cint(muka) + 1
%>

<table align = "center" border="0" width="100%" cellspacing="0" >
<tr>
  <td><font face="Trebuchet MS" size="1"><%=oz("tkh")%></font></td>
  <td align="right"><font face="Trebuchet MS" size="1">Page : <%=muka%></font></td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
    <td align="center"><b><font size="3" face="Trebuchet MS"> MAJLIS PERBANDARAN 
      SEBERANG PERAI<br>
      LAPORAN KESALAHAN MENGIKUT TRED <%=tred%> - <%=objRsrr("terangz")%><br>
      PADA <%=tkh1%> HINGGA <%=tkh2%><br>
      &nbsp; </font></b></td>
</tr>
</table>
<%		b = " select to_date('"&tkh1&"','ddmmyyyy') as tkha,"
		b = b & " to_date('"&tkh2&"','ddmmyyyy') as tkhb from dual "
		Set objRsb = Server.CreateObject ("ADODB.Recordset")
   		Set objRsb = objConn.Execute(b)
   		
   		if objRsb.eof then
   		response.write "<script language=""javascript"">"
       response.write "var timeID = setTimeout('invalid_date(""  "");',1) "
       response.write "</script>"
 	
		else
 		tkha = objRsb("tkha")
 		tkhb = objRsb("tkhb")	
 		
 
		
		d = " select rowid, no_akaun,perkara, no_rujukan2, nama, to_char(tkh_masuk,'ddmmyyyy') tkh_masuk, "
		d = d & " nvl(amaun_bayar,0) amaun, to_char(tkh_bayar,'ddmmyyyy') tkh_bayar, no_resit "
		d = d & " from hasil.bil "
		d = d & " where tred = '"& tred &"' "
		d = d & " and tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		'd = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%') "
		'18042014 : jun tambah boleh papar kod 76413 jugak, sebelum ni 76410 & 76420 saja
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%') "
		d = d & " and (perkara <> 'P01' or perkara is null) "
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
		d = d & " order by no_akaun "
	

		Set objRs2 = Server.CreateObject ("ADODB.Recordset")
		Set objRs2 = objConn.Execute(d)
end if
	
%>
<table border="0" cellpadding="0" cellspacing="0" width="92%" align="center">
  <tr> 
    <td width="28" align="center"><b><font size="2" face="Verdana" color="#000000">Bil</font></b></td>
    <td width="86" align="center"><b><font size="2" face="Verdana" color="#000000">No 
      Kompaun</font></b></td>
    <td width="45" align="center"><b><font size="2" face="Verdana" color="#000000">Akta</font></b></td>
    <td width="44" align="center"><b><font size="2" face="Verdana" color="#000000">Jenis</font></b></td>
    <td width="204" align="center"><b><font size="2" face="Verdana" color="#000000">Nama 
      </font></b></td>
    <td width="57" align="center"><b><font size="2" face="Verdana" color="#000000">Tarikh&nbsp;</font></b></td>
    <td width="51" align="center"><b><font size="2" face="Verdana" color="#000000">Amaun</font></b></td>
    <td width="65" align="center"><b><font size="2" face="Verdana" color="#000000">Tkh 
      Bayar&nbsp;</font></b></td>
    <td width="48" align="center"><b><font size="2" face="Verdana" color="#000000">No 
      Resit</font></b></td>
  </tr>
  <tr> 
    <td align="center" colspan="9" bgcolor="<%=color1%>"> <hr> </td>
  </tr>
  <%
  		bil = 0
   		belum = 0
   			
    	Do while not objRs2.EOF
    	
    '	rekod = objRs2("rekod")
    '	sudah = objRs2("sudah")
    	
    '	belum = cdbl(rekod) - cdbl(sudah)
    	bil = bil + 1
%>
  <tr bgcolor="lightgrey"> 
    <td width="28" align="center"><font size="2" face="Verdana"><%=bil%></font></td>
  <td width="86" align="center"><font size="2" face="Verdana"><%=objRs2("no_akaun")%></font></td>
  <%	kara = objRs2("perkara")
  		if kara <> "" then	%>
    <td width="45" align="center" onMouseover="this.style.backgroundColor='#FFFFCC'" onMouseout="this.style.backgroundColor='#CCCCCC'"><a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"> 
      <font size="2" face="Verdana" color="Blue"><b><%=objRs2("perkara")%></b></font></a></td>
	<%	else  %>
  <td width="45" align="center"><font size="2" face="Verdana" color="Blue"><%=objRs2("perkara")%></font></td>
	<%	end if		%>
  <%	ruj2 = objRs2("no_rujukan2") 
  		if ruj2 <> "" then	%>	
    <td width="44" align="center" onMouseover="this.style.backgroundColor='#FFFFCC'" onMouseout="this.style.backgroundColor='#CCCCCC'"> 
      <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRs2("no_rujukan2")%>"> 
      <font size="2" face="Verdana"><%=objRs2("nama")%></font><font size="2" face="Verdana" color="Blue"></font></a></td>
	<%	else	%>
    <td width="44" align="center"><font size="2" face="Verdana"><%=objRs2("tkh_masuk")%></font></td>
	<%	end if		%>
    <td width="204" align="left"><font size="2" face="Verdana"><%=FormatNumber(objRs2("amaun"),2)%></font></td>  
    <td width="57" align="center"><font size="2" face="Verdana"><%=objRs2("tkh_bayar")%></font></td>  
    <td width="51" align="right"><font size="2" face="Verdana"><%=objRs2("no_resit")%>&nbsp;&nbsp;</font></td> 

  </tr>
  <%
  	objRs2.MoveNext			
  	Loop
%>
  <tr> 
    <td colspan="9"><font size="2" face="Trebuchet MS"> 
      <hr>
      </font></td>
  </tr>
</table>
<%end sub %>

</body>