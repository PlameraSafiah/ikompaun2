<!-- '#INCLUDE FILE="ik.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Kompaun Mengikut Daerah</title>
</head>

<body onload='self.print()' topmargin="0" leftmargin="0">
<%	'Set objConn = Server.CreateObject("ADODB.Connection")
'	objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
	
  proses = Request.form("B2")
  dae = Request.form("fdae")
  kod = Request.form("fkod")
  tkh1 = Request.form("ftkh1")
  tkh2= Request.form("ftkh2")
tkh1a = cstr(mid(tkh1,1,2)) + "/" + cstr(mid(tkh1,3,2)) + "/" + cstr(mid(tkh1,5,4))
  tkh2a = cstr(mid(tkh2,1,2)) + "/" + cstr(mid(tkh2,3,2)) + "/" + cstr(mid(tkh2,5,4))
	
	if proses = "Cetak" then
	    
		
		d = " select no_akaun, perkara, no_rujukan2 ,initcap(nama) nama, to_char(tkh_masuk,'ddmmyyyy') tkh_masuk, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'ddmmyyyy') tkh_bayar, no_resit, jabatan "
		d = d & " from hasil.bil "
		d = d & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran-mimi-pn.raja
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%'or no_akaun like '76441'||'%') "
		d = d & " and substr(no_akaun,5,2) = '"& dae &"' "
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
		d = d & " order by no_rujukan "
		Set objRs2 = objConn.Execute(d)
		

	end if
	
	
		
	
	    if dae = "01" then nama = "Seberai Perai Utara" 
	    if dae = "02" then nama = "Seberai Perai Tengah" 
	    if dae = "03" then nama = "Seberai Perai Selatan" 
	
	f="select to_char(sysdate,'dd-mm-yyyy  hh24:mi:ss') as tkhs from dual "
   	Set objRs1a = Server.CreateObject ("ADODB.Recordset")
   	Set objRs1a = objConn.Execute(f)	
   	tkhs = objrs1a("tkhs")
   	
   	
	s = " select nama from majlis.syarikat "     	
	Set objRss = objConn.Execute(s)
	namas = objRss("nama")
       
     
	 			

		
		
    muka = 0
    mula
    badan     
%>

<%	sub mula	
		muka = cdbl(muka) + 1
%>
<table width="100%" border="0" >
<tr>
<td width="20%" align="left" ><i><font size="2" color="#COCOCO"><%=tkhs%></font></i></td>
<td width="60%" align="center"></td>
<td width="20%" align="right" ><font size="2" color="#COCOCO">Page <%=muka%></font></td>
</tr>
</table>

<table width="100%" border="0">
<tr>
    <td width="100%" align="center"><font face="MS Serif" size="4"><b><%=namas%></font></td>
</tr>
<tr>
    <td width="100%" align="center"><font face="MS Serif" size="3"><b>LAPORAN 
      KOMPAUN MENGIKUT DAERAH</b></font></td>
</tr></table>
<%'response.write "nopek"&nopek&""%>
<p></p>
<table width="85%" align="center" border="0">
  <tr> 
    <td width="14%" nowrap><font size="2" face="Verdana"><b> Tarikh</b></font></td>
    <td width="1%" ><strong><font size="2" face="Arial"> :</font></strong></td>
    <td width="85%" ><font size="2" face="Verdana"><%=tkh1a%><b> Hingga</b> <%=tkh2a%></font></td>
  </tr>
  <tr> 
    <td nowrap><font size="2" face="Verdana"><b> Daerah</b></font></td>
    <td ><strong><font size="2" face="Arial">:</font></strong></td>
    <td ><font size="2" face="Verdana"><%=dae%>-<%=nama%></font></td>
  </tr>
</table>
<hr>
<%	end sub	%>



<%	sub badan		%>

<center>
<table border="0" width="100%" align="center">
  <tr> 
    <td align="center"><b><font size="2" face="Arial">Bil</font></b></td>
    <td align="center"><b><font size="2" face="Arial">No Kompaun</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Akta</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Jenis</font></b></td>
    <td width="40%" align="center"><b><font size="2" face="Arial">Nama</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Tkh Kompaun</font></b></td>
    <td align="center"><strong><font size="2" face="Arial">Amaun</font></strong></td>
    <td align="center"><b><font size="2" face="Arial">Tkh Bayar</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Resit</font></b></td>
   <td align="center"><b><font size="2" face="Arial">Jabatan</font></b></td>
  </tr>
  <%	ctr = 0
  	ctrz = 0
	bil = 0

	Do while not objRs2.eof
	
	tb=objrs2("tkh_bayar")
	name=objrs2("nama")
	nama=mid(name,1,30)
		 if tb="" then
		  tkh="-"
		  else
		  tkh=objrs2("tkh_bayar")
		  end if
	ab = objRs2("amaun_bayar")
		total_ab = cdbl(total_ab) + cdbl(ab)
	bil = bil + 1
	ctr = cdbl(ctr) + 1
    	ctrz = cdbl(ctrz) + 1
    	if ctr = 30 then
    		ctr = 1  
			
			

				
  %>
</table>

<%mula%>

<center>
<table width="100%" height="46" border="0" align="center">
  <tr> 
    <td align="center"><b><font size="2" face="Arial">Bil</font></b></td>
    <td align="center"><b><font size="2" face="Arial">No Kompaun</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Akta</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Jenis</font></b></td>
    <td width="40%" align="center"><b><font size="2" face="Arial">Nama</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Tkh Kompaun</font></b></td>
    <td align="center"><strong><font size="2" face="Arial">Amaun</font></strong></td>
    <td align="center"><b><font size="2" face="Arial">Tkh Bayar</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Resit</font></b></td>
    <td align="center"><b><font size="2" face="Arial">Jabatan</font></b></td>
  </tr>
  <%	end if
  
  		jab1 = objRs2("jabatan")
				
		jab = " select kod, keterangan from majlis.jabatan "
		jab = jab & " where kod = '"& jab1 &"' "
		
		set objJab = ObjConn.Execute(jab)
		
		if objJab.eof then
		
		end if
		
			%>
  <tr> 
    <td height="20" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=bil%></font></td>
    <td height="20" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=objRs2("no_akaun")%></font></td>
    <td align="left" nowrap><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=objRs2("perkara")%></font></td>
    <td align="left" nowrap><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<%=objRs2("no_rujukan2")%></font></td>
    <td width="40%" align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=name%></font></td>
    <td height="20" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;<%=objRs2("tkh_masuk")%></font></td>
    <td align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=formatnumber(objRs2("amaun_bayar"),2)%></font></td>
    <td height="20" align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=tkh%>&nbsp;&nbsp;</font></td>
    <td height="20" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=objRs2("no_resit")%></font></td>
    <td height="20" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=objJab("keterangan")%></font></td>
  </tr>
  <%	objRs2.MoveNext
	Loop
%>
 <tr>
 <td colspan="6" align="right"><b>Jumlah</b></td>
 <td align="right"><b><%=FormatNumber(total_ab,2)%></b></td></tr>
</table>
<%	end sub	%>

</body>
</html>

