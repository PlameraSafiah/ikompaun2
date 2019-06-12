<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Ringkasan Mengikut Pegawai</title>
</head>

<body topmargin="0" leftmargin="0">
<%
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
	
	proses = Request.form("B2")
	nopek = Request.form("fnopek")
	tkh1 = Request.form("ftkh1")
	tkh2 = Request.form("ftkh2")
	
	
	
		d = " select perkara,jabatan, no_rujukan2, count(1) bilsalah,count(amaun_bayar) bilbayar from hasil.bil "
		d = d & " where no_pekerja = '"& nopek &"' "
		d = d & " and tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%') "
		'18042014 : jun tambah boleh papar kod 76413 jugak, sebelum ni 76410 & 76420 saja
		d = d & " and perkara <> 'P01'  "
		d = d & " group by perkara, no_rujukan2,jabatan "
		d = d & " order by perkara, no_rujukan2 "
		Set objRs2 = Server.CreateObject ("ADODB.Recordset")
		Set objRs2 = objConn.Execute(d)
		jab = objRs2("jabatan")
	
	
	c = " select initcap(keterangan) terang from payroll.ptj where kod = '"&jab&"' "
	Set objRsc = Server.CreateObject("ADODB.Recordset")
	Set objRsc = objConn.Execute(c)
	
	if not objRsc.EOF then
		terang = objRsc("terang")
	else
		terang = ""
	end if	


		f = " select nama from payroll.paymas where no_pekerja = '"& nopek &"' "
	 	f = f & " union "
	 	f = f & " select nama from payroll.paymas_sambilan where no_pekerja = '"& nopek &"' "
	 	Set objRs3 = objConn.Execute(f)
	 
	 
	 	if not objRs3.eof then
	 	 	napek = objRs3("nama")
	 	else napek = ""
	 	end if	
	
			
		f="select to_char(sysdate,'dd-mm-yyyy  hh24:mi:ss') as tkhs from dual "
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
<td width="20%" align="right" ><font size="2" color="#COCOCO">Muka :&nbsp;<%=muka%></font></td>
</tr>
</table>

<table width="100%" border="0">
<tr>
<td width="100%" align="center"><font face="MS Serif" size="4"><b><%=namas%></font></b></td>
</tr>
<tr>
<td width="100%" align="center"><font face="MS Serif" size="3"><b>LAPORAN RINGKASAN KOMPAUN MENGIKUT PEGAWAI</b></font></td>
</tr></table>

<p></p>
<table width="100%" border="0" align="center">
<tr>
<td><font face="Verdana" size="2"><b>No Pekerja :&nbsp;</b><%=nopek%> - <%=napek%></font></td>
</tr>
</table>

<table width="100%" border="0" align="center">
<tr>
<td><font face="Verdana" size="2"><b>Jabatan &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;</b><%=objRs2("jabatan")%> - <%=terang%></font></td>
</tr>
</table>

<table width="100%" border="0" align="center">
<tr>
<td><font face="Verdana" size="2"><b>Tarikh Dari :</b> <%=tkh1%>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <b>Hingga :</b> <%=tkh2%></font></td>
</tr>
</table>

<hr>
<p></p>
<%	end sub	%>



<%	sub badan		%>

<table width="99%" border=1 borderColor=#000000 cellPadding=1 cellSpacing=0 rules=all align="center"
style="border-collapse: collapse; border: 1px solid black">
<tr bgcolor="#FFFFFF"> 
<td width="5%" align="center"><b><font color="#000000" size="2" face="Arial">Bil</font></b></td>
<td width="8%" align="center"><b><font color="#000000" size="2" face="Arial">Akta</font></b></td>
<td width="50%" align="center"><b><font color="#000000" size="2" face="Arial">Kod
  / Keterangan</font></b></td>
<td width="12%" align="center"><b><font color="#000000" size="2" face="Arial">Bil Kompaun<br>Belum Bayar</font></b></td>
<td width="12%" align="center"><b><font color="#000000" size="2" face="Arial">Bil Kompaun<br>Sudah Bayar</font></b></td>
<td width="12%" align="center"><b><font color="#000000" size="2" face="Arial">Bil Kompaun</font></b></td>
</tr>
 <%		bil = 0
 		ctr = 0
  		ctrz = 0

   		Do while not objRs2.EOF	
   		
   		perkara = objRs2("perkara")
   		no_rujukan2 = objRs2("no_rujukan2")
   		k = "     select upper(keterangan) keter from kompaun.jenis_kesalahan "
   		k = k & " where perkara = '"& perkara &"' and kod = '"& no_rujukan2 &"' "
   		set ok = objConn.Execute(k)
   			if not ok.eof then
				keter = ok("keter")   			
   			end if
   			
   		bilsalah = objRs2("bilsalah")
   		bilkpn = cdbl(bilkpn) + cdbl(bilsalah)
   		bilbayar = objRs2("bilbayar")
		
   		bilbelum = cdbl(bilsalah) - cdbl(bilbayar)
   		jbilbayar = cdbl(jbilbayar) + cdbl(bilbayar)
   		jbilbelum = cdbl(jbilbelum) + cdbl(bilbelum)
   		bil = bil + 1
   		ctr = cdbl(ctr) + 1
    	ctrz = cdbl(ctrz) + 1
    	if ctr = 39 then
    		ctr = 1  	
%> 
</table>

<%mula%> 
<table width="99%" border=1 borderColor=#000000 cellPadding=1 cellSpacing=0 rules=all align="center"
style="border-collapse: collapse; border: 1px solid black" >
<tr bgcolor="#FFFFFF"> 
<td width="5%" align="center" ><b><font color="#000000" face="Arial" size="2">Bil</font></b></td>
<td width="8%" align="center" ><b><font color="#000000" face="Arial" size="2">Akta</font></b></td>
<td width="50%" align="center" ><b><font color="#000000" face="Arial" size="2">Kod / Keterangan</font></b></td>
<td width="12%" align="center" ><b><font color="#000000" face="Arial" size="2">Bil Kompaun<br>Belum Bayar</font></b></td>
<td width="12%" align="center" ><b><font color="#000000" face="Arial" size="2">Bil Kompaun<br>Sudah Bayar</font></b></td>
<td width="12%" align="center" ><b><font color="#000000" face="Arial" size="2">Bil Kompaun</font></b></td>
</tr>

<%	end if	%>

<tr bgcolor="#FFFFFF"> 
<td width="5%" align="center" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=bil%></font>&nbsp;</td>
<td width="8%" align="center" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=perkara%></font>&nbsp;</td>
<td width="50%" align="left" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=no_rujukan2%>&nbsp;<%=keter%></font></td>
<td width="12%" align="center" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=bilbelum%></font>&nbsp;</td>
<td width="12%" align="center" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=bilbayar%></font>&nbsp;</td>
<td width="12%" align="center" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=bilsalah%></font>&nbsp;</td>
</tr>
 <%
   		objRs2.MoveNext
   		Loop
 %> 
<tr> 
<td colspan="3" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">JUMLAH</font></td>
<td align="center" width="12%" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=jbilbelum%></font>&nbsp;</td>
<td align="center" width="12%" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=jbilbayar%></font>&nbsp;</td>
<td align="center" width="12%" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=bilkpn%></font>&nbsp;</td>
</tr>
</table>

<%	end sub	 %>

</body>

</html>