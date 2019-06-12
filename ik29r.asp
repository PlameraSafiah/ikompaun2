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

<%Set objConn = Server.CreateObject("ADODB.Connection")
  objconn.Open "dsn=12c;uid=majlis;pwd=majlis;"
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
		
		
		pekz = request.cookies("gnop")
		'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		if not objRLokasi.eof then
		lok = objRLokasi("lokasi")
		end if 
		
				
		end if
		'end view ikut jabatan
		'******************************************************************
   		
		ff = " select stesyen, no_akaun,nvl(amaun_bayar,0) am,to_char(tarikh,'yyyy') yy,to_char(tarikh,'mm') mm ,to_char(tarikh,'dd') dd, no_resit "
 	    ff = ff & " from kutipan.kutipan where tarikh between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		ff = ff & " and to_date('"& tkh2 &"' , 'ddmmyyyy') and (status <> 'B' or status is null) and post is null"
		ff = ff & " and no_akaun <> '764102101353'"
		if lok = "112" then 'jabatan perlesenan
		ff = ff & " and no_akaun like '76410'||'%' "
		end if 
		
		if lok = "105" then  'bangunan 
		ff = ff & " and no_akaun like '76101'||'%' "
		end if
		
		if lok = "103" then  'perbandaran 
		ff = ff & " and (no_akaun like '76412'||'%' or no_akaun like '76415'||'%' or no_akaun like '76416'||'%' )"
		end if 
		
		if lok = "113" then  'kesihatan & veterinar 
		ff = ff & " and ( no_akaun like '76413'||'%' or no_akaun like '76420'||'%' ) "
		end if 
		
		ff = ff & " union "
		ff = ff & " select stesyen,no_akaun, nvl(amaun_bayar,0) am,to_char(tkh_bayar,'yyyy') yy,to_char(tkh_bayar,'mm') mm,"
		ff = ff & " to_char(tkh_bayar,'dd')dd,no_resit "
	    ff = ff & " from hasil.bil where tkh_bayar between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		ff = ff & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		ff = ff & " and jabatan = '"& lok &"' and no_resit is not null"
		
		if lok = "112" then 'jabatan perlesenan
		ff = ff & " and no_akaun like '76410'||'%' "
		end if 
		
		if lok = "105" then  'bangunan 
		ff = ff & " and no_akaun like '76101'||'%' "
		end if
		
		if lok = "103" then  'perbandaran 
		ff = ff & " and (no_akaun like '76412'||'%' or no_akaun like '76415'||'%' or no_akaun like '76416'||'%' )"
		end if 
		
		if lok = "113" then  'kesihatan & veterinar 
		ff = ff & " and ( no_akaun like '76413'||'%' or no_akaun like '76420'||'%' ) "
		end if 
		
		ff = ff & " union "
		
		ff = ff & " select 'MPSPPAY' stesyen,  no_akaun,nvl(amaun,0) am,to_char(tarikh,'yyyy') yy,to_char(tarikh,'mm') mm,"
 	    ff = ff & " to_char(tarikh,'dd')dd,no_resit "
 	    ff = ff & " from hasil.ebayar_trxid where tarikh between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		ff = ff & " and to_date('"& tkh2 &"' , 'ddmmyyyy') and flag = 'SUCCESSFUL' and status_kutipan is null "
		if lok = "112" then 'jabatan perlesenan
		ff = ff & " and no_akaun like '76410'||'%' "
		end if 
		
		if lok = "105" then  'bangunan 
		ff = ff & " and no_akaun like '76101'||'%' "
		end if
		
		if lok = "103" then  'perbandaran 
		ff = ff & " and (no_akaun like '76412'||'%' or no_akaun like '76415'||'%' or no_akaun like '76416'||'%' )"
		end if 
		
		if lok = "113" then  'kesihatan & veterinar 
		ff = ff & " and ( no_akaun like '76413'||'%' or no_akaun like '76420'||'%' ) "
		end if 
		
		ff = ff & " order by yy asc,mm asc,dd asc "
		''response.Write(ff)
		Set objRs2 = objConn.Execute(ff)
		
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
    <td width="30%" align="left" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Nama</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Tarikh&nbsp;</font></b></td>
    <td width="11%" align="left" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Tkh 
      Bayar&nbsp;</font></b></td>
    <td width="9%" align="left" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">No 
      Resit</font></b></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><b><font size="2" face="Verdana" color="#000000">Amaun</font></b></td>
  </tr>
  <%	
	bil = 0
	ctrz = 0
	
	do while not objRs2.eof
	bil = bil + 1
	ctrz = cdbl(ctrz) + 1
			
		noakaun = objRs2("no_akaun")
		stesyen = objRs2("stesyen")
	 	amq = objRs2("am")
		yy = objRs2("yy")
		mm = objRs2("mm")
		dd = objRs2("dd")		
		tkhq = cstr(dd)+"/"+cstr(mm)+"/"+cstr(yy)
		resitq = objRs2("no_resit")
		
			
			
		d = " select no_akaun,perkara, no_rujukan2, nama, to_char(tkh_masuk,'ddmmyyyy') tkh_masuk,jabatan, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'ddmmyyyy') tkh_bayar, no_resit "
		d = d & " from hasil.bil "
		'd = d & " where (perkara <> 'P01' or perkara is null) "
		d = d & " where no_akaun = '"& noakaun &"' "
		
		set objRss = objconn.execute(d)
		'response.Write "<br><br>"&(d)
		if not objRss.eof then
		abayar = objRss("amaun_bayar")
		perkara = objRss("perkara")
		no_rujukan2 = objRss("no_rujukan2")
		nama = objRss("nama")
		tkh_masuk = objRss("tkh_masuk")
		jabatan = objRss("jabatan")
		tkh_bayar = objRss("tkh_bayar") 
		no_resit = objRss("no_resit")
		
				else
		perkara = ""
		no_rujukan2 = ""
		nama = ""
		tkh_masuk = ""
		jabatan = ""
		tkh_bayar = ""
		no_resit = ""

		end if 
		
		'----------jumlah keseluruhan	
			jamaun = cdbl(jamaun) + cdbl(abayar)  
			
	   %>
            
            
  <tr bgcolor="#FFFFFF" > 
    <td align="center"><font color="#000000" size="2" face="Trebuchet MS">&nbsp;<%=bil%></font></td>
    <td width="10%" align="center"><font color="#000000" size="2" face="Verdana"><%=noakaun%></font></td>
    <%	kara = perkara
  		if kara <> "" then
  %>
    <td width="7%" align="center" > <font size="2" face="Verdana" color="#000000"><b><%=perkara%></b></font></td>
    <%	else	%>
    <td width="7%" align="center"> <font size="2" face="Verdana" color="#000000"><b><%=perkara%></b></font></td>
    <%	end if		
  		rujuk2 = no_rujukan2
  		if rujuk2 <> "" then
  %>
    <td width="10%" align="center" > 
      <font size="2" face="Verdana" color="#000000"><b><%=no_rujukan2%></b></font></td>
    <%	else		%>
    <td width="10%" align="center">
      <font size="2" face="Verdana" color="#000000"><b><%=no_rujukan2%></b></font></td>
    <%	end if		%>
    <td width="28%" align="center"><font size="2" face="Verdana">
      <p align="left"><font color="#000000"><%=nama%></font></font></td>
    <td width="11%" align="center"><font color="#000000" size="2" face="Verdana"><%=tkh_masuk%> 
      </font></td>
    <% if no_resit <> "" then%>
  <td><%=tkh_bayar%></td>
  <td><%=no_resit%></td>
  <td align="right" ><%=FormatNumber(abayar,2)%></td>
  <% else %>
   <td colspan="3"><font color="#FF0066" size="-2">Bayaran Telah Dibuat Pada <%=tkhq%>
	  di Stesyen <%=stesyen%>. <br>Jumlah: RM <%=amq%>. No Resit:<%=resitq%> . Data Belum DiSahkan </font></td>

  <% end if %>
  <%	objRs2.MoveNext			
  	Loop
%>

<tr>
<td colspan="9"><hr></td>
</tr>

  <tr bgcolor="#FFFFFF" > 
    <td colspan="2" align="center">&nbsp;</td>
    <td colspan="6" align="right"><b><font size="2" face="Trebuchet MS">&nbsp; 
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