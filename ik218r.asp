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
		
		d = " select no_akaun, perkara, no_rujukan2 ,initcap(nama) nama, to_char(tkh_masuk,'dd/mm/yyyy') tkh_masuk, jabatan, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'dd/mm/yyyy') tkh_bayar, no_resit, "
		d = d & " nvl(amaun_bayar,0) amaun_bayart from hasil.bil "
		d = d & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		''d = d & " and (no_akaun like '"& kod &"'||'%' ) "
		'd = d & " and substr(no_akaun,5,2) = '"& dae &"' "
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
      KOMPAUN KESELURUHAN DARI TARIKH<br>
      <%=tkh1%>&nbsp;HINGGA&nbsp;<%=tkh2%></font></font></b></td>
    </tr>
  </table>
  <br>

<hr>

<table border="0" width="100%">
<tr align="center"> 
<td width="19"  class="hd1"><b><font size="2" face="Verdana" color="#000000">Bil</font></b></td>
<td width="95"  class="hd1"><b><font size="2" face="Verdana" color="#000000">No Kompaun</font></b></td>
<td width="34"  class="hd1"><b><font size="2" face="Verdana" color="#000000">Akta</font></b></td>
<td  class="hd1"><b><font size="2" face="Verdana" color="#000000">Jenis</font></b></td>
<td width="30"  class="hd1"><b><font size="2" face="Verdana" color="#000000">Nama</font></b></td>
<td width="98"  class="hd1"><b><font size="2" face="Verdana" color="#000000">Tarikh Masuk</font></b></td>
<td width="70"  class="hd1" ><b><font size="2" face="Verdana" color="#000000">Status Bayaran</font></b></td>
<td width="50"  class="hd1"><b><font size="2" face="Verdana" color="#000000">Tarikh Bayaran</font></b></td>
<td width="20" class="hd1" ><b><font size="2" face="Verdana" color="#000000">Amaun Bayar</font></b></td>
<td class="hd1"><b><font size="2" face="Verdana" color="#000000">No Resit</font></b></td>
<td width="250" class="hd1"><b><font size="2" face="Verdana" color="#000000">Jabatan </font></b></td>
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


  <tr> 
  <td><font size="2" face="Verdana" color="#000000"><%=bil%></font></td>
  <td><font size="2" face="Verdana" color="#000000"><%=objRs2("no_akaun")%></font></td>
  <%	kara = objRs2("perkara")
  		if kara <>"" then		%>
  <td><font size="2" face="Verdana" color="#000000">
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></a></font></td>
	<%	else	%>
  <td width="39" align="center" ><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></td>
  <%	end if		%>
  <%	ruj2 = objRs2("no_rujukan2")
  		if ruj2 <>"" then		%>
  <td> 
  <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRS2("no_rujukan2")%>">
  <font size="2" color="Blue"><b><%=objRs2("no_rujukan2")%></b></font></a></td>
	<%	else	%>
 <td><font size="2" face="Verdana" color="#000000"><%=objRs2("no_rujukan2")%></font></td>
	<%	end if		%>
  <td><font size="2" face="Verdana" color="#000000"><%=objRs2("nama")%></font></td>
  <td align="center"><font size="2" face="Verdana" color="#000000"><%=objRs2("tkh_masuk")%></font></td>
  
  <td align="center"><font size="2" face="Verdana" color="#000000">
   <% tarikhbayar = objRs2("tkh_bayar")
   if tarikhbayar <> "" then %> Y <%else%> T <% end if %>
	</font></td>
  <td width="42" align="center"><font size="2" face="Verdana" color="#000000"><%=objRs2("tkh_bayar")%></font></td>
  <td width="" align="center"><font size="2" face="Verdana" color="#000000"><%=FormatNumber(objRs2("amaun_bayar"),2)%></font></td>
  <td width="18" align="center"><font size="2" face="Verdana" color="#000000"><%=objRs2("no_resit")%></font></td>
   <td ><font size="2" face="Verdana" color="#000000">&nbsp;&nbsp;&nbsp;<%=namaJabatan%></font></td>
  </tr>




  <%	objRs2.MoveNext			
  	Loop
%>
  <tr bgcolor="#FFFFFF" > 
    <td colspan="2" align="center">&nbsp;</td>
    <td colspan="6" align="right"><b><font size="2" face="Trebuchet MS">&nbsp; 
     <b> Jumlah :&nbsp;&nbsp;</font></b></td>
    <td align="center" ><b><font size="2" face="Trebuchet MS"><%=formatnumber(jamaun,2)%>&nbsp;&nbsp;&nbsp;</font></b></td>
  </tr>
</table>
<%
end if
end if%>
<hr>
</body>

</html>