<%response.cookies("ikmenu") = "ik218.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Kompaun Keseluruhan</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "tkhdari";
netscape="";
ver=navigator.appVersion; len=ver.length;
for(iln=0;iln<len;iln++)if(ver.charAt(iln)=="(")break;
netscape=(ver.charAt(iln+1).toUpperCase()!="C");

function keyDown(DnEvents){
k=(netscape)?DnEvents.which:window.event.keyCode;
if(k==13){
if(nextfield=='done')return true;
else{
eval('document.myform.'+nextfield+'.focus()');
return false;
		}
	}
}
document.onkeydown=keyDown;
if(netscape)document.captureEvents(Event.KEYDOWN|Event.KEYUP);
//End -->
</script>
<script language="javascript">
   function invalid_data(a)
  {  
       alert (a+" Tiada Rekod ");
		return(true);
  }
     function invalid_tarikh(b)
  {  
       alert (b+" Tarikh Salah ");
		return(true);
  }
  
</script>
</head>
<form name=myform method="POST" action="ik218.asp">



<%
	'Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   	
   	proses = Request.form("B1")
   		
	if proses <> "Cari" then
		
		e = " select '01'||to_char(sysdate,'mm')||to_char(sysdate,'yyyy') as tkh1s, to_char(sysdate,'ddmmyyyy') as tkh2s from dual "
		Set objRse = Server.CreateObject ("ADODB.Recordset")
   		Set objRse = objConn.Execute(e)	
   		tkh1 = objRse("tkh1s")	
   		tkh2 = objRse("tkh2s")
   		
  	end if

	
	if proses = "Cari" then
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")
	end if



	dtkh1 = Request.QueryString("dtkh1")

	if dtkh1 <> "" then
		tkh1 = Request.QueryString("dtkh1")
		tkh2 = Request.QueryString("dtkh2")
	end if


%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
    <td class="hd">Tarikh Dari</td><td> &nbsp;&nbsp;<input type="text" name="tkhdari" value="<%=tkh1%>" size="8" maxlength="8" onFocus="nextfield='tkhhingga';">&nbsp;&nbsp;
  &nbsp;&nbsp;&nbsp;&nbsp;Hingga&nbsp; <input type="text" name="tkhhingga" value="<%=tkh2%>" size="8" maxlength="8" onFocus="nextfield='done';">&nbsp;
  <input type="submit" value="Cari" name="B1" class="button"></td>
</tr>
<script>
	document.myform.tkhdari.focus();
</script>
</table>   
<%	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 50 '----------tukar jd 30 kejap..utk test jumlah-------

	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If



	if proses = "Cari" or dtkh1 <>"" then
		
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
   		
		d = " select no_akaun, perkara, no_rujukan2 ,initcap(nama) nama, to_char(tkh_masuk,'dd/mm/yyyy') tkh_masuk, jabatan,  "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'dd/mm/yyyy') tkh_bayar, no_resit, "
		d = d & " nvl(amaun_bayar,0) amaun_bayart from hasil.bil "
		d = d & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		''d = d & " and (no_akaun like '"& kod &"'||'%' ) "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%') "

		'd = d & " and substr(no_akaun,5,2) = '"& dae &"' "
		d = d & " and (perkara <> 'P01' or perkara is null) "
		'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		d = d & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
		d = d & " order by  to_date(tkh_masuk,'dd/mm/yyyy') "
		

		
		
	''response.Write(d)
		Set objRs2 = Server.CreateObject ("ADODB.Recordset")
		objRs2.PageSize = iPageSize
		objRs2.CacheSize = iPageSize
 		
		objRs2.CursorLocation = 3
		objRs2.Open d, objConn
		iPageCount = objRs2.PageCount 
		
		
		if not objRs2.bof and not objRs2.eof then
		kira=objRs2.recordcount
		rekod="ada"
		If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
		If iPageCurrent < 1 Then iPageCurrent = 1

		bil=0
		bilangan=Request.QueryString("bilangan")
		ms=Request.QueryString("ms")
		
		If bilangan <>"" and ms="next" then
			bil = bilangan
		End If
		If bilangan <>"" and ms="pre" then
			bil = bilangan
		End If
		
		If iPageCount <> 0 Then
			objRs2.AbsolutePage = iPageCurrent
   			iRecordsShown = 0
			count = 0
		Do While iRecordsShown <iPageSize And Not objRs2.eof 
			iRecordsShown = iRecordsShown + 1
			count = count + 1
			bil=bil + 1
		objRs2.movenext
		loop
		end if
		end if

'-----------------kira total keseluruhan------------

		dd = " select sum(nvl(amaun_bayar,0)) amaun_bayart  "
		dd = dd & " from hasil.bil "
		dd = dd & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
	
		dd = dd & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
			dd = dd & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		dd = dd & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		dd = dd & " or no_akaun like '76101'||'%') "

		dd = dd & " and (perkara <> 'P01' or perkara is null) "
		'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		admind = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmind = objConn.Execute(admind)
		
		if objRAdmind.eof then
		
		lokasid = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasid = objConn.Execute(lokasid)
		
		lok = objRLokasid("lokasi")
		
		dd = dd & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
		Set dd2 = objConn.Execute(dd)
'----------------end kira total keseluruhan-------------
'-----------------kira kompaun dh bayau-----------------
	
		d2 = " select count(*) dh_bayar  "
		d2 = d2 & " from hasil.bil "
		d2 = d2 & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d2 = d2 & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d2 = d2 & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d2 = d2 & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d2 = d2 & " or no_akaun like '76101'||'%') "

		d2 = d2 & " and (perkara <> 'P01' or perkara is null) "
		d2 = d2 & " and tkh_bayar is not null "
		response.Write(d2)
		'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		admin2 = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin2 = objConn.Execute(admin2)
		
		if objRAdmin2.eof then
		
		lokasi2 = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi2 = objConn.Execute(lokasi2)
		
		lok = objRLokasi2("lokasi")
		
		d2 = d2 & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
		Set dsv = objConn.Execute(d2)
'-----------------end kira kompaun dh bayau---------------				
			
		if objRs2.bof and objRs2.eof then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
       	response.write "</script>"
		else

		if kira > 0 then
%>  <br/>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
  <tr align="right">
  	<td align="left" width="25%">Jumlah Rekod : <%=kira%></td>  
    </tr>
    <tr align="right">
  	<td align="left" width="25%">Bilangan Kompaun Dibayar : <%=dsv("dh_bayar")%></td>  
    </tr>
    <tr align="right">
  	<td align="left" width="25%">Jumlah Kompaun Dibayar : <b><%=FormatNumber(dd2("amaun_bayart"),2)%></b></td>  
    </tr>
    
    
    <tr align="right">
    
    <td width="75%" >
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik218.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik218.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik218.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik218.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %>	
	</td>
  </tr>
  </table>

 <table width="100%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="29" class="hd1">Bil</td>
<td width="74" class="hd1">No Kompaun</td>
<td width="41" class="hd1">Akta</td>
<td width="43" class="hd1">Jenis</td>
<td width="51" class="hd1">Nama</td>
<td width="91" class="hd1">Tarikh Masuk</td>
<td class="hd1" colspan="2">Status Bayaran</td>
<td width="61" class="hd1">Tarikh Bayaran</td>
<td width="55" class="hd1" align="right">Amaun Bayar</td>
<td width="50" class="hd1">No Resit</td>
<td width="63" class="hd1">Jabatan </td>
  </tr>
<tr align="center"> 
<td class="hd1" colspan="6">&nbsp;</td>
<td width="61" class="hd1" align="center" >Ya</td>
<td width="65" class="hd1" align="center">Tidak</td>
<td class="hd1" colspan="4">&nbsp;</td>

  </tr>
  
<%		bil = 0
		ctrz = 0
		ab = 0
		total_ab = 0
		total_abs = 0
	
		bilangan=Request.QueryString("bilangan")
		page = Request.QueryString("page")
		ms=Request.QueryString("ms")

		If bilangan <>"" and ms="next" then
			bil = bilangan
		End If
		If bilangan <>"" and ms="pre" then
			bil = bilangan
		End If
		If iPageCount <> 0 Then
			objRs2.AbsolutePage = iPageCurrent
   			iRecordsShown = 0
			count = 0
			
		Do While iRecordsShown <iPageSize And Not objRs2.eof 


		bil = bil + 1
		ctrz = cdbl(ctrz) + 1
		
		
		ab = objRs2("amaun_bayar")
		total_ab = cdbl(total_ab) + cdbl(ab)
		
		kodJbtn = objRs2("jabatan")
		q1="select keterangan from payroll.ptj where kod='"&kodJbtn&"'"
		set rq1 = objConn.execute(q1)
				
		if not rq1.eof then namaJabatan = rq1("keterangan")
		

%>
  <tr> 
  <td><%=bil%></td>
  <td><%=objRs2("no_akaun")%></td>
  <%	kara = objRs2("perkara")
  		if kara <>"" then		%>
  <td>
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></a></td>
	<%	else	%>
  <td width="43" align="center" ><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></td>
  <%	end if		%>
  <%	ruj2 = objRs2("no_rujukan2")
  		if ruj2 <>"" then		%>
  <td> 
  <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRS2("no_rujukan2")%>">
  <font size="2" color="Blue"><b><%=objRs2("no_rujukan2")%></b></font></a></td>
	<%	else	%>
 <td><%=objRs2("no_rujukan2")%></td>
	<%	end if		%>

  <td><%=objRs2("nama")%></td>
  <td align="center"><%=objRs2("tkh_masuk")%></td>
    <% '----------------------------------------
	noa=objRs2("no_akaun")
	
	ff = " select stesyen, nvl(amaun_bayar,0) am,to_char(tarikh,'yyyy') yy,to_char(tarikh,'mm') mm ,to_char(tarikh,'dd') dd, no_resit "
    ff = ff & " from kutipan.kutipan where no_akaun = '"& noa &"' and (status <> 'B' or status is null) and post is null "
'	ff = ff & " union "
'	ff = ff & " select stesyen, nvl(amaun_bayar,0) am,to_char(tkh_bayar,'yyyy') yy,to_char(tkh_bayar,'mm') mm,"
'    ff = ff & " to_char(tkh_bayar,'dd')dd,no_resit "
'    ff = ff & " from hasil.bil2 where no_akaun = '"& no &"' and status is null "
	ff = ff & " union "
	ff = ff & " select 'MPSPPAY' stesyen, nvl(amaun,0) am,to_char(tarikh,'yyyy') yy,to_char(tarikh,'mm') mm,"
    ff = ff & " to_char(tarikh,'dd')dd,no_resit "
    ff = ff & " from hasil.ebayar_trxid where no_akaun = '"& noa &"' and flag = 'SUCCESSFUL' and status_kutipan is null "
	'response.write(ff)
	Set objRsf = objConn.Execute(ff)
	
	if not objRsf.eof then
	
		stesyen = objRsf("stesyen")
	 	amq = objRsf("am")
		yy = objRsf("yy")
		mm = objRsf("mm")
		dd = objRsf("dd")		
		tkhq = cstr(dd)+"/"+cstr(mm)+"/"+cstr(yy)
		resitq = objRsf("no_resit")

  %>
    
    <td colspan="5" align="center"><font color="#FF0066" size="-2">Bayaran Telah Dibuat Pada <%=tkhq%>
	  di Stesyen <%=stesyen%>. Jumlah: RM <%=amq%>. No Resit:<%=resitq%> . Data Belum DiSahkan </font></td>
  <%else%>
  
    <% tarikhbayar = objRs2("tkh_bayar")
  
  if tarikhbayar <> "" then %>
  
  <td align="center">Y</td>
	<td align="center">&nbsp;</td>
  <% else %>
<td align="center">&nbsp;</td>
  <td width="63" align="center">T</td>
  <% end if %>
  
  <td width="80" align="center"><%=objRs2("tkh_bayar")%></td>
  <td width="17" align="right"><%=FormatNumber(objRs2("amaun_bayar"),2)%></td>
  <td width="45" align="center"><%=objRs2("no_resit")%></td>
  
  <% end if %> 
<%'----------------------------------------------%>


  <td width="133"><%=namaJabatan%></td>
  </tr>
<%	iRecordsShown = iRecordsShown + 1
	count = count + 1	
  	objRs2.MoveNext			
  	Loop
		
%> 
 <tr>
 <td colspan="9" align="right"><b>Jumlah</b></td>
 <td align="right"><b><%=FormatNumber(total_ab,2)%></b></td></tr>
  <tr>
 <td colspan="9" align="right"><b>Jumlah Keseluruhan</b></td>
 <td align="right"><b><%=FormatNumber(dd2("amaun_bayart"),2)%></b></td></tr>

 </table>
<%  	end if		
		end if
		end if
		end if
  		end if  
  		end if
  		
  		'end if
%>
</td>
</tr>
</table>  
</form>
<form method="post" action="ik218r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
<input type="submit" value="Cetak" name="B2" class="button">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">
<input type="hidden" name="fkod" value="<%=kod%>">
</td></tr>
</table>
</form>
</body>
</html>

