<%response.cookies("ikmenu") = "ik30.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Kompaun Sudah Bayar</title>
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
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top"> 

<td width="100%">
<form name=myform method="POST" action="ik30.asp">
<%
	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   	
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
    <td class="hd">Tarikh Bayar</td><td> Dari&nbsp;&nbsp;<input type="text" name="tkhdari" value="<%=tkh1%>" size="8" maxlength="8" onFocus="nextfield='tkhhingga';">&nbsp;&nbsp;
  &nbsp;&nbsp;&nbsp;&nbsp;Hingga&nbsp; <input type="text" name="tkhhingga" value="<%=tkh2%>" size="8" maxlength="8" onFocus="nextfield='done';">&nbsp;
  <input type="submit" value="Cari" name="B1" class="button"></td>
</tr>
<script>
	document.myform.tkhdari.focus();
</script>
</table>   
<%	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 15

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
		d = d & " order by tkh_bayar "
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

			
			
		if objRs2.bof and objRs2.eof then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
       	response.write "</script>"
		else

		if kira > 0 then
%>  <br/>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="right">
  	<td align="left" colspan=2 >Jumlah Rekod : <%=kira%> </td>
    <td colspan=8>
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik30.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik30.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik30.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik30.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %>	
	</td>
  </tr></table>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
    		<td width="2%" class="hd1">Bil.</td>
			<td width="5%" class="hd1">No Kompaun </td>
			<td width="4%" class="hd1">Akta </td>
			<td width="4%" class="hd1">Jenis&nbsp; </td>
			<td width="30%" class="hd1">Nama </td>
			<td width="10%" class="hd1">Tarikh&nbsp; </td>
			<td width="11%" class="hd1">Tkh Bayar&nbsp; </td>
			<td width="9%" class="hd1">No Resit </td>
			<td width="4%" class="hd1">Amaun </td>
            <td width="15%" class="hd1">Jabatan </td>
  </tr>
<%		bil = 0
		ctrz = 0
	
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
			abayar = objRs2("amaun_bayar")
			
			jamaun = cdbl(jamaun) + cdbl(abayar)
			kodJbtn = objRs2("jabatan")
		
		q1="select keterangan from payroll.ptj where kod='"&kodJbtn&"'"
		set rq1 = objConn.execute(q1)
		
		if not rq1.eof then namaJabatan = rq1("keterangan")
			
%>
  <tr align="center"> 
  <td><%=bil%></td>
  <td><%=objRs2("no_akaun")%></td>  
  <%	kara = objRs2("perkara") 
  		if kara <> "" then
  %>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"><a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a></td> <%	else	%>
  <td><a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a></td>
  <%	end if		
  		rujuk2 = objRs2("no_rujukan2")
  		if rujuk2 <> "" then
  %>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"><a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRs2("no_rujukan2")%>">
  <%=objRs2("no_rujukan2")%></a></td>
  <%	else		%>
  <td><a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRs2("no_rujukan2")%>">
 <%=objRs2("no_rujukan2")%></a></td>
  <%	end if		%>
  <td><%=objRs2("nama")%></td>
  <td><%=objRs2("tkh_masuk")%></td>
  <td><%=objRs2("tkh_bayar")%></td>
  <td><%=objRs2("no_resit")%></td>
  <td align="right" ><%=FormatNumber(objRs2("amaun_bayar"),2)%></td>
  <td><%=namaJabatan%></td>
  </tr>
<% 	iRecordsShown = iRecordsShown + 1
	count = count + 1

  	objRs2.MoveNext			
  	Loop
	
		d2 = " select sum(nvl(amaun_bayar,0)) amaun_bayart "
		d2 = d2 & " from hasil.bil "
		d2 = d2 & " where tkh_bayar between to_date('"& tkh1 &"', 'ddmmyyyy') "
		d2 = d2 & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d2 = d2 & " and (perkara <> 'P01' or perkara is null) "
		d2 = d2 & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%') "
		d2 = d2 & " and amaun_bayar > 0 "
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
		d2 = d2 & " order by tkh_bayar "
		Set of2 = objConn.Execute(d2)
  %> 
  <tr>
  <td align="right" colspan="7">
  <b>Jumlah :&nbsp;RM </b></td>
  <td align="right"><b><%=FormatNumber(jamaun,2)%></b></td>

  </tr>
  <tr>
  <td align="right" colspan="7">
  <b>Jumlah Keseluruhan :&nbsp;RM </b></td>
  <td align="right"><b><%=FormatNumber(of2("amaun_bayart"),2)%></b></td>

  </tr></table>
  <%  	end if		
  		end if
  		end if
  		end if
  		end if 
  		end if 
  %>
</form>
<form method="post" action="ik29r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">

</td></tr>
</table>
</form>
</body>