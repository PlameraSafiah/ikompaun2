<%response.cookies("ikmenu") = "ik218.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Laporan Keseluruhan</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<style>
<!-- a {text-decoration:none}
//-->
</style>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "tkh1";
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln)=="(")break;
netscape = (ver.charAt(iln+1).toUpperCase()!="C");

function keyDown(DnEvents){
k = (netscape)?DnEvents.which : window.event.keyCode;
if(k==13){//enter key pressed
if (nextfield=='done') return true; //submit
else{//send focus to next box
eval('document.myform.'+nextfield + '.focus()');
return false;
	}
  }
 }
document.onkeydown = keyDown;// work together to analyze keystrokes
if (netscape)document.captureEvents(Event.KEYDOWN|Event.KEYUP);
//End-->
</script>
<script language="javascript">

function invalid_data(b)
  {  
       alert (b+" Tiada Rekod ");
		return(true);
  }
function invalid_nopekerja(c)
	{
		alert (c+" No Pekerja Salah !!! ");
		return(true);
	}
	function invalid_tarikh(d)
    {  
       alert (d+" Tarikh Salah !!! ");
		return(true);
    }	
</script>

</head>

<form name=myform method="POST" action="ik218.asp">
<%
	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   
	response.cookies("amenu") = "ik218.asp"
   	proses = Request.form("B1")
    nopek1 = request.cookies("gnop")


	if proses <> "Cari" then
		
		e = " select '01'||to_char(sysdate,'mm')||to_char(sysdate,'yyyy') as tkh1s , "
		e = e & " to_char(sysdate,'ddmmyyyy') as tkh2s from dual "
		Set objRse = Server.CreateObject ("ADODB.Recordset")
   		Set objRse = objConn.Execute(e)	   		
   		tkh1 = objRse("tkh1s")
   		tkh2 = objRse("tkh2s")
  		
	end if

'////kod jabatan penyedia////
	 np = " select lokasi from payroll.paymas where no_pekerja = '"&nopek1&"' "
	 Set np = objConn.Execute(np)
	 if not np.eof then
	  lokasi = np("lokasi")
	 end if

if lokasi = "103" then 
	kod = "76410"
elseif lokasi = "109" then
	kod = "76420"
elseif lokasi = "107" then
	kod = "76410"
end if

	if proses = "Cari" then
	
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")
	end if
	

%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">

<!--<script>
	document.myform.dae.focus();
</script>-->
<tr >  
<td class="hd">Tarikh Dari&nbsp;</td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" size="10" maxlength="8" onFocus="nextfield='tkhhingga';">&nbsp;Hingga
<input type="text" name="tkhhingga" value="<%=tkh2%>" size="10" maxlength="8" onFocus="nextfield='done';">&nbsp;
<input type="submit" value="Cari" name="B1" class="button"></td>
</tr></table>
  
<%	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 50

	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If
	



	if proses = "Cari" or tkh1 <> "" or tkh2 <> "" then
	
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
	
	if proses = "Cari" then
	
		'if proses = "CARI" then	
			
		d = " select no_akaun, perkara, no_rujukan2 ,initcap(nama) nama, to_char(tkh_masuk,'dd/mm/yyyy') tkh_masuk, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'dd/mm/yyyy') tkh_bayar, no_resit, "
		d = d & " nvl(amaun_bayar,0) amaun_bayart from hasil.bil "
		d = d & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		''d = d & " and (no_akaun like '"& kod &"'||'%' ) "
		'd = d & " and substr(no_akaun,5,2) = '"& dae &"' "
		d = d & " and (perkara <> 'P01' or perkara is null) "

		'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		''response.Write(d)
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		d = d & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
		d = d & " order by no_rujukan "
		Set objRs2 = Server.CreateObject ("ADODB.Recordset")
		
		
	
	'response.write "<p>dae"&dae&"</p>"
'response.write kod
'response.write tkh1
'response.write tkh2
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

			
			
		if proses = "Cari" and objRs2.bof and objRs2.eof then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
       	response.write "</script>"
		else

		if kira > 0 then

%>  
<%'response.write dae%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
  <tr align="right">
  	<td align="left" width="25%">Jumlah Rekod : <%=kira%></td>
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

 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="5%" class="hd1">Bil</td>
<td width="10%" class="hd1">No Kompaun</td>
<td width="3%" class="hd1">Akta</td>
<td width="3%" class="hd1">Jenis</td>
<td width="20%" class="hd1">Nama</td>
<td width="10%" class="hd1">Tarikh Masuk</td>
<td width="3%" class="hd1" colspan="2">Status Bayaran</td>
<td width="10%" class="hd1">Tarikh Bayaran</td>
<td width="10%" class="hd1" align="right">Amaun Bayar</td>
<td width="10%" class="hd1">No Resit</td>
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

%>
  <tr> 
  <td><%=bil%></td>
  <td><%=objRs2("no_akaun")%></td>
  <%	kara = objRs2("perkara")
  		if kara <>"" then		%>
  <td>
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></a></td>
	<%	else	%>
  <td width="10%" align="center" ><font size="2" color="Blue"><b><%=objRs2("perkara")%></b></font></td>
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
  <% tarikhbayar = objRs2("tkh_bayar")
  
  if tarikhbayar <> "" then %>
  
  <td align="center">Y</td>
  <td align="center">&nbsp;</td>
  <% else %>
  <td align="center">&nbsp;</td>
  <td align="center">T</td>
  <% end if %>
  
  <td align="center"><%=objRs2("tkh_bayar")%></td>
  <td align="right"><%=FormatNumber(objRs2("amaun_bayar"),2)%></td>
  <td><%=objRs2("no_resit")%></td>
  </tr>
<%	iRecordsShown = iRecordsShown + 1
	count = count + 1
		
  	objRs2.MoveNext			
  	Loop
	
		d2 =     " select sum(nvl(amaun_bayar,0)) amaun_bayart from hasil.bil "
		d2 = d2 & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d2 = d2 & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d2 = d2 & " and (no_akaun like '"& kod &"'||'%' ) "
		d2 = d2 & " and substr(no_akaun,5,2) = '"& dae &"' "
		d2 = d2 & " and (perkara <> 'P01' or perkara is null) "
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
		d2 = d2 & " order by no_rujukan "
		Set of2 = objConn.Execute(d2)

	
%> 
 <tr>
 <td colspan="6" align="right"><b>Jumlah</b></td>
 <td align="right"><b><%=FormatNumber(total_ab,2)%></b></td></tr>
  <tr>
 <td colspan="6" align="right"><b>Jumlah Keseluruhan</b></td>
 <td align="right"><b><%=FormatNumber(amaun_bayart,2)%></b></td></tr>

 </table>
<%  	end if		
		end if
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
<form method="post" action="ik30r.asp"> 
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

