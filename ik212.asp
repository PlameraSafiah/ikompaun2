<%response.cookies("ikmenu") = "ik212.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Kompaun Mengikut Daerah</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<style>
<!-- a {text-decoration:none}
//-->
</style>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "dae";
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
function invalid_nopek(a)
  {  
       alert (a+" Masukkan Daerah !!! ");
		return(true);
  }
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
<form name=myform method="POST" action="ik212.asp">
<%
	'Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   
	response.cookies("amenu") = "ik212.asp"
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
		dae = Request.form("dae")
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")
	end if

	
	ddae = Request.QueryString("ddae")

	if ddae <> "" then
		dae = Request.QueryString("ddae")
		tkh1 = Request.QueryString("dtkh1")
		tkh2 = Request.QueryString("dtkh2")
	end if

	
	

%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
<td class="hd">Daerah</td>
<td>
	  <input type="radio" name="dae" value="01" <%if dae="01" then%>checked<%end if%>>&nbsp;SPU&nbsp;
      <input type="radio" name="dae" value="02" <%if dae="02" then%>checked<%end if%>>&nbsp;SPT&nbsp;
      <input type="radio" name="dae" value="03" <%if dae="03" then%>checked<%end if%>>&nbsp;SPS&nbsp;
</td>
</tr>
<!--<script>
	document.myform.dae.focus();
</script>-->
<tr >  
<td class="hd">Tarikh Dari</td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" size="10" maxlength="8" onFocus="nextfield='tkhhingga';"> Hingga <input type="text" name="tkhhingga" value="<%=tkh2%>" size="10" maxlength="8" onFocus="nextfield='done';">&nbsp;
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
	



	if proses = "Cari" or dae <> "" or tkh1 <> "" or tkh2 <> "" then
	
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
	
	if proses = "Cari" and dae = "" then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_nopek(""  "");',1) "
       	response.write "</script>"
       	proses = "Cari"
       	
	else
	

		
		'if proses = "CARI" then	
			
		d = " select perkara, no_rujukan2, count(1) bilsalah "
		d = d & " from hasil.bil "
		d = d & " where tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%'or no_akaun like '76441'||'%') "
		d = d & " and substr(no_akaun,5,2) = '"& dae &"' "
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
		d = d & " group by perkara, no_rujukan2  "
		d = d & " order by perkara, no_rujukan2  "
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
  	<td align="left" colspan=2 width="143">Jumlah Rekod : <%=kira%></td>
    <td align="right" colspan=2 width="570" >
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik212.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&ddae=<%=dae%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik212.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&ddae=<%=dae%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik212.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&ddae=<%=dae%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik212.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&ddae=<%=dae%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %>
	</td>
  </tr>
<tr >  
</table>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="10%" class="hd1">Bil</td>
<td width="30%" class="hd1">Akta</td>
<td width="30%" class="hd1">Jenis</td>
<td width="30%" class="hd1">Bil Kesalahan</td>
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
%>
  <tr> 
  <td align="center"><%=bil%></td>
  <td align="center" onMouseover="this.style.backgroundColor='#CCCCCC'" onMouseout="this.style.backgroundColor='#FFE9E8'">
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a></td>
  <td onMouseover="this.style.backgroundColor='#CCCCCC'" onMouseout="this.style.backgroundColor='#FFE9E8'"> 
  <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRS2("no_rujukan2")%>">
 <%=objRs2("no_rujukan2")%></a></td>
  <td align="center"><%=objRs2("bilsalah")%></td>
  </tr>
<%	iRecordsShown = iRecordsShown + 1
	count = count + 1

  	objRs2.MoveNext			
  	Loop
%> 
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
<form method="post" action="ik212r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">
<input type="hidden" name="fdae" value="<%=dae%>">
<input type="hidden" name="fkod" value="<%=kod%>">
</td></tr>
</table>
</form>
</body>
</html>



















































