<%response.cookies("ikmenu") = "ik27.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Pengeluar Kompaun</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "nop";
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
       alert (a+" Masukkan No Pekerja !!! ");
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

<form name=myform method="POST" action="ik27.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top"> 
<td width="100%">


<%
	'Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   
	response.cookies("amenu") = "ik27.asp"
   	proses = Request.form("B1")
   		
	if proses <> "Cari" then
		
		e = " select '01'||to_char(sysdate,'mm')||to_char(sysdate,'yyyy') as tkh1s , "
		e = e & " to_char(sysdate,'ddmmyyyy') as tkh2s from dual "
		Set objRse = Server.CreateObject ("ADODB.Recordset")
   		Set objRse = objConn.Execute(e)	   		
   		tkh1 = objRse("tkh1s")
   		tkh2 = objRse("tkh2s")
  		
	end if

	
	if proses = "Cari" then
		nopek = Request.form("nop")
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")
	end if

	
	dnopek = Request.QueryString("dnopek")

	if dnopek <> "" then
		nopek = Request.QueryString("dnopek")
		tkh1 = Request.QueryString("dtkh1")
		tkh2 = Request.QueryString("dtkh2")
	end if

	
	'05062013 : bukak report ni utk akaun yang start dengan 76413 jugak
	'n = " select initcap(nama) nama from payroll.paymas where no_pekerja = '"&nopek&"' and lokasi = 103 "
	n = " select initcap(nama) nama from payroll.paymas where no_pekerja = '"&nopek&"' " 'and (lokasi = 103 or lokasi = 113)
	n = n & " union "
	'n = n & " select initcap(nama) nama from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' and lokasi = 103 "
	n = n & " select initcap(nama) nama from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' "'and (lokasi = 103 and lokasi = 113) 
	Set objRsn = Server.CreateObject("ADODB.Recordset")
	'response.Write(n)
	Set objRsn = objConn.Execute(n)
	
	if not objRsn.eof then
		napek = objRsn("nama")
	else
		napek = ""
	end if		
	

%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
    <td class="hd">No Pekerja&nbsp; </td>
<td><input type="text" name="nop" value="<%=nopek%>" size="10" maxlength="5" onFocus="nextfield='tkhdari';">
&nbsp;&nbsp;-&nbsp;&nbsp;<%=napek%></td>
</tr>
<script>
	document.myform.nop.focus();
</script>
<tr >  
<td class="hd">Tarikh Dari&nbsp; </td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" size="10" maxlength="8" onFocus="nextfield='tkhhingga';">Hingga
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
	



	if proses = "Cari" or nopek <> "" or tkh1 <> "" or tkh2 <> "" then
	
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
	
	if proses = "Cari" and Len(nopek) = 0 then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_nopek(""  "");',1) "
       	response.write "</script>"
       	proses = "Cari"
       	
	else
	
	
	
		k = " select no_pekerja from payroll.paymas where no_pekerja = '"&nopek&"' and lokasi = 101 "
		k = k & " union "
		k = k & " select no_pekerja from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' and lokasi = 101 "
		Set objRsk = Server.CreateObject("ADODB.Recordset")
		Set objRsk = objConn.Execute(k)
		
		if proses = "CARI" and objRsk.eof  then	
			response.write "<script language=""javascript"">"
			response.write "var timeID = setTimeout('invalid_nopekerja(""  "");',1) "
			response.write "</script>"
			proses = "Cari"
			
		else
		
		d = " select no_akaun, perkara, no_rujukan2 ,initcap(nama) nama, to_char(tkh_masuk,'dd/mm/yyyy') tkh_masuk, "
		d = d & " nvl(amaun_bayar,0) amaun_bayar, to_char(tkh_bayar,'dd/mm/yyyy') tkh_bayar, no_resit,jabatan "
		d = d & " from hasil.bil "
		d = d & " where no_pekerja = '"& nopek &"' "
		d = d & " and tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran-mimi-pn.raja
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%'or no_akaun like '76441'||'%') "
		d = d & " order by no_rujukan "
		'response.Write(d)
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

			
			
		if proses = "Cari" and objRs2.bof and objRs2.eof then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
       	response.write "</script>"
		else

		if kira > 0 then
%>  <br/>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
  <tr align="right">
  	<td align="left" colspan=2 width="143">Jumlah Rekod : <%=kira%></td>
    <td colspan=8 width="570" >
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik27.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik27.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik27.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik27.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %></font>	
	</td>
  </tr>
  </table>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="5%" class="hd1">Bil</td>
<td width="15%" class="hd1">No Kompaun</td>
<td width="10%" class="hd1">Akta</td>
<td width="10%" class="hd1">Jenis</td>
<td width="25%" class="hd1">Nama</td>
<td width="10%" class="hd1">Tarikh&nbsp;</td>
<td width="10%" class="hd1">Amaun</td>
<td width="10%" class="hd1">No Resit</td>
<td width="10%" class="hd1">Jabatan</td>
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
		ab = objRs2("amaun_bayar")
		total_ab = cdbl(total_ab) + cdbl(ab)
		
		kodJbtn = objRs2("jabatan")
		
		q1="select keterangan from payroll.ptj where kod='"&kodJbtn&"'"
		set rq1 = objConn.execute(q1)
		
		if not rq1.eof then namaJabatan = rq1("keterangan")
%>
 <tr align="center">
<td><%=bil%></td>
  <td><%=objRs2("no_akaun")%></td>
  <%	kara = objRs2("perkara")
  		if kara <>"" then		%>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'">
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a></td>
	<%	else	%>
  <td><%=objRs2("perkara")%></td>
  <%	end if		%>
  <%	ruj2 = objRs2("no_rujukan2")
  		if ruj2 <>"" then		%>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"> 
  <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRS2("no_rujukan2")%>">
 <%=objRs2("no_rujukan2")%></a></td>
	<%	else	%>
 <td><%=objRs2("no_rujukan2")%></td>
	<%	end if		%>
  <td><%=objRs2("nama")%></td>
  <td><%=objRs2("tkh_masuk")%></td>
  <td><%=FormatNumber(objRs2("amaun_bayar"),2)%></td>
  <td><%=objRs2("no_resit")%></td>
  <td><%=namaJabatan%></td>
  </tr>
<%	iRecordsShown = iRecordsShown + 1
	count = count + 1

  	objRs2.MoveNext			
  	Loop
	
		d2 = " select sum(nvl(amaun_bayar,0)) amaun_bayart "
		d2 = d2 & " from hasil.bil "
		d2 = d2 & " where no_pekerja = '"& nopek &"' "
		d2 = d2 & " and tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "		
		d2 = d2 & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d2 = d2 & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%') "
		d2 = d2 & " order by no_rujukan "
		Set of2 = objConn.Execute(d2)
%> 
 <tr>
 <td colspan="6" align="right"><b>Jumlah</b></td>
 <td align="right"><b><%=FormatNumber(total_ab,2)%></b></td></tr>
  <tr>
 <td colspan="6" align="right"><b>Jumlah Keseluruhan</b></td>
 <td align="right"><b><%=FormatNumber(of2("amaun_bayart"),2)%></b></td></tr>

</table>
<%  	end if		
		end if
		end if
		end if
  		end if  
  		end if
  		end if
  		end if
%>
</td>
</tr>
</table>  
</form>
<form method="post" action="ik27r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">
<input type="hidden" name="fno" value="<%=nopek%>">
</td></tr>
</table>
</form>
</body>
</html>