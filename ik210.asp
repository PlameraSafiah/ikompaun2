<%response.cookies("ikmenu") = "ik210.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Ringkasan Mengikut Pegawai</title>
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
<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolorlight="#003366">
<tr valign="top"> 
<td width="100%">

<form name=myform method="POST" action="ik210.asp">
<%
	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   
   	response.cookies("amenu") = "ik210.asp"
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

	
	
	n = " select initcap(nama) nama from payroll.paymas where no_pekerja = '"&nopek&"' "
	n = n & " union "
	n = n & " select initcap(nama) nama from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' "
	Set objRsn = Server.CreateObject("ADODB.Recordset")
	Set objRsn = objConn.Execute(n)
	
	if not objRsn.eof then
		napek = objRsn("nama")
	else
		napek = ""
	end if		
%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr> 
            <td class="hd">No Pekerja</td>
            <td>
<input type="text" name="nop" value="<%=nopek%>" onFocus="nextfield='tkhdari';" size="10" maxlength="5">
<b>&nbsp;&nbsp;-&nbsp;&nbsp;<%=napek%></b></td>
</tr>
<script>
	document.myform.nop.focus();
</script>
<tr >  
            <td class="hd">Tarikh Dari</td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" onFocus="nextfield='tkhhingga';" size="10" maxlength="8">&nbsp; Hingga &nbsp;
              <input type="text" name="tkhhingga" value="<%=tkh2%>" onFocus="nextfield='done';" size="10" maxlength="8">&nbsp;
<input type="submit" value="Cari" name="B1" class="button"></td>
</tr></table>

<%	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 15

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
	
	
	
		k = " select no_pekerja from payroll.paymas where no_pekerja = '"&nopek&"' "
		k = k & " union "
		k = k & " select no_pekerja from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' "
		Set objRsk = Server.CreateObject("ADODB.Recordset")
		Set objRsk = objConn.Execute(k)
		
		if proses = "Cari" and objRsk.eof then	
			response.write "<script language=""javascript"">"
			response.write "var timeID = setTimeout('invalid_nopekerja(""  "");',1) "
			response.write "</script>"
			proses = "Cari"
			
		else
		
		d = " select perkara, no_rujukan2, count(1) bilsalah "
		d = d & " from hasil.bil "
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
		d = d & " group by perkara, no_rujukan2 "
		d = d & " order by perkara, no_rujukan2 "
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
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr align="right">
  	<td align="left" colspan=2 width="141">Jumlah Rekod : <%=kira%></td>
    <td colspan=8 width="572" >
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik210.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik210.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik210.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik210.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dnopek=<%=nopek%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %>	
	</td>
  </tr>
  </table>
   <table width="50%" cellpadding="1" cellspacing="5" class="hd">
<tr >  
  <td width="25%" class="hd">Bil</td>
  <td width="25%"  class="hd">Akta</td>
  <td width="25%"  class="hd">Jenis</td>
  <td width="25%"  class="hd">Bil Kesalahan</td>
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
  <td><%=bil%></font> </td>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'">
  <a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a></td>
  <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"> 
  <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRS2("no_rujukan2")%>"><%=objRs2("no_rujukan2")%></a></td>
  <td><%=objRs2("bilsalah")%></td>
  </tr>
<%	iRecordsShown = iRecordsShown + 1
	count = count + 1

  	objRs2.MoveNext			
  	Loop
%> 
</table>

</form>

<form method="post" action="ik210r.asp"> 
<p align="center"><input type="submit" value="Cetak" name="B2" class="button"></p>
<input type="hidden" name="fnopek" value="<%=nopek%>">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">
</form>

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

</body>