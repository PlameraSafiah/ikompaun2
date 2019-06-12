<%response.cookies("ikmenu") = "ik26test.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Komapun, Nama & Alamat</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "no_rujuk2";
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
function invalid_data(a)
  {  
       alert (a+" Tiada Rekod ");
		return(true);
  }
function invalid_input(b)
  {  
       alert (b+" Masukkan Pilihan !!! ");
		return(true);
  }
function invalid_akaun(b)
  {  
       alert (b+" Sila Masukkan Kod Akaun Kompaun Yg Betul !!! ");
		return(true);
  }
</script>

</head>

<form name=myform method="POST" action="ik26test.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top"> 
<td width="100%">



<%
	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   	
   	proses = Request.form("B1")
 	if proses="Reset" then
     no_rujuk2=""
	 nama=""
	 alamat=""
	 
	end if
   	
   	if proses = "Cari" then
   	
   		rujuk = Request.form("no_rujuk2")
   		nama = ucase(Request.form("nama"))
   		add = ucase(Request.form("alamat"))
   	
	end if
	
	dadd= Request.QueryString("dadd")
	dnama = Request.QueryString("dnama")
	drujuk = Request.QueryString("drujuk")
	
	if dadd <> "" then
		add = Request.QueryString("dadd")
	end if
	
	if dnama <> "" then
		nama = Request.QueryString("dnama")
	end if	

	if drujuk <> "" then
		rujuk = Request.QueryString("drujuk")
	end if
 	
 	nama2=Replace(nama,"'","''")
 	add2=Replace(add,"'","''")

%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
    <td class="hd">No Kompaun</td> 
<td><input type="text" name="no_rujuk2" value="<%=rujuk%>" onFocus="nextfield='nama';" maxlength="11" size="12"></td>
</tr>
<script>
	document.myform.no_rujuk2.focus();
</script>
<tr>    
<td class="hd">Nama&nbsp;</td>
<td><input type="text" name="nama" value="<%=nama%>" onFocus="nextfield='alamat';" size="50" maxlength ="50"></td>
</tr>  
<tr>
<td class="hd">Alamat</td>
<td><input type="text" name="alamat" value="<%=add%>" onFocus="nextfield='done';" size="40" maxlength="40" >&nbsp;&nbsp;&nbsp;&nbsp;
  <input type="submit" value="Cari" name="B1" class="button">
              <input name="B1" type="submit" id="B1" class="button" value="Reset"></td></tr>
</table>
<%
	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 10

	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If
	
	if proses = "Cari" or rujuk <> "" or nama <> "" or add <> "" then
		
		if Len(rujuk) =0 and Len(nama) = 0 and Len(add) = 0 then
			response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_input(""  "");',1) "
       	response.write "</script>"
       	proses = "Cari"

		else
		
		d =		   " select no_akaun,rowid, perkara, no_rujukan2 , initcap(nama) nama, to_char(tkh_masuk,'dd/mm/yyyy') tkh_masuk, " 
		d = d & " nvl(amaun_bayar,0) amaun_bayar, decode(tkh_bayar,null,'0',to_char(tkh_bayar,'dd/mm/yyyy')) tkh_bayar, no_resit "
		d = d & " from hasil.bil "
		d = d & " where (perkara <> 'P01' or perkara is null) "
		d = d & " and no_akaun like '"& rujuk &"'||'%' "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%') "
		'18042014 : jun tambah boleh papar kod 76413 jugak, sebelum ni 76410 & 76420 saja
		d = d & " and nama like '"& nama2 &"'||'%' "
		d = d & " and alamat1||alamat2||alamat3 like decode('"& add2 &"',null, "
		d = d & " alamat1||alamat2||alamat3,'%'||'"& add2 &"'||'%') "	
		d = d & " order by no_rujukan "  			
		Set objRs2 = Server.CreateObject("ADODB.Recordset")
		
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
       	proses = "CARI"
   
   		else    
   		
   		if kira > 0 then
%>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
  <tr align="right">
  	<td colspan="3" align="left" >Jumlah Rekod : <%=kira%></td>
    <td colspan="6">
	  <% If iPageCurrent <> 1 Then %>
      <a href="ik26test.asp?page=1&bilangan=0&ms=pre&drujuk=<%=rujuk%>&dnama=<%=nama%>&dadd=<%=add%>&proses=Cari">
		<img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a>  
      <% End If %>
      <% If iPageCurrent <> 1 Then%>
      <a href="ik26test.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&drujuk=<%=rujuk%>&dnama=<%=nama%>&dadd=<%=add%>&proses=Cari">
      	<img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
      <% End If %>
      Halaman <%=iPageCurrent%>/<%if iPageCount=0 then%>1<%else%><%=iPageCount%><%end if%>
      <% If iPageCurrent < iPageCount Then	%>
      <a href="ik26test.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&drujuk=<%=rujuk%>&dnama=<%=nama%>&dadd=<%=add%>&proses=Cari">
      	<img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
      <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
      <a href="ik26test.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&drujuk=<%=rujuk%>&dnama=<%=nama%>&dadd=<%=add%>&proses=Cari">
      	<img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
      <% End If %>	
      </font>	
	</td>
  </tr>
</table>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="27" class="hd1">Bil </td>
<td width="84" class="hd1">No Kompaun </td>
<td width="47" class="hd1">Akta </td>
<td width="49" class="hd1">Jenis </td>
<td width="184" class="hd1">Nama </td>
<td width="59" class="hd1">Tarikh </td>
<td width="67" class="hd1">Amaun </td>
<td width="63" class="hd1">Tkh Bayar </td>
<td width="56" class="hd1">No Resit </td>
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
 <tr align="center">
<td><%=bil%></td>
<!--<td width="84" bgcolor="#CCCCCC" align="center"><font face="Verdana" size="2"><%=objRs2("no_akaun")%></font></td>-->
<td onMouseover="this.style.backgroundColor='CC6666'" onMouseout="this.style.backgroundColor='FFE9E8'"><a href="ik26btest.asp?kod=<%=objRs2("no_akaun")%>"><%=objRs2("no_akaun")%></a>&nbsp;</td>
<%	kara = objRs2("perkara")
	if kara <> "" then	%>
<td onMouseover="this.style.backgroundColor='CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'">
<a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"><%=objRs2("perkara")%></a>&nbsp;</td>
<%	else	%>
<td><%=objRs2("perkara")%></td>
<%	end if	%>
<%	ruj2 = objRs2("no_rujukan2")
	if ruj2 <> "" then	%>
<td onMouseover="this.style.backgroundColor='CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'" align="center">
<a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRs2("no_rujukan2")%>"><%=objRs2("no_rujukan2")%>&nbsp;</a></td>
<%	else	%>
<td><%=objRs2("no_rujukan2")%>
<%	end if

t_bayar = objRs2("tkh_bayar")
if t_bayar = "0" then
 tb="-"
 else
 tb=t_bayar
 end if	%></td>
<td align="right"><%=objRs2("nama")%></td>
<td><%=objRs2("tkh_masuk")%></td>
<td><%=FormatNumber(objRs2("amaun_bayar"),2)%></td>
<td><%=tb%></td>
<td><%=objRs2("no_resit")%></td>
</tr>
<%		iRecordsShown = iRecordsShown + 1
		count = count + 1
		ab = objRs2("amaun_bayar")
		total_ab = cdbl(total_ab) + cdbl(ab)


  		objRs2.MoveNext			
  		Loop
		
		d2 =		   " select " 
		d2 = d2 & " sum(nvl(amaun_bayar,0)) amaun_bayart "
		d2 = d2 & " from hasil.bil "
		d2 = d2 & " where (perkara <> 'P01' or perkara is null) "
		d2 = d2 & " and no_akaun like '"& rujuk &"'||'%' "
		d2 = d2 & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%') "
		d2 = d2 & " and nama like '"& nama2 &"'||'%' "
		d2 = d2 & " and alamat1||alamat2||alamat3 like decode('"& add2 &"',null, "
		d2 = d2 & " alamat1||alamat2||alamat3,'%'||'"& add2 &"'||'%') "	
		d2 = d2 & " order by no_rujukan "  
		'response.write d2			
		Set of2 = objConn.Execute(d2)

%> 
  <tr>
 <td colspan="6" align="right"><b>Jumlah</b></td>
 <td align="right"><b>RM <%=FormatNumber(total_ab,2)%></b></td></tr>
  <tr>
    <tr>
 <td colspan="6" align="right"><b>Jumlah Keseluruhan</b></td>
 <td align="right"><b>'''<%=of2("amaun_bayart")%><%=FormatNumber(of2("amaun_bayart"),2)%></b></td></tr>

</table>
  
<%  	end if
		end if
  		end if		 
 		end if
 		'end if  		
%>
</td>
</tr>
</table>
</form>

<form method="post" action="ik26r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button">
<input type="hidden" name="frujuk1" value="<%=rujuk%>">
<input type="hidden" name="fnama2" value="<%=nama2%>">
<input type="hidden" name="fadd2" value="<%=add2%>">
</td></tr>
</table>
</form>
<%end if  		%>
</body>

</html>