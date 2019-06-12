<%response.cookies("ikmenu") = "ik219.asp"%>
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
function dopopup1() {
	window.open('ik219r.asp','LAMPIRAN','toolbar=no,location=no,status=no,menu bar=no,scrollbars=yes,left=50,top=50,width=800,height=600')
}

function dopopup2() {
	window.open('ik219rr.asp','LAMPIRAN','toolbar=no,location=no,status=no,menu bar=no,scrollbars=yes,left=50,top=50,width=800,height=600')
}

function dopopup3() {
	window.open('ik219rrr.asp','LAMPIRAN','toolbar=no,location=no,status=no,menu bar=no,scrollbars=yes,left=50,top=50,width=800,height=600')
}
function dopopup4() {
	window.open('ik219_excel3.asp','LAMPIRAN','toolbar=no,location=no,status=no,menu bar=no,scrollbars=yes,left=50,top=50,width=800,height=600')
}
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

<form name=myform method="POST" action="ik219.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top"> 
<td width="100%">



<p>
  <%


	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
 	
	session("pekz1") = pekz
 
	
   	proses = Request.form("B1")
	proses3 = Request.form("B3")
	
 	if proses="Reset" then
     no_rujuk2=""
	 nama=""
	 noic=""
	 alamat =""
	 
	end if
	
	mula
	if proses = "Nama" then
		subnama 
	end if 
	
	if proses = "No Kad Pengenalan" then
		subnoic 
	end if 
	
	if proses = "No Pendaftaran Syarikat" then
		subnoreg
	end if 
	
	if proses3 ="Muat Turun Excel" then
	response.Write("mimi")
	response.redirect "ik219_excel3.asp"
	end if 
	
	
	



%>
  <% sub mula 


		%>
</p>
<table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr align="center" bgcolor="<%=color2%>"> 
<td class="hd" height="40">
<input type="submit" value="Nama" name="B1" class="button2">
<input type="submit" value="No Kad Pengenalan" name="B1" class="button1">
<input type="submit" value="No Pendaftaran Syarikat" name="B1" class="button1">
<input type="submit" value="Reset" name="B1" class="button">
</td></tr></table>
<% end sub %>


<% sub subnoreg 


	
	
 det = "No Pendaftaran Syarikat"
		'******************************************************************
		'mimi tambah user view jabatan masing2.admin view semua (09082018)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		
				
		end if
		'end view ikut jabatan
		'******************************************************************
		

		a =		"select count(distinct(no_akaun)) as no_akaun,perkara5 from hasil.bil"
		a = a & " where "
		a = a & " perkara5 is not null and perkara5 <> '0' and perkara5 <> '-' "
		a = a & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		a = a & " or no_akaun like '76415'||'%' or no_akaun like '76412'||'%' or no_akaun like '76416'||'%'"
		a = a & " or no_akaun like '76101'||'%' )" 'or no_akaun like '76441'||'%') "		d = d & " and tkh_bayar is null "
		a = a & " and tkh_bayar is null "
		a = a & " and perkara <> 'P01'  "
		a = a & " and jabatan = '"& lok &"' "
		a = a & " group by perkara5 "
		a = a & " having count(distinct(no_akaun)) >= '2'"
		a = a & " order by no_akaun desc "
		'response.Write(a)
		set objRa = objConn.execute(a)
		

%>
<br>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr align="center"> 
    <td width="1" class="hd1">Bil </td>
<td width="10" class="hd1">Bil Kompaun</td>
<td width="84" class="hd1">No Kompaun </td>
</tr>
<%		

bil = 0
		ctrz = 0
do while not objRa.eof 


		
		bil = bil + 1
		ctrz = cdbl(ctrz) + 1
		
		
%>
<tr>
<td align="center"><%=bil%></td>
<td align="center"><%=objRa("no_akaun")%></td>
<td><a href="ik219p.asp?brn=<%=objRa("perkara5")%> "><%=objRa("perkara5")%></a></td>
<tr>

<%
objRa.movenext
loop
%>

</table>
  
  </td>

</table>
</form>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr>
      <td align="center" class="hd" valign="middle"> 
        <input type="submit" value="Cetak" name="B2" class="button" onClick="dopopup3();" >
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ik219_excel3.asp"><font color="#FFFFFF" class="button2" style="font-size:14px"> Export Ke Excel </font></a>
</td></tr>
</table>

<%  	


end sub		
%>

<% sub subnoic 
 det = "No Kad Pengenalan"

		'******************************************************************
		'mimi tambah user view jabatan masing2.admin view semua (09082018)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		
				
		end if
		'end view ikut jabatan
		'******************************************************************

		a =		"select count(distinct(no_akaun)) as no_akaun,replace(kp,'-','')kp from hasil.bil"
		a = a & " where "
		a = a & " kp is not null and kp <> '0' and kp <> '-' "
		a = a & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		a = a & " or no_akaun like '76415'||'%' or no_akaun like '76412'||'%' or no_akaun like '76416'||'%'"
		a = a & " or no_akaun like '76101'||'%' )" 'or no_akaun like '76441'||'%') "		d = d & " and tkh_bayar is null "
		a = a & " and tkh_bayar is null "
		a = a & " and perkara <> 'P01'  "
		a = a & " and jabatan = '"& lok &"' "
		a = a & " group by kp "
		a = a & " having count(distinct(no_akaun)) >= '2'"
		a = a & " order by no_akaun desc "
		'response.Write(a)
		set objRa = objConn.execute(a)
		

%>
<br>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr align="center"> 
    <td width="1" class="hd1">Bil </td>
<td width="27" class="hd1">Bil Kompaun</td>
<td width="84" class="hd1"><%=det%> </td>
</tr>
<%		

bil = 0
		ctrz = 0
do while not objRa.eof 


		
		bil = bil + 1
		ctrz = cdbl(ctrz) + 1
		
		
%>
<tr>
<td align="center"><%=bil%> </td>
<td align="center"><%=objRa("no_akaun")%></td>
<td><a href="ik219p.asp?nokp=<%=objRa("kp")%> "><%=objRa("kp")%></a></td>
<tr>

<%
objRa.movenext
loop
%>

</table>
</td>
</tr>
</table>
</form>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button" onClick="dopopup2();" >
        
        
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ik219_excel2.asp"><font color="#FFFFFF" class="button2" style="font-size:14px"> Export Ke Excel </font></a>
</td></tr>
</table>

<%  	


end sub		
%>


<% sub subnama

 det = "Nama"
		'******************************************************************
		'mimi tambah user view jabatan masing2.admin view semua (09082018)
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		
				
		end if
		'end view ikut jabatan
		'******************************************************************

		a =		"select count(distinct(no_akaun)) as no_akaun,nama from hasil.bil"
		a = a & " where "
		a = a & " nama is not null and nama <> '0' and nama <> '-' "
		a = a & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		a = a & " or no_akaun like '76415'||'%' or no_akaun like '76412'||'%' or no_akaun like '76416'||'%'"
		a = a & " or no_akaun like '76101'||'%' )" 'or no_akaun like '76441'||'%') "		d = d & " and tkh_bayar is null "
		a = a & " and tkh_bayar is null "
		a = a & " and perkara <> 'P01'  "
		a = a & " and jabatan = '"& lok &"' "
		a = a & " group by nama "
		a = a & " having count(distinct(no_akaun)) >= '2'"
		a = a & " order by no_akaun desc "
		'response.Write(a)
		set objRa = objConn.execute(a)
		

%>
<br>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr align="center"> 
    <td width="1" class="hd1">Bil</td>
<td width="27" class="hd1">Bil Kompaun</td>
<td width="84" class="hd1">Nama </td>
</tr>
<%		

		bil = 0
		ctrz = 0
do while not objRa.eof 


		
		bil = bil + 1
		ctrz = cdbl(ctrz) + 1
		
		
%>
<tr>
<td align="center"><%=bil%></td>
<td align="center"><%=objRa("no_akaun")%></td>
<td><a href="ik219p.asp?nama=<%=objRa("nama")%> "><%=objRa("nama")%></a></td>
<tr>

<%
objRa.movenext
loop
%>

</table>
  



</td>
</tr>
</table>


 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" value="Cetak" name="B2" class="button" onClick="dopopup1();" >
        
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ik219_excel1.asp"><font color="#FFFFFF" class="button2" style="font-size:14px"> Export Ke Excel </font></a>
</td></tr>
</table>

<%  	
end sub		
%>
</form>
<%'end if  		%>
</body>

</html>