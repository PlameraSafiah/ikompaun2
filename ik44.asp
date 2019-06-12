<%response.cookies("ikmenu") = "ik44.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Penyelenggaraan Kod Jadual Pembayaran</title>
<script language="javascript">
	function papar1(form)
		{var item = form.drop1.selectedIndex; choice = form.drop1.options[item].value; if (choice!="x") 
		top.location.href=""+(choice); };
</script> 
<script language="javascript">
	function papar2(form)
		{var item = form.drop2.selectedIndex; choice = form.drop2.options[item].value; if (choice!="x")
		top.location.href=""+(choice); };
</script>
<script language="javascript">
   function invalid_data(a)
    {  
       alert (a+" Tiada Maklumat ");
		return(true);
    }
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" >
<table border="0" width="100%" height="15">
<tr><td width="100%" height="15"></td></tr>
</table>
<% 
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
	
	response.cookies("amenu") = "ik44.asp"	
	proses = Request.Form("B4")
		vrowid1 = request.form("rowid1")
         sdiv1 = request.form("sdiv1")
         sdiv2 = request.form("sdiv2")
         sdiv3 = request.form("sdiv3")
         snia1 = request.form("snia1")
         snia2 = request.form("snia2")
         snia3 = request.form("snia3")
         smak = request.form("smak")
         sdus1 = request.form("sdus1")
         sdus2 = request.form("sdus2")
         sdus3 = request.form("sdus3")              

	if proses = "kemaskini" then         
         		
         			b1 = " update kompaun.jenis_kesalahan "
					b1 = b1 & " set div1 = '"& sdiv1 &"', "
					b1 = b1 & "div2 = '"& sdiv2 &"', "
					b1 = b1 & "div3 = '"& sdiv3 &"', "
					b1 = b1 & "nia1 = '"& snia1 &"', "
					b1 = b1 & "nia2 = '"& snia2 &"', "
					b1 = b1 & "nia3 = '"& snia3 &"', "
					b1 = b1 & "dus1 = '"& sdus1 &"', "
					b1 = b1 & "dus2 = '"& sdus2 &"', "
					b1 = b1 & "dus3 = '"& sdus3 &"', "
					b1 = b1 & "maksima = '"& smak &"' "
					b1 = b1 & " where rowid = '"& vrowid1 &"' "
						
					Set objRs1b = Server.CreateObject("ADODB.Recordset")
					Set objRs1b = objConn.Execute(b1)
					
		kod11 = Request.QueryString("tuju") 
		kod22 = Request.QueryString("tuju1")
		
		if kod22 = "" then
			kod22 = Request.form("fjab")	
		end if					
  		end if	
%> 
<table width="80%" cellspacing="1">
<form method="POST" action="ik44.asp">
<tr><td width="11%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>Akta/UUK</b></font></td> 
<td width="69%" bgcolor="lightgrey">
<select size="1" name="drop1" onChange="papar1(this.form);">
<option selected value="">Sila Pilih </option>
<%  mtuju = request.querystring("tuju")
	Set objRszz = Server.CreateObject("ADODB.Recordset")
	zz = "select kod,initcap(keterangan) as terang from kompaun.perkara order by kod"
	Set objRszz = objConn.Execute(zz)	
	do while not objRszz.EOF	
	if mtuju <> "" and mtuju = objRszz("kod") then%>
	<option selected value="ik44.asp?tuju=<%=objRszz("kod")%>"><%=objRszz("kod")%>-<%=objRszz("terang")%></option>
	<%else%>
<option value="ik44.asp?tuju=<%=objRszz("kod")%>"><%=objRszz("kod")%>-<%=objRszz("terang")%></option>
<%end if
  objRszz.Movenext
  loop
 objRszz.close 
 
 kod11 = Request.QueryString("tuju")%>
</select>
</td></tr>
</form>
<form method="Post" action="ik44.asp?tuju=<%=kod11%>">
<tr><td width="11%" bgcolor="lightgrey"><font face="Verdana" size="2"><b>Kesalahan</b></font></td>
<td width="69%" bgcolor="lightgrey">
<select size="1" name="drop2" onChange="papar2(this.form);">
<option selected value="">Pilih Kesalahan</option>
<%
	jenis = request.querystring("tuju1")
	b = "select kod,keterangan from kompaun.jenis_kesalahan "
	b = b & " where perkara = '"&kod11&"' order by kod "
	Set objRs1 = Server.CreateObject("ADODB.Recordset")
	Set objRs1 = objConn.Execute(b)

	 do while not objRs1.eof 
	 if jenis <> "" and jenis = objRs1("kod") then %>
<option selected value="ik44.asp?tuju=<%=kod11%>&tuju1=<%=objRs1("kod")%>"><%=objRs1("kod")%>
- <%=objRs1("keterangan")%></option>
<%   else %>
<option value="ik44.asp?tuju=<%=kod11%>&tuju1=<%=objRs1("kod")%>"><%=objRs1("kod")%>
- <%=objRs1("keterangan")%></option>
<% end if 
   objRs1.Movenext
   loop
   objRs1.close 
	 %>
</select>
</td></tr>
</form>
<%	proses = Request.Form("B4")
	kod22 = Request.QueryString("tuju1")
	kodee = kod22	
	
	if kod11 <> "" and kod22 <> "" then 
	bb = "select rowid,kod,initcap(keterangan) as terang1,nvl(div1,0)as div1,nvl(div2,0) as div2,nvl(div3,0) as div3, "
	bb = bb & " nvl(nia1,0) as nia1,nvl(nia2,0) as nia2,nvl(nia3,0) as nia3, "
	bb = bb & " nvl(dus1,0) as dus1,nvl(dus2,0) as dus2,nvl(dus3,0) as dus3,nvl(maksima,0) as maksima "
	bb = bb & " from kompaun.jenis_kesalahan where perkara = '"& kod11 &"' and kod = '"& kod22 &"' "		
	
	Set objRs2 = Server.CreateObject("ADODB.Recordset")
	Set objRs2 = objConn.Execute(bb)
 	terang = objRs2("terang1")
 		 	
		if objRs2.eof then
			response.write "<script language=""javascript"">"
        	response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
       	response.write "</script>"
      		     		
		end if 
%>
</table>
<hr width="98%" align="left" color="#660000">
<p></p>
<form method="Post" action="ik44.asp?tuju=<%=kod11%>&tuju1=<%=kod22%>">
  <table width="86%" align="center" border="1" bgcolor="#9D2024">
    <tr bgcolor="#330000"> 
      <td width="10%" align="center" bgcolor="#9D2024"></td>
      <td width="16%" align="center" bgcolor="#9D2024"> <font face="verdana"><b><font color="#FFFFFF"><font size="2">kesalahan(1)</font> 
        <font size="1" color="#FFFFFF">(RM)</font></font></b></font></td>
      <td width="18%" align="center" bgcolor="#9D2024"> <font face="verdana"><b><font color="#FFFFFF"><font size="2">kesalahan(2)</font> 
        <font size="1" color="#FFFFFF">(RM)</font></font></b></font></td>
      <td width="19%" align="center" bgcolor="#9D2024"> <font face="verdana"><b><font color="#FFFFFF"><font size="2">kesalahan(3)</font> 
        <font size="1" color="#FFFFFF">(RM)</font></font></b></font></td>
      <td width="18%" align="center" bgcolor="#9D2024"> <font face="verdana"><b><font color="#FFFFFF"><font size="2">Maksima</font> 
        <font size="1" color="#FFFFFF">(RM)</font></font></b></font></td>
      <td width="19%" align="center" bgcolor="#9D2024"><font face="verdana"></font></td>
    </tr>
    <tr> 
      <td width="10%" bgcolor="#9D2024"  align="center"> <font face="Verdana" size="2"><b><font color="#FFFFFF">INDIVIDU</font></b></font></td>
      <td width="16%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdiv1" value="<%=formatNumber(objRs2("div1"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdiv2" value="<%=formatNumber(objRs2("div2"),2)%>">
        </b></font></td>
      <td width="19%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdiv3" value="<%=formatNumber(objRs2("div3"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"><font face="verdana"></font></td>
      <td width="19%"  align="center" bgcolor="#9D2024"><font face="verdana"></font></td>
    </tr>
    <tr> 
      <td width="10%" bgcolor="#9D2024"  align="center"> <font face="Verdana" size="2"><b> 
        <font color="#FFFFFF">PENIAGA</font></b></font></td>
      <td width="16%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="snia1" value="<%=formatNumber(objRs2("nia1"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="snia2" value="<%=formatNumber(objRs2("nia2"),2)%>">
        </b></font></td>
      <td width="19%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="snia3" value="<%=formatNumber(objRs2("nia3"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="smak" value="<%=formatNumber(objRs2("maksima"),2)%>">
        </b></font></td>
      <td width="19%" align="center" bgcolor="#9D2024"> <font face="verdana"> 
        <input type="submit"  name="B4" value="kemaskini" >
        <input type="hidden" name="rowid1" value="<%=objRs2("rowid")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="10%" bgcolor="#9D2024"  align="center"> <font face="Verdana" size="2"><b> 
        <font color="#FFFFFF">INDUSTRI</font></b></font></td>
      <td width="16%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdus1" value="<%=formatNumber(objRs2("dus1"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdus2" value="<%=formatNumber(objRs2("dus2"),2)%>">
        </b></font></td>
      <td width="19%" bgcolor="lightgrey"  align="center"> <font face="verdana"><b> 
        <input type="text" size="9" name="sdus3" value="<%=formatNumber(objRs2("dus3"),2)%>">
        </b></font></td>
      <td width="18%" bgcolor="lightgrey"  align="center"><font face="verdana"></font></td>
      <td width="19%"  align="center" bgcolor="#9D2024"><font face="verdana"></font></td>
    </tr>
    <%		end if		%>
  </table>
</form>
</body>




































































