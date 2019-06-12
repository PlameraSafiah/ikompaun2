<%response.cookies("ikmenu") = "ik41.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Penyelenggaraan Akta/UUK</title>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin

nextfield = "fkod";
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln)=="(")break;
netscape = (ver.charAt(iln+1).toUpperCase()!="C");

function keyDown(DnEvents){
k = (netscape)?DnEvents.which : window.event.keyCode;
if(k==13){//enter key pressed
if (nextfield=='done') return true; //submit
else{//send focus to next box
eval('document.form.'+nextfield + '.focus()');

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
		alert(a+" Kod Telah Wujud !!!");
		return(true);
	}
function invalid_keter(b)
	{
		alert(b+" Masukkan Keterangan Kod !!!");
		return(true);
	}
</script>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">

<table border="0" width="100%" height="15">
<tr><td width="100%" height="15"></td></tr>
</table>


<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolorlight="#003366" >
<tr valign="top"> 
          
<td width="100%"> 
<form name="form" method ="Post" action="ik41.asp">
<% 
	Set objConn = Server.CreateObject ("ADODB.Connection")
  	ObjConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

	proses = Request.Form("B1")
	proces = Request.Form("B2")
	response.cookies("amenu") = "ik41.asp"

	
	if proces = "res" then	
	vkod = ucase(Request.Form("fkod"))
	vketer = ucase(Request.Form("fketer"))
	vrowid = request.form("rowid")
	
		vkod = ""
		vketer = ""
		vrowid = ""
	end if			
	
	
	if proses <> "+" then
		bilrec = request.form("bilrec")
	else
		bilrec = ""
	end if
	
	
	
	if proses = "+" then
	vkod = ucase(Request.Form("fkod"))
	vketer = ucase(Request.Form("fketer"))
	vrowid = request.form("rowid")
	
		if vrowid = "" then			
			if vkod <> "" then
				if vketer <> "" then
				
				d = " select kod from kompaun.perkara where kod = '"&vkod&"' "
				Set objRsd = Server.CreateObject("ADODB.Recordset")
				Set objRsd = objConn.Execute(d)	
					if not objRsd.eof then
						response.write "<script language=""javascript"">"
						response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
						response.write "</script>"
						proses = "+"
					else
						a1 = " insert into kompaun.perkara(kod,keterangan) values ('"&vkod&"','"&vketer&"') "
						Set objRs1a = Server.CreateObject("ADODB.Recordset")
						Set objRs1a = objConn.Execute(a1)
					end if	
				
				else
					response.write "<script language=""javascript"">"
					response.write "var timeID = setTimeout('invalid_keter(""  "");',1) "
					response.write "</script>"
					proses = "+"
				end if
			end if
		else
		
		b1 = " update kompaun.perkara "
		b1 = b1 & " set kod = '"&vkod&"', "
		b1 = b1 & "     keterangan = '"&vketer&"' "
		b1 = b1 & " where rowid = '"&vrowid&"' "
		Set objRs1b = Server.CreateObject("ADODB.Recordset")
		Set objRs1b = objConn.Execute(b1)	
		
		end if
		proses1 = ""
	end if	
	
 
 '*******************************************************************************************************
 
	
	if bilrec <> "" then
     for i = 1 to bilrec
         rowid = "rowid" + cstr(i)
         kod = "kod" + cstr(i)
         keter = "keter" + cstr(i)
         b1 = "b1" + cstr(i)
         b2 = "b2" + cstr(i)
        
         rowid = request.form(""& rowid &"")
         vkod = request.form(""& kod &"")
         vketer = request.form(""& keter &"")
         b1 = request.form(""& b1 &"")
         b2 = request.form(""& b2 &"")
       

         if b1 = "-" then
     
            db = " delete kompaun.perkara where rowid = '"& rowid &"' "
            Set objRsdb = Server.CreateObject("ADODB.Recordset")
            Set objRsdb = objConn.Execute(db)
         elseif b2 = "e" then 
         		fkod = vkod
         		fketer = vketer
         		frowid = rowid
          		proses1 = "+"
         end if
     next
  end if
 
 
 '*****************************************************************************************************
 
		
	b = " select rowid, kod, initcap(keterangan) keterangan from kompaun.perkara order by kod"
	Set objRs2 = Server.CreateObject("ADODB.Recordset")
	Set objRs2 = objConn.Execute(b)

%>

<table width="80%" align="center">
<tr> 
<td width="5%" bgcolor="#FFFFFF"  align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Kod</b></font></td>
<td width="40%" bgcolor="#FFFFFF" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Keterangan</b></font></td>
<td width="15%" bgcolor="#FFFFFF" align="center"></td>
</tr>
<%	if proses1 <> "" then		%>
<tr> 
<td width="5%" bgcolor="#FFFFFF" align="center"><input type="text" name="fkod" size="3" value="<%=fkod%>" maxlength="3" onFocus="nextfield='fketer';"></td>
<td width="40%" bgcolor="#FFFFFF" align="center"><input type="text" name="fketer" size="66" value="<%=fketer%>" maxlength="65" onFocus="nextfield='B1';"></td>
<td width="15%" bgcolor="#FFFFFF" align="center">
<input type="submit" value="+" name="B1" onFocus="nextfield='done';"><input type="submit" value="res" name="B2">
<input type="hidden" name="rowid" value="<%=frowid%>" ></td>
</tr>
<%		else	%>
<tr> 
<td width="5%" bgcolor="#FFFFFF" align="center"><input type="text" name="fkod" size="3" maxlength="3" onFocus="nextfield='fketer';" ></td>
<td width="40%" bgcolor="#FFFFFF" align="center"><input type="text" name="fketer" size="66" maxlength="65" onFocus="nextfield='B1';" ></td>
<td width="15%" bgcolor="#FFFFFF" align="center">
<input type="submit" value="+" name="B1" onFocus="nextfield='done';"><input type="submit" value="res" name="B2">
</td>
</tr>
</table>

<%	end if
	if not objRs2.eof then %>
	
<table width="80%" align="center" border="0"  bgcolor="#330000" cellspacing="1">
<tr width="67%"  bgcolor="#9D2024"> 
<td width="5%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Kod</b></font></td>
<td width="52%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Keterangan</b></font></td>
<td width="5%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Edit</b></font></td>
<td width="5%" bgcolor="#9D2024" align="center">
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Hapus</b></font></td>
</tr>
<%
	bil = 0
	do while not objRs2.eof
	bil = bil + 1
%>
<tr width="67%"> 
<td width="5%" bgcolor="lightgrey" align="center">
<font face="Verdana" size="2"><%=objRs2("kod")%></font></td>
<td width="52%" bgcolor="lightgrey" align="left">
<font face="Verdana" size="2"><%=objRs2("keterangan")%></font></td>
<td width="5%" bgcolor="lightgrey" align="center">
<font face="Verdana" size="2">
<input type="submit"  name="b2<%=bil%>" value="e" >
</font>
<td width="5%" bgcolor="lightgrey" align="center">
<font face="Verdana" size="2">
<input type="submit" onClick="return confirm('Hapus Satu Data ?')" name="b1<%=bil%>" value="-" >
<input type="hidden" name="rowid<%=bil%>" value="<%=objRs2("rowid")%>" >
<input type="hidden" name="kod<%=bil%>" value="<%=objRs2("kod")%>" >
<input type="hidden" name="keter<%=bil%>" value="<%=objRs2("keterangan")%>" ></font></td>
</tr>
<%	objRs2.movenext
	loop					%>
</table>
<input type="hidden" name="bilrec" value="<%=bil%>" >
<%		end if		%>
</form>

</td>
</tr>
</table>
</body>
</html>
