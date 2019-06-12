<!-- '#INCLUDE FILE="ik.asp" -->
<!-- #INCLUDE FILE="adovbs.inc" -->
<%Response.Buffer = True%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>i-Kompaun : Kompaun Mengikut Daerah</title>
<script language="JavaScript">
	function showpage(form)
	{ var item = form.perkara.selectedIndex; choice = form.perkara.options[item].value; if (choice!="x") top.location.href=""+(choice); };
</script>
<script language="JavaScript">
	function page(form)
	{ var item = form.rujukan.selectedIndex; choice = form.rujukan.options[item].value; if (choice!="x") top.location.href=""+(choice); };
</script>
<script language="javascript">
function invalid_data(a)
    {  
       alert (a+" Tiada Maklumat ");
		return(true);
    }
function invalid_tarikh(b)
    {  
       alert (b+" Tarikh Salah !!! ");
		return(true);
    }
</script>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<form method = "Post" action= "ha21216c.asp">
<table border="0" width="100%" height="21">
  <tr> 
    <td width="100%" height="21"></td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#003366">
    <tr valign="top"> 
      <td width="39%"> 
	  
	  <%	'Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"  %> <%	
		tkh1a = Request.querystring("tkh1")
		tkh2a = Request.querystring("tkh2")
		proses="Cari"
	
t1=mid(tkh1a,1,2)+cstr("/")+mid(tkh1a,3,2)+cstr("/")+mid(tkh1a,5,4)
t2=mid(tkh2a,1,2)+cstr("/")+mid(tkh2a,3,2)+cstr("/")+mid(tkh2a,5,4)
	
	
%> 
    <tr bgcolor="#FFFFFF">
      <td colspan="2">
        <div align="center">
          <p>&nbsp;</p>
          <p><b><font face="Verdana" size="2">MAJLIS PERBANDARAN SEBERANG PERAI</font></b></p>
        </div>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> 
        <div align="center"><b><font face="Verdana" size="2">Laporan Bilangan 
          Kompaun Mengikut Daerah</font></b></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" align="center"> 
        <div align="center"></div>
        <div align="center"><b><font face="Verdana" size="2">Tarikh: <%=t1%> Hingga 
          <%=t2%></font></b> </div>
        <p>&nbsp;</p>
      </td>
    </tr>
    <%	
	
	if proses = "Cari" then   		
%> 
  </table>
      <%		b = " select to_date('"&tkh1a&"','ddmmyyyy') as tkha,"
		b = b & " to_date('"&tkh2a&"','ddmmyyyy') as tkhb from dual "
		Set objRsb = Server.CreateObject ("ADODB.Recordset")
		Set objRsb = objConn.Execute(b)
		
		If objRsb.eof then
        	
      		response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_tarikh(""  "");',1) "
       	response.write "</script>"
       
       else
       	tkha = objRsb("tkha")
			tkhb = objRsb("tkhb") 
			
       	if tkhb < tkha then
       	response.write "<script language=""javascript"">"
       	response.write "var timeID = setTimeout('invalid_tarikh(""  "");',1) "
       	response.write "</script>"
       	
       	else
		
		
		
		d = " select count(1) rekod,substr(no_akaun,1,6) aa1, "
		d = d & " decode(substr(no_akaun,1,6),764101,'SPU 76410',764102,'SPT 76410',764103,'SPS 76410',764201,'SPU 76420',764202,'SPT 76420',764203,'SPS 76420',null) dae "
	d = d & " from hasil.bil"
	d = d & " where (no_akaun like '76410%' or no_akaun like '76420%') "
	d = d & " and tkh_masuk between to_date('"&tkh1a&"','ddmmyyyy') and to_date('"&tkh2a&"','ddmmyyyy') "
	d = d & " and (perkara <>'P01' or perkara is null) "
	d = d & " and substr(no_akaun,6,1) in ('1','2','3') "
	 	'******************************************************************
		'ika tambah user view jabatan masing2.admin view semua (23/09/2016)
		pekz = request.cookies("gnop")
		admin = "select id from hasil.superadmin where id='"&pekz&"' "
		'response.Write(admin)
		Set objRAdmin = objConn.Execute(admin)
		
		if objRAdmin.eof then
		
		lokasi = "select lokasi from payroll.paymas where no_pekerja='"&pekz&"' "
		Set objRLokasi = objConn.Execute(lokasi)
		
		lok = objRLokasi("lokasi")
		
		d = d & " and jabatan = '"& lok &"' "
		
		end if
		'end view ikut jabatan
		'******************************************************************
	d = d & " group by substr(no_akaun,1,6) "
	d = d & " order by substr(no_akaun,1,6) "
		Set objRs2 = objConn.Execute(d)
	
		if not objRs2.eof then
%>
  <table width="70%" border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all align="center"
style="border-collapse: collapse; border: 1px solid black">
    <tr bgcolor="#FFFFFF"> 
      <td align="center"><b><font size="2" face="Verdana" color="#000000">&nbsp;Daerah&nbsp;- 
        Kompaun </font></b></td>
      <td align="center" nowrap><b><font size="2" face="Verdana" color="#000000">Bilangan 
        Kompaun </font></b></td>
      <td align="center" nowrap><b><font size="2" face="Verdana" color="#000000">Bayaran</font></b></td>
 <td align="Center" nowrap ><b><font size="2" face="Verdana" color="#000000">Bilangan Bayar</font></b></td>
     
</tr>
    <tr bgcolor="#FFFFFF"> 
      <td  align="center"><font size="2" face="Verdana">&nbsp;</font></td>
      <td  align="center"> 
        <div align="center"><font size="2" face="Verdana">&nbsp;</font></div>
      </td>
      <td  align="center">&nbsp;</td>
    </tr>
    <% 
   		bil = 0
   		belum = 0
   			
    	Do while not objRs2.EOF
    	
    	rekod = objRs2("rekod")
    	dae = objRs2("dae")
aa1 = objRs2("aa1")
'response.write tkh2a
bb="select count(1) kir ,sum(nvl(amaun_bayar,0)) amb from hasil.bil where no_akaun like  '"&aa1&"'||'%' "
bb = bb & " and amaun_bayar is not null and tkh_masuk between to_date('"&tkh1a&"','ddmmyyyy') and to_date('"&tkh2a&"','ddmmyyyy') "
Set objbb = objConn.Execute(bb)
amb = objbb("amb")
kirbaya=objbb("kir")

'response.write "s" & amb & "s" 
if isnull(objbb("amb")) then amb="0"
'response.write "s" & amb & "s"     	
    
    
    	bil = bil + 1
 %> 
    <tr bgcolor="#FFFFFF"> 
      <td  align="center" nowrap><font size="2" face="Verdana">&nbsp;<%=dae%>&nbsp;</font></td>
      <td  align="center"> 
        <div align="center"><font size="2" face="Verdana"><%=rekod%></font></div>
      </td>
      <td  width="80"  align="right"><font size="2" face="Verdana"><%=formatnumber(amb,2)%></font></td>
<td  width="80"  align="right"><font size="2" face="Verdana"><%=kirbaya%></font></td>
 
    </tr>
    <%
  	objRs2.MoveNext	
if bil=3 then%> 
    <tr bgcolor="#FFFFFF"> 
      <td width="74" align="center"><font size="2" face="Verdana">&nbsp;</font></td>
      <td width="121" align="center"> 
        <div align="center"><font size="2" face="Verdana">&nbsp;</font></div>
      </td>
      <td width="121" align="center">&nbsp;</td>
    </tr>
    <% end if
Loop
%> 
  </table>

<%	
 	else  
 	
 	response.write "<script language=""javascript"">"
    response.write "var timeID = setTimeout('invalid_data(""  "");',1) "
    response.write "</script>"
	 
 		end if
 		end if
 		end if
		end if
%> 
</tr>
</table>
</form>
</body>
