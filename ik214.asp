<%response.cookies("ikmenu") = "ik214.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Kompaun Mengikut Daerah</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

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
<form method = "Post" action= "ik214.asp">
      <%	'Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"  %>
     
	
  <%	
		tkh1 = Request.form("tkhdari")
		tkh2 = Request.form("tkhhingga")
	
	t = " select '01'||to_char(sysdate,'mm')||to_char(sysdate,'yyyy') tkh1, to_char(sysdate,'ddmmyyyy') tkh2 from dual "
	Set objRst = Server.CreateObject ("ADODB.Recordset")
   	Set objRst = objConn.Execute(t)
	
	if tkh1 = "" then
	tkh1 = objRst("tkh1")
	tkh2 = objRst("tkh2")
	end if

	
	
%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
<td class="hd">Tarikh</td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" size="8" maxlength="8">
        &nbsp; Hingga &nbsp; 
        <input type="text" name="tkhhingga" value="<%=tkh2%>" size="8" maxlength="8"> &nbsp;
      <input type="submit" value="Cari" name="B1" class="button">&nbsp;
        <input name="B1" type="submit" id="B1" class="button" value="Cetak">
        <input type="hidden" name="kode" value="<%=vkod%>" > <input type="hidden" name="dkod" value="<%=perk1%>" > 
    
    </td>
  </tr>
  <%	
	proses = Request.form("B1")
	ko21 = Request.form("kod1")
	dkod = Request.form("dkod")
	kod11 = Request.QueryString("rujuk")  
	kod22 = Request.QueryString("rujuk1")
			
	 if proses = "Cetak" then
		 response.redirect "ik214c.asp?tkh1="&tkh1&"&tkh2="&tkh2&""
    end if
	
	if proses = "Cari" then   		
%>
 
</table>
      <%		b = " select to_date('"&tkh1&"','ddmmyyyy') as tkha,"
		b = b & " to_date('"&tkh2&"','ddmmyyyy') as tkhb from dual "
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
		
		
		
		d = " select count(1) rekod,substr(no_akaun,1,6) aa, "
		d = d & " decode(substr(no_akaun,1,6),764101,'SPU 76410',764102,'SPT 76410',764103,'SPS 76410',764201,'SPU 76420',764202,'SPT 76420',764203,'SPS 76420',null) dae "
		d = d & " from hasil.bil"
		d = d & " where (no_akaun like '76410%' or no_akaun like '76420%') "
		d = d & " and tkh_masuk between to_date('"&tkh1&"','ddmmyyyy') and to_date('"&tkh2&"','ddmmyyyy') "
		d = d & " and (perkara <>'P01' or perkara is null) "
		d = d & " and substr(no_akaun,6,1) in ('1','2','3') "
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
		d = d & " group by substr(no_akaun,1,6)  "
		d = d & " order by substr(no_akaun,1,6)  "
		Set objRs2 = objConn.Execute(d)
	
		if not objRs2.eof then
%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="25%" class="hd1">Daerah</td>
<td width="25%" class="hd1">Bilangan Kompaun</td>
<td width="25%" class="hd1">Bayaran</td>
<td width="25%" class="hd1">Bilangan Bayar</td>
    </tr>
    <% 
   		bil = 0
   		belum = 0
   			
    	Do while not objRs2.EOF
    	
    	rekod = objRs2("rekod")
    	dae = objRs2("dae")
aa = objRs2("aa")
'response.write "s" & aa & "s"   
bb="select count(1) kir ,sum(nvl(amaun_bayar,0)) amb from hasil.bil where no_akaun like  '"&aa&"'||'%' "
bb = bb & " and amaun_bayar is not null and tkh_masuk between to_date('"&tkh1&"','ddmmyyyy') and to_date('"&tkh2&"','ddmmyyyy') "
Set objbb = objConn.Execute(bb)
amb = objbb("amb")
kirbaya=objbb("kir")

'response.write "s" & amb & "s" 
if isnull(objbb("amb")) then amb="0"
'response.write "s" & amb & "s"     	
    
    	bil = bil + 1
		ab = objbb("amb")
		total_ab = cdbl(total_ab) + cdbl(ab)
 %>
    <tr align="center"> 
      <td><%=dae%></td>
      <td><%=rekod%></td>
<td align="right"><%=formatnumber(amb,2)%></td>
<td><%=kirbaya%></td>
    </tr>
    <%
  	objRs2.MoveNext			
  	Loop
%>
 <tr>
 <td colspan="2" align="right"><b>Jumlah</b></td>
 <td align="right"><b>RM <%=FormatNumber(total_ab,2)%></b></td></tr>
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