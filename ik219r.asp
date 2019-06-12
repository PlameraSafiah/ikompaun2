
<%response.cookies("ikmenu") = "ik219.asp"%>
<%Response.Buffer = True%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>i-Kompaun : Laporan Kompaun Belum Jelas (Nama)</title>
</head>
<script language="javascript">
print();
</script>
<style>
TABLE.mailer {BORDER-COLLAPSE: collapse; FONT-SIZE:11pt; FONT-FAMILY:Arial; WIDTH: 99%;
page-break-after:always;}

</style>

<%


	Set objConn = Server.CreateObject("ADODB.Connection")
   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
	  pekz = session("pekz1")
%>

<body>

<% 


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
		a = a & " nama is not null and nama <> '0' and nama <> '-'"
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

 <table width="80%" cellspacing="0" cellpadding="1" border="1" frame="box" rules="all" align="center" style="font-family: Calibri; font-size: 10pt;">
    <tr align="center"> 
    <td width="10%" class="hd1">Bil</td>
<td width="40%" class="hd1">Bil Kompaun</td>
<td width="40%" class="hd1" align="left">Nama </td>
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
<td><%=objRa("nama")%></td>
<tr>

<%
objRa.movenext
loop
%>

</table>
  
  </td>
</tr>
</table>



</body>
</html>
