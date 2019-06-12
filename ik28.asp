<%response.cookies("ikmenu") = "ik28.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Mengikut Tred</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<script language="javascript">
function invalid_data(a)
  {  
       alert (a+" Tiada Rekod ");
		return(true);
  }
function invalid_tred(b)
    {  
       alert (b+" Sila Pilih Jenis Tred !!! ");
		return(true);
    }
function invalid_tkh(c)
    {  
       alert (c+" Sila Masukkan Tarikh !!! ");
		return(true);
    }
function invalid_tarikh(d)
    {  
       alert (d+" Tarikh Salah !!! ");
		return(true);
    }
</script>

</head>
<form method="POST" action="ik28.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr valign="top"> 
<td width="100%">


<%
'	Set objConn = Server.CreateObject("ADODB.Connection")
'   objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
   	
   	proses = Request.form("B1")
   		
	if proses = "Cari" then
		tred = Request.form("ftred")
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")
	else
		e = " select '01'||to_char(sysdate,'mm')||to_char(sysdate,'yyyy') as tkh1s,to_char(sysdate,'ddmmyyyy') as tkh2s from dual "
		Set objRse = Server.CreateObject ("ADODB.Recordset")
   		Set objRse = objConn.Execute(e)	
   		tkh1 = objRse("tkh1s")
   		tkh2 = objRse("tkh2s")	
	end if

'response.Write"<p>tred"&tred&"</p>"	
'response.Write"<p>tkh1"&tkh1&"</p>"	
'response.Write"<p>tred"&tkh2&"</p>"	
	if proses = "Cetak" then
		tred = Request.form("ftred")
		tkh1 = Request.form("tkhdari")	
		tkh2 = Request.form("tkhhingga")

     response.cookies("tkh1") = tkh1
	 response.cookies("tkh2") = tkh2
	 response.cookies("tred") = tred
     response.redirect "ik28r.asp"
   
    end if
	
	dtred = Request.QueryString("dtred")
	
	if dtred <> "" then
		tred = Request.QueryString("dtred")
		tkh1 = Request.QueryString("dtkh1")	
		tkh2 = Request.QueryString("dtkh2")	
	end if
	
	
		r = " select kod, initcap(keterangan) terang from lesen.tred "
   		Set objRsr = Server.CreateObject ("ADODB.Recordset")
   		Set objRsr = objConn.Execute(r)	
   		
   		
   		rr = " select initcap(keterangan) terangz from lesen.tred "
   		rr = rr & " where kod = '"& tred &"' "
   		Set objRsrr = Server.CreateObject ("ADODB.Recordset")
   		Set objRsrr = objConn.Execute(rr)
 %>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
    <td class="hd">Jenis Tred</td>
<td><select size="1" name="ftred" value="<%%>">
<%	if tred = "" then		%>
<option value="">Sila Buat Pilihan </option>

<%	else	%>

<option value="<%=tred%>" ><%=tred%> - <%=objRsrr("terangz")%> </option>
<%		end if
     	do while not objRsr.EOF 
%> 
<option value="<%=objRsr("kod")%>"><%=objRsr("kod")%> - <%=objRsr("terang")%></option>
<%
     objRsr.MoveNext
     loop
%> 
</select></td>
</tr>
<tr>
<td class="hd">Tarikh Dari</td>
<td><input type="text" name="tkhdari" value="<%=tkh1%>" size="8" maxlength="8">&nbsp;&nbsp;
&nbsp;&nbsp;Hingga&nbsp;
<input type="text" name="tkhhingga" value="<%=tkh2%>" size="8" maxlength="8">&nbsp;
<input type="submit" value="Cari" name="B1" class="button">

</td>
 </tr></table> 
<%	

	Dim iPageSize,iPageCount,iPageCurrent,iRecordsShown
	Dim S
	iPageSize = 50

	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If



	if proses = "Cari" or tred <> "" or tkh1 <> "" then
	
		if proses = "Cari" and tred = ""	then
				response.write "<script language=""javascript"">"
       		response.write "var timeID = setTimeout('invalid_tred(""  "");',1) "
       		response.write "</script>"
       		proses = "Cari"
		else
		if tkh1 = ""	or tkh2 = "" then
				response.write "<script language=""javascript"">"
       		response.write "var timeID = setTimeout('invalid_tkh(""  "");',1) "
       		response.write "</script>"
       		proses = "Cari"		
		else
		
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
		
		d = " select rowid, no_akaun,perkara, no_rujukan2, nama, to_char(tkh_masuk,'ddmmyyyy') tkh_masuk, "
		d = d & " nvl(amaun_bayar,0) amaun, to_char(tkh_bayar,'ddmmyyyy') tkh_bayar, no_resit,jabatan "
		d = d & " from hasil.bil "
		d = d & " where tred = '"& tred &"' "
		d = d & " and tkh_masuk between to_date('"& tkh1 &"', 'ddmmyyyy') "
		d = d & " and to_date('"& tkh2 &"' , 'ddmmyyyy') "
		d = d & " and (no_akaun like '76410'||'%' or no_akaun like '76420'||'%' or no_akaun like '76413'||'%' "
		''kompaun perbandaran
		d = d & " or no_akaun like '76412'||'%' or no_akaun like '76415'||'%' "
		'kompaun bangunan & veterinar
		d = d & " or no_akaun like '76101'||'%'or no_akaun like '76441'||'%') "
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
		d = d & " order by no_akaun "
	
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

%>  
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
          <tr align="right"> 
            <td align="left" colspan=3 >Jumlah 
              Rekod : <%=kira%></td>
            <td colspan=8>
              <% If iPageCurrent <> 1 Then %>
              <a href="ik28.asp?page=1&bilangan=0&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dtred=<%=tred%>&proses=Cari"> 
              <img name="firstrec" border="0" src="images/firstrec.jpg" width="20" height="20" alt="Halaman Mula"></a> 
              <% End If %>
              <% If iPageCurrent <> 1 Then%>
              <a href="ik28.asp?page=<%= iPageCurrent - 1 %>&bilangan=<%=bil-count-iPageSize%>&ms=pre&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dtred=<%=tred%>&proses=Cari"> 
              <img name="previous" border="0" src="images/previous.jpg" width="20" height="20" alt="Rekod Sebelum"></a> 
              <% End If %>
              Halaman <%=iPageCurrent%>/ 
              <%if iPageCount=0 then%>
              1 
              <%else%>
              <%=iPageCount%> 
              <%end if%>
              <% If iPageCurrent < iPageCount Then	%>
              <a href="ik28.asp?page=<%= iPageCurrent + 1 %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dtred=<%=tred%>&proses=Cari"> 
              <img name="next" border="0" src="images/next.jpg" width="20" height="20" alt="Rekod Seterusnya"></a> 
              <% End If 
	  If iPageCurrent < iPageCount Then
	  bil = (iPageCount - 1) * iPageSize %>
              <a href="ik28.asp?page=<%=iPageCount %>&bilangan=<%=bil%>&ms=next&dtkh1=<%=tkh1%>&dtkh2=<%=tkh2%>&dtred=<%=tred%>&proses=Cari"> 
              <img name="lastrec" border="0" src="images/lastrec.jpg" width="20" height="20" alt="Halaman Akhir"></a> 
              <% End If %>
              </td>
          </tr></table>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
			<td width="5%" class="hd1">Bil</font></b></td>
            <td width="12%" class="hd1">No Kompaun</td>
            <td width="6%" class="hd1">Akta</td>
            <td width="8%" class="hd1">Jenis</td>
            <td width="14%" class="hd1">Nama</td>
            <td width="10%"  class="hd1">Tarikh&nbsp;</td>
            <td width="12%" class="hd1">Amaun</td>
            <td width="10%" class="hd1">Tkh Bayar&nbsp;</td>
            <td width="9%" class="hd1">No Resit</td>
            <td width="9%" class="hd1">Jabatan</td>
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
		
		kodJbtn = objRs2("jabatan")
		
		q1="select keterangan from payroll.ptj where kod='"&kodJbtn&"'"
		set rq1 = objConn.execute(q1)
		
		if not rq1.eof then namaJabatan = rq1("keterangan")
		
%>
          <tr align="center"> 
            <td><%=bil%></td>
            <td><%=objRs2("no_akaun")%></td>
            <%	kara = objRs2("perkara")
  		if kara <> "" then	%>
            <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"><a href="akta2.asp?rujuk=<%=objRs2("perkara")%>"> 
             <%=objRs2("perkara")%></a></td>
            <%	else		%>
            <td><%=objRs2("perkara")%></td>
            <%	end if		%>
            <%	ruj2 = objRs2("no_rujukan2") 
  		if ruj2 <> "" then	%>
            <td onMouseover="this.style.backgroundColor='#CC6666'" onMouseout="this.style.backgroundColor='#FFE9E8'"> 
              <a href="salah2.asp?rujuk=<%=objRs2("perkara")%>&rujuk1=<%=objRs2("no_rujukan2")%>"> 
              <%=objRs2("no_rujukan2")%></a></td>
            <%	else	%>
            	
            <td><%=objRs2("no_rujukan2")%></td>
            <%	end if		%>
           	<td><%=objRs2("nama")%></td>
			<td><%=objRs2("tkh_masuk")%></td>
            <td><%=FormatNumber(objRs2("amaun"),2)%></td>
            <td><%=objRs2("tkh_bayar")%></td>
            <td><%=objRs2("no_resit")%>&nbsp;&nbsp;</td>
            <td><%=namaJabatan%></td>
            </tr>
          <%
  	iRecordsShown = iRecordsShown + 1
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
  		end if
  %>
</td>
 </tr>
 </table>  
 </form>
 <form method="post" action="ik28r.asp"> 
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr>
      <td align="center" class="hd"> 
        <input type="submit" name="B1" value="Cetak" class="button">
<input type="hidden" name="ftkh1" value="<%=tkh1%>">
<input type="hidden" name="ftkh2" value="<%=tkh2%>">
<input type="hidden" name="ftred1" value="<%=tred1%>">

</form>
</body>
</html>