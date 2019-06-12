<%response.cookies("ikmenu") = "ik215.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Statistik Kompaun</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<STYLE TYPE="text/css" MEDIA="screen">
</STYLE>

<form name=su31 method="POST" action="ik215.asp">
  <%
  
'  Set objConn = Server.CreateObject("ADODB.Connection")
'  objconn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  
  thn = request.form("thn")
  rosak = request.form("rosak")
  bhantar = request.form("bhantar")
  breset = request.form("breset")
  proses = request.form("bcetak")
  
  if bhantar="" and breset="" then
  	mula
  end if
 if proses = "Cetak" then
   thn = request.form("thn")
  rosak = request.form("rosak")
		 response.redirect "ik215c.asp?tkh1="&thn&"&tkh2="&rosak&""
    end if
  
  if bhantar="Hantar" then
  	mula
  	hantar
  end if
  
  if breset="Reset" then
	rosak= ""
	thn= ""
	mula
  end if  

  sub mula
ss="select to_char(sysdate,'yyyy') s from dual "
set Rs=objconn.execute(ss)
thnini=Rs("s")
  
%>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr align="center"> 
      <td colspan="2" class="hd" >STATISTIK 
          KOMPAUN
      </td>
    </tr>
    <tr> 
      <td class="hd">
        Bayaran</td>
      <td> 
        <input type="radio" name="rosak" value="1" <%if rosak="1" then%>checked<%end if%>>
         Bayar 
        <input type="radio" name="rosak" value="0" <%if rosak="0" then%>checked<%end if%>>
        Tak bayar</td>
    </tr>
    <tr> 
      <td class="hd">Tahun</td>
      <td> 
        <select size="1" name="thn" onFocus="nextfield='bhantar';">
          <option value="" selected>[Semua 
          Tahun]</option>
          <% 	if thn <> "" then
				
	    		d = "select  distinct(tahun)tahun from hasil.kesihatan_view where   "
			    d = d & "     tahun='"&thn&"' order by tahun desc"
				Set rsd= Server.CreateObject("ADODB.Recordset")
    			Set rsd = objConn.Execute(d)
	  %>
          <option selected value="<%=rsd("tahun")%>"><%=rsd("tahun")%></option>
          <%end if
			   
				d1 = "select distinct(tahun)tahun from hasil.kesihatan_view     "
				d1 = d1 & "   order by tahun desc"
				Set rsd1= Server.CreateObject("ADODB.Recordset")
    			Set rsd1 = objConn.Execute(d1)
				Do while not rsd1.EOF				
	  %>
          <option value="<%=rsd1("tahun")%>"><%=rsd1("tahun")%></option>
          <% 	rsd1.MoveNext
				Loop			
	  %>
        </select>
        <input type="submit" value="Hantar" name="bhantar" onFocus="nextfield='done';" class="button">
        <input type="submit" value="Reset" name="breset" onFocus="nextfield='done';" class="button">
        <input type="submit" value="Cetak" name="bcetak" onFocus="nextfield='done';" class="button">
        </td>
    </tr>
  </table>
  <%end sub%>
   <%sub hantar
  
    ss= "select  tahun,spt,spu,sps,jspt,jspu,jsps "
	'ss= ss & " sum(decode(substr(tempat,1,2),'BM',1,0))BM,"
 '	ss = ss & "sum(decode(substr(tempat,1,2),'NT',1,0))NT"
	ss= ss & " from hasil.kesihatan_view  "
	if thn <> "" then 
	ss = ss & " where  tahun='"&thn&"' and "
	else
	ss = ss & " where    "
	end if
 	ss= ss & "   bayar='"&rosak&"' "
         ' ss= ss & "   and tahun between '1990' and '"&thnini&"' "
	'ss= ss & " group by kekerapan order by kekerapan desc"
	set rsss = objconn.execute(ss)%>
<br> <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
    <tr align="center"> 
<td width="20%" class="hd1" rowspan="2">Tahun</td>
<td width="20%" class="hd1" colspan="2">SPU</td>
<td width="20%" class="hd1" colspan="2">SPT</td>
<td width="20%" class="hd1" colspan="2">SPS</td>
<td width="20%" class="hd1" colspan="2">Jumlah</td>
    </tr>
    
    <tr align="center"> 
      <td align="center" class="hd1">Bil</td>
      <%'if rosak="1" then%>
        <td align="center" class="hd1">Amaun</td>
      <%'end if%>
      <td align="center" class="hd1">Bil</td>
      <td align="center" class="hd1">Amaun</td>
      <td align="center" class="hd1">Bil</td>
      <td align="center" class="hd1">Amaun</td>
      <td align="center" class="hd1">Bil</td>
      <td align="center" class="hd1">Amaun</td>
    </tr>
    <%
	ctrz = 0
   	bilrec = 0
	jum=0	
	 Do While Not rsss.eof
	if not rsss.eof then 
		 
		bw=rsss("spu")
		bm=rsss("spt")
		nt=rsss("sps")
		jbw=rsss("jspu")
		jbm=rsss("jspt")
		jnt=rsss("jsps")
		 
		tahun=rsss("tahun")
  	end if
	
	jum=cdbl(bm)+cdbl(bw)+cdbl(nt)
	jumbm=cdbl(jumbm)+cdbl(bm)
	jumbw=cdbl(jumbw)+cdbl(bw)
	jumnt=cdbl(jumnt)+cdbl(nt)
	tahunrm=cdbl(jbm)+cdbl(jnt)+cdbl(jbw)
    ajumbm=cdbl(ajumbm)+cdbl(jbm)
	ajumbw=cdbl(ajumbw)+cdbl(jbw)
	ajumnt=cdbl(ajumnt)+cdbl(jnt)

%>
    <tr align="center"> 
      <td><%=tahun%></td>
      <td><%=bw%></td>
      <td>RM <%=formatnumber(jbw,2)%></td>
      <%'else%>
      <td><%=bm%></td>
      <td>RM <%=formatnumber(jbm,2)%></td>
      <td><%=nt%></td>
      <td>RM <%=formatnumber(jnt,2)%></td>
      <%'else%>
      <td><%=jum%></td>
      <td>RM <%=formatnumber(tahunrm,2)%></td>
      <%'else%>
    </tr>
    <%
  rsss.movenext
	loop
	jumsum=cdbl(jumbm)+cdbl(jumbw)+cdbl(jumnt)
	ajum =cdbl(ajumbm)+cdbl(ajumbw)+cdbl(ajumnt)
	%>
    <tr align="center"> 
      <td><b>Jumlah</b></td>
      <td><b><%=jumbw%></b></td>
      <td><b>RM <%=formatnumber(ajumbw,2)%></b></td>
      <%'else%>
      <td><b><%=jumbm%></b></td>
      <td><b>RM <%=formatnumber(ajumbm,2)%></b></td>
      <td><b><%=jumnt%></b></td>
      <td><b>RM <%=formatnumber(ajumnt,2)%></b></td>
      <td><b><%=jumsum%></b></td>
     <td><b>RM <%=formatnumber(ajum,2)%></b></td>
      <%'else%>
      <%'else%>
    </tr>
  </table>
  <input type="hidden" value="<%=ctrz%>" name="bilrec">

  <%end sub %>
</form>
</body>

