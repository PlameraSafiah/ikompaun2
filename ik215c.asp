 
<%response.buffer=true%>

<head>
<STYLE TYPE="text/css" MEDIA="screen">
</STYLE>

<title>i-Kompaun : Statistik Kompaun</title></head>

<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color2%>" >
<form name=su3a1 method="POST" action="ik215c.asp">
  <%Set objConn = Server.CreateObject("ADODB.Connection")
  objconn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  
  thn = request.querystring("tkh1")
  rosak = request.querystring("tkh2")
  

 mula
hantar

  sub mula
ss="select to_char(sysdate,'dd/mm/yyyy') s from dual "
set Rs=objconn.execute(ss)
thnini=Rs("s")
  
%>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" bgcolor="#FFFFFF" >&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" bgcolor="#FFFFFF" >&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" bgcolor="#FFFFFF" > 
        <div align="center"><b>MAJLIS PERBANDARAN SEBERANG PERAI</b></div>
      </td>
    </tr>
    <tr bgcolor="#CCCCCC"> 
      <td colspan="2" bgcolor="#FFFFFF"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b>STATISTIK 
          KOMPAUN KESIHATAN</b></font></div>
      </td>
    </tr>
    <tr bgcolor="#CCCCCC"> 
      <td colspan="2" bgcolor="#FFFFFF"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b>YANG<%IF rosak="1" then%> 
          TELAH DIBAYAR<%else%> BELUM DIBAYAR<%end if%></b></font></div>
      </td>
    </tr>
    <tr bgcolor="#CCCCCC"> 
      <td colspan="2" bgcolor="#FFFFFF"> 
        <div align="center"><b>Sehingga <%=thnini%></b></div>
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
  <table width="99%" height="77" border="1" align = "center" cellpadding="1" cellspacing="0" bordercolor="#000000" bgcolor="#CCCCCC" style="border-collapse: collapse; border: 1px solid black" >
    <tr bgcolor="#FFFFFF" valign="top"> 
      <td width="16%" height="9" rowspan="2" align="center"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Tahun</font></b></font></td>
      <td height="4" colspan="2" align="center"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">SPU</font></b></font></td>
      <td height="4" colspan="2" align="center"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">SPT</font></b></font></td>
      <td height="4" colspan="2" align="center"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">SPS</font></b></font></td>
      <td height="4" colspan="2" align="center"><font size="2" color="#000000"><b><font face="Arial, Helvetica, sans-serif">Jumlah</font></b></font></td>
    </tr>
    <tr bgcolor="#990033" valign="top"> 
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bil</font></b></font></td>
      <%'if rosak="1" then%> 
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Amaun 
        Bayar </font></b></font></td>
      <%'end if%> 
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bil</font></b></font></td>
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Amaun 
        </font><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bayar</font></b></font><font size="2" face="Arial, Helvetica, sans-serif"> 
        </font></b></font></td>
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bil</font></b></font></td>
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Amaun 
        </font><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bayar</font></b></font><font size="2" face="Arial, Helvetica, sans-serif"> 
        </font></b></font></td>
      <td  height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Bil</font></b></font></td>
      <td   height="4" align="center" bgcolor="#FFFFFF"><font color="#000000"><b><font size="2" face="Arial, Helvetica, sans-serif">Amaun 
        </font></b></font></td>
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
    <tr bgcolor="#FFFFFF"> 
      <td width="16%" height="14" onMouseOver="this.style.backgroundColor='#ffffff'" > 
        <div align="center"><font size="2" color="#000000" face="Arial, Helvetica, sans-serif"> 
          <%=tahun%> </font></div>
      </td>
      <td height="14" onMouseOver="this.style.backgroundColor='#ffffff'"  > 
        <div align="center"><font size="2" color="#0033FF" face="Arial, Helvetica, sans-serif"> 
          &nbsp;<%=bw%> </font></div>
      </td>
      <td onMouseOver="this.style.backgroundColor='#ffffff'" > 
        <div align="left"><b></b></div>
        <div align="left"><font color="#000000"> <font size="2" face="Arial, Helvetica, sans-serif"><font size="2" face="Arial, Helvetica, sans-serif">RM<b>&nbsp;</b></font><%=formatnumber(jbw,2)%></font> 
          </font></div>
      </td>
      <%'else%> 
      <td height="14" onMouseOver="this.style.backgroundColor='#ffffff'" > 
        <div align="center"><font size="2" color="#0033FF" face="Arial, Helvetica, sans-serif"> 
          &nbsp;<%=bm%></font></div>
      </td>
      <td onMouseOver="this.style.backgroundColor='#ffffff'"  > 
        <div align="left"><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><font size="2" face="Arial, Helvetica, sans-serif">RM<b>&nbsp;</b></font></font></font><font size="2" color="#000000" face="Arial, Helvetica, sans-serif"><%=formatnumber(jbm,2)%></font></div>
      </td>
      <td height="14" onMouseOver="this.style.backgroundColor='#ffffff'"  > 
        <div align="center"><font size="2" color="#0033FF" face="Arial, Helvetica, sans-serif"> 
          <%=nt%> </font> </div>
      </td>
      <td onMouseOver="this.style.backgroundColor='#ffffff'" > 
        <div align="left"></div>
        <div align="left"><font color="#000000"> <font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><font size="2" face="Arial, Helvetica, sans-serif">RM<b>&nbsp;</b></font></font></font><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(jnt,2)%></font></font></div>
      </td>
      <%'else%> 
      <td width="6%" height="14" onMouseOver="this.style.backgroundColor='#ffffff'" > 
        <div align="center"><font size="2" color="#0033FF" face="Arial, Helvetica, sans-serif"><%=jum%> 
          </font> </div>
      </td>
      <td height="14" colspan="2" onMouseOver="this.style.backgroundColor='#ffffff'" ><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif">RM&nbsp;<%=formatnumber(tahunrm,2)%></font></font></td>
      <%'else%> </tr>
    <%
  rsss.movenext
	loop
	jumsum=cdbl(jumbm)+cdbl(jumbw)+cdbl(jumnt)
	ajum =cdbl(ajumbm)+cdbl(ajumbw)+cdbl(ajumnt)
	%> 
    <tr bgcolor="#FFFFFF"> 
      <td height="69"> 
        <div align="center"><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif">Jumlah</font></font></strong></div>
      </td>
      <td height="69" onMouseOver="this.style.backgroundColor='#ffffff'" onMouseOut="this.style.backgroundColor='#ffffff'"> 
        <div align="center"><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><%=jumbw%></font></font></strong></div>
      </td>
      <td height="69" > 
        <div align="right"><strong><font color="#000000">RM&nbsp;&nbsp;<font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(ajumbw,2)%></font></font></strong></div>
      </td>
      <%'else%> 
      <td height="69" > 
        <div align="center"><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><%=jumbm%></font></font></strong></div>
      </td>
      <td onMouseOver="this.style.backgroundColor='#ffffff'" onMouseOut="this.style.backgroundColor='#ffffff'" height="69"> 
        <div align="right"><strong><font color="#000000">RM&nbsp;&nbsp;<font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(ajumbm,2)%></font></font></strong></div>
      </td>
      <td height="69" onMouseOver="this.style.backgroundColor='#ffffff'" onMouseOut="this.style.backgroundColor='#ffffff'"> 
        <div align="center"><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><%=jumnt%></font></font></strong></div>
      </td>
      <td height="69" > 
        <div align="right"><strong><font color="#000000">RM&nbsp;<font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(ajumnt,2)%></font></font></strong></div>
      </td>
      <td height="69" > 
        <div align="center"><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif"><%=jumsum%></font></font></strong></div>
      </td>
      <td height="69" colspan="3" ><strong><font color="#000000"><font size="2" face="Arial, Helvetica, sans-serif">RM&nbsp;<%=formatnumber(ajum,2)%></font></font></strong> 
        <div align="center"></div>
      </td>
      <%'else%> <%'else%> </tr>
  </table>
  <input type="hidden" value="<%=ctrz%>" name="bilrec">

  <%end sub %>
</form>
</body>

