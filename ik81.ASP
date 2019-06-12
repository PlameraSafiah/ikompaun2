<%response.cookies("ikmenu") = "ik81.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<%Response.Buffer = True%>
<html>
<head>
<title>i-Kompaun : Daftar Pengguna</title>
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">


<script language="javascript">
 function invalid_kod(a)
 { alert(a+" Kod Sudah Ada!!! ");
   return(true);
 }
 function keluar(f)
 { location=window.close();
 }
</script>
</head>
<body>
<%response.cookies("amenu") = "ik81.asp"%>

<FORM name="kew" action="ik81.asp" method="post" >

<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"

  proses = Request.Form("B1")
  pekerja = request.form("pekerja")
  sistem = request.form("sistem")
  sistem2 = request.form("sistem2")
  
  mula2
  
  b2 = request.form("b2")
  if b2 = "Hantar" then
     proses = "Hantar2"
  end if
  
  proses2 = Request.Form("siap")
  
  if proses = "Hantar" then    hantar
  
  if proses = "Hantar2" then     senarai

  if proses = "Simpan" then
     bilrec = request.form("bilrec")
         ss = "delete majlis.kebenaran_2002 where no_pekerja = '"& pekerja &"' and sistem = 'ik' "
         Set Rsss = objConn.Execute(ss)
         
  if bilrec <> "" then
     proses="z"
     for i = 1 to bilrec
         rowid = "rowid" + cstr(i)
         kod1 = "kod1" + cstr(i)
         kod2 = "kod2" + cstr(i)
         kod3 = "kod3" + cstr(i)
         kod4 = "kod4" + cstr(i)
         ket1 = "ket1" + cstr(i)
         ket2 = "ket2" + cstr(i)
         ket3 = "ket3" + cstr(i)
         ket4 = "ket4" + cstr(i)

         rowid = request.form(""& rowid &"")
         kod1 = request.form(""& kod1 &"")
         kod2 = request.form(""& kod2 &"")
         kod3 = request.form(""& kod3 &"")
         kod4 = request.form(""& kod4 &"")
         ket1 = request.form(""& ket1 &"")
         ket2 = request.form(""& ket2 &"")
         ket3 = request.form(""& ket3 &"")
         ket4 = request.form(""& ket4 &"")

         if ket1 = "1" then
            i1 = " insert into majlis.kebenaran_2002 (no_pekerja,sistem,skrin) "
            i1 = i1 & " values ('"& pekerja &"','ik','"& kod1 &"') "
            Set Rsi1 = objConn.Execute(i1)
         end if

         if ket2 = "1" then
            i2 = " insert into majlis.kebenaran_2002 (no_pekerja,sistem,skrin) "
            i2 = i2 & " values ('"& pekerja &"','ik','"& kod2 &"') "
            Set Rsi2 = objConn.Execute(i2)
         end if

         if ket3 = "1" then
            i3 = " insert into majlis.kebenaran_2002 (no_pekerja,sistem,skrin) "
            i3 = i3 & " values ('"& pekerja &"','ik','"& kod3 &"') "
            Set Rsi3 = objConn.Execute(i3)
         end if

         if ket4 = "1" then
            i4 = " insert into majlis.kebenaran_2002 (no_pekerja,sistem,skrin) "
            i4 = i4 & " values ('"& pekerja &"','ik','"& kod4 &"') "
            Set Rsi4 = objConn.Execute(i4)
         end if

     next
     
     hantar
  end if
  end if
  
  bilrecsen = request.form("bilrecsen")
  if bilrecsen <> "" then
     proses = "z"
     for i = 1 to bilrecsen
     
         d = "d" + cstr(i)
         e = "e" + cstr(i)
         nopekerja = "nopekerja" + cstr(i)
         nama = "nama" + cstr(i)
         
         d = request.form (""& d &"")
         e = request.form (""& e &"")
         nopekerja = request.form (""& nopekerja &"")
         nama = request.form (""& nama &"")
         
         if d = "Hapus" then
            db = " delete majlis.kebenaran_2002 where no_pekerja = '"& nopekerja &"' "
            set objrsdb = objConn.Execute(db)
            mula3
            senarai
         elseif e = "Edit" then
             pekerja = nopekerja
             sistem = sistem2
             hantar
         end if
     next   
  end if
  
  if proses = "" then
     gmenu = request.cookies("gmenu")
     if gmenu = "daftar.asp" then
        proses = "z"
        sistem = request.cookies("gsistem")
        pekerja = request.querystring("kod")
     else
     end if
     response.cookies("gsistem") = ""
     response.cookies("gmenu") = ""
  end if
 sub mula2 %>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
  <tr> 
  <td class="hd">No Pekerja </td><td>&nbsp;
      <input type="text" name="pekerja" value="<%=pekerja%>" size="5"  onKeyDown="if(event.keyCode==13) event.keyCode=9;">
<%if pekerja <> "" then
  z0 = "select initcap(nama) nama from payroll.tetap_sambilan where no_pekerja = '"& pekerja &"' "
  Set Rsz0 = objConn.Execute(z0) 
  if not rsz0.eof then
     nama = rsz0("nama")
  else
     nama = ""
  end if
  end if	%>
      <%=nama%>&nbsp;  
      <input type="submit" name="B1" value="Hantar" onFocus="nextfield='done';"  class="button">
      </td>
</tr>
</table>
<%end sub

  sub hantar %>
<br/>
 <table width="50%" cellpadding="1" cellspacing="5" class="hd1">
<%b1 =     "select rowid,kod,keterangan,paras,tamat "
  b1 = b1 & "  from majlis.menu_2002 where paras=1 and sistem = 'ik' order by kod "
  Set Rsb1 = objConn.Execute(b1)
 
  ctrz = 0
  Do while not Rsb1.EOF
     ctrz = cdbl(ctrz) + 1
     ket1 = rsb1("keterangan")
     kod1 = rsb1("kod") 
     rowid = rsb1("rowid")
     tamat1 = rsb1("tamat")
%>
<tr > 
  <td colspan=3 class="hd"><%=ket1%>
  <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
  <input type="hidden" name="kod1<%=ctrz%>" value="<%=kod1%>">
  </td>
</tr>

<%   b2 =      "select rowid,kod,keterangan,paras,tamat "
     b2 = b2 & "  from majlis.menu_2002 where kod like '"& kod1 &"'||'%' and paras=2 and sistem = 'ik' order by kod "
     Set Rsb2 = objConn.Execute(b2)
 
     Do while not Rsb2.EOF
     ctrz = cdbl(ctrz) + 1
     ket2 = rsb2("keterangan")
     kod2 = rsb2("kod") 
     rowid = rsb2("rowid")
     tamat2 = rsb2("tamat")

     b2a =       "select 'x' from majlis.kebenaran_2002 "
     b2a = b2a & " where no_pekerja = '"& pekerja &"' and sistem = 'ik' and skrin = '"& kod2 &"' "
     Set Rsb2a = objConn.Execute(b2a)
 
     b3 =      "select rowid,kod,keterangan,paras,tamat "
     b3 = b3 & "  from majlis.menu_2002 where kod like '"& kod2 &"'||'%' and paras=3 and sistem = 'ik' order by kod "
     Set Rsb3 = objConn.Execute(b3)
 
     if rsb3.eof then	%>
<tr > 
  <%if ket2z <> ket2 then%>
  <td colspan=3>
      <%if tamat2 = "Y" then%><input type="checkbox" name="ket2<%=ctrz%>" value="1"
      <% if not rsb2a.eof then%>checked<%end if%> >
      <%end if%>
      <%=ket2%>
      <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
      <input type="hidden" name="kod2<%=ctrz%>" value="<%=kod2%>">
      </td>
  <%else%>
  <td colspan=3 bgcolor="white" width="33%">&nbsp;</td>
  <%end if%>
</tr>
<%else
     Do while not Rsb3.EOF
     ctrz = cdbl(ctrz) + 1
     ket3 = rsb3("keterangan")
     kod3 = rsb3("kod") 
     tamat3 = rsb3("tamat")
     rowid = rsb3("rowid")

     b3a =       "select 'x' from majlis.kebenaran_2002 "
     b3a = b3a & " where no_pekerja = '"& pekerja &"' and sistem = 'ik' and skrin = '"& kod3 &"' "
     Set Rsb3a = objConn.Execute(b3a)
 
     b4 =      "select rowid,kod,keterangan,paras,tamat "
     b4 = b4 & "  from majlis.menu_2002 where kod like '"& kod3 &"'||'%' and paras=4 and sistem = 'ik' order by kod "
     Set Rsb4 = objConn.Execute(b4)
 
     if rsb4.eof then
%>
<tr > 
  <%if ket2z <> ket2 then%>
  <%if tamat2 = "Y" then%>
  <td bgcolor<%=color3%>>
      <input type="checkbox" name="ket2<%=ctrz%>" value="1"
      <% if not rsb2a.eof then%>checked<%end if%> ><%=ket2%>
      <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
      </td>
  <%else%>
  <td> <%=ket2%>  <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
      </td>
  <%end if%>
  <%else%>
  <td bgcolor="white" >&nbsp;</td>
  <%end if%>
  <td colspan=2 bgcolor="<%=color4%>" >
      <%if tamat3 = "Y" then%><input type="checkbox" name="ket3<%=ctrz%>" value="1"
      <% if not rsb3a.eof then%>checked<%end if%> >
      <%end if%>    <%if ket3z <> ket3 then%><%=ket3%><%end if%>
      <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
      <input type="hidden" name="kod3<%=ctrz%>" value="<%=kod3%>">
      </td>
</tr>
<%else
     Do while not Rsb4.EOF
     ctrz = cdbl(ctrz) + 1
     ket4 = rsb4("keterangan")
     kod4 = rsb4("kod") 

     b4a =       "select 'x' from majlis.kebenaran_2002 "
     b4a = b4a & " where no_pekerja = '"& pekerja &"' and sistem = 'ik' and skrin = '"& kod4 &"' "
     Set Rsb4a = objConn.Execute(b4a)	%> 
<tr > 
  <%if ket2z <> ket2 then%>
  <td bgcolor="<%=color3%>" >
      <%if tamat2 = "Y" then%><input type="checkbox" name="ket2<%=ctrz%>" value="1"
      <% if not rsb2a.eof then%>checked<%end if%> >
      <%end if%>
    <%=ket2%></td>
  <%else%>
  <td>&nbsp;</td>
  <%end if%>
  <%if ket3z <> ket3 then%>
  <td width="33%" bgcolor="<%=color4%>">
      <%if tamat3 = "Y" then%><input type="checkbox" name="ket3<%=ctrz%>" value="1"
      <% if not rsb3a.eof then%>checked<%end if%> >
      <%end if%><%=ket3%></td>
  <%else%>
  <td width="33%">&nbsp;</td>
  <%end if%>
  <td bgcolor="<%=color4%>"><input type="checkbox" name="ket4<%=ctrz%>" value="1"
      <% if not rsb4a.eof then%>checked<%end if%> >
<%=ket4%>
      <input type="hidden" name="rowid<%=ctrz%>" value="<%=rowid%>">
      <input type="hidden" name="kod4<%=ctrz%>" value="<%=kod4%>">
      </td>
</tr>
<%ket2z = ket2
  ket3z = ket3
  Rsb4.Movenext
  Loop   
  end if
  ket2z = ket2
  ket3z = ket3
  Rsb3.Movenext
  Loop 
  end if 
  Rsb2.Movenext
  Loop 
  Rsb1.Movenext
  Loop %> 
<tr> 
  <td colspan=3 align="center" class="hd"><input type="submit" name="B1" value="Simpan"  class="button"></td>
</tr>
</table>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
<%end sub %>
</form>
</body>
</html>