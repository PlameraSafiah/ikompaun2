<%response.cookies("ikmenu") = "ik82.asp"%>
<!-- '#INCLUDE FILE="ik.asp" -->
<META content="text/html; charset=iso-8859-1" http-equiv=Content-Type>
<link type="text/css" href="menu.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<form method="Post" action="ik82.asp" name="ik82">   
    <%
	bilrek=request.form("bilrek")
	saves=request.form("savea")
	reset=request.form("reset")
	edmode=request.form("edmode")
	
	if reset="Reset" then
		bilrek=""
		edmode="off"
		emode="off"
		saves=""
		koda=""
	end if	
	
	if saves="Simpan" then
		bilrek="S" 
	end if
		
	if bilrek="" then
		emode="off"
		papar
		
	elseif bilrek<>"" and bilrek<>"S" then
		saves=""		
		for i=0 to bilrek
			rowid="rowid"+cStr(i)
			edit="edit"+cStr(i)
			hap="hap"+cStr(i)
			kod="kod"+cStr(i)
			
			rowidz=request.form(""&rowid&"")
			editz=request.form(""&edit&"")
			hapz=request.form(hap)
			kodz=request.form(""&kod&"")
			
			if editz="Edit" then
				koda=kodz
				rowida=rowidz						
				emode="on"
				papar												

			elseif hapz="Hapus" then
	   			del = " delete hasil.superadmin where rowid = '"& rowidz &"' "
			   	Set objRsdel = objConn.Execute(del)
				emode="off"				
				papar				
								
			end if			
					
		next
	end if
	
	if saves="Simpan" and bilrek="S" then			
			kodc=request.form("kodTxt")
			rowidc=request.form("rid")

			if edmode="on" then
				if kodc="" then
					response.write "<script language = ""vbscript"">"
		    		response.write " MsgBox ""Sila masukkan nilai ke dalam kotak input."", vbInformation, ""Perhatian!"" "
					response.write "</script>"					
					emode="off"
				
				else	
					ed="update hasil.superadmin "
					ed=ed&"set id='"&kodc&"' "
					ed=ed&"where rowid='"& rowidc &"'"			
					Set objRSed = objConn.Execute(ed)			
					emode="off"
				
				end if

			elseif edmode="off" then
				if kodc="" then
					response.write "<script language = ""vbscript"">"
		    		response.write " MsgBox ""Sila masukkan nilai ke dalam kotak input."", vbInformation, ""Perhatian!"" "
					response.write "</script>"					
					emode="off"					
			
				else			
    				check = " select 'x' from hasil.superadmin where id = '"& kodc &"' "
    				Set objRScheck = objConn.Execute(check)				
				
		    		if not objRScheck.eof then
        				response.write "<script language = ""vbscript"">"
		        		response.write " MsgBox ""Kod " + kodc + " sudah wujud! Sila pilih kod lain."", vbInformation, ""Perhatian!"" "
		        		response.write "</script>"
						emode="off"
						bilrek=""				

		    		else
						ins="insert into hasil.superadmin "
						ins=ins&"values ('"&kodc&"')"
						Set objRSadd = objConn.Execute(ins)
						emode="off"
						bilrek=""					
					
					end if
				end if	
					
			end if							
			papar
		
	 end if
 
	sub papar
%>
    <!--Bahagian Edit dan Tambah Kod-->
    <input type="hidden" name="edmode" value="<%=emode%>">
 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
    <tr >
	 <td colspan="4" align="center" ><b>Daftar Superadmin </b></td>
    </tr>
    <tr >
     <td width="25%" class="hd1">Id Pekerja</td>
     <td><input type="text" name="kodTxt" size="5" maxlength="5" value="<%=koda%>" onKeyDown="if(event.keyCode==13) event.keyCode=9;">
      *</td>
    </tr>
    <tr >
     <td colspan="4" align="center"><input type="submit" name="savea" value="Simpan" onFocus="nextfield='done';" class="button" >
      <input type="hidden" name="rid" value="<%=rowida%>">
      <input type="submit" name="reset" value="Reset"  class="button"></td>
    </tr>
</table>
<br/>

 <table width="50%" cellpadding="1" cellspacing="5" class="hd">
      <tr bgcolor="<%=color2%>" align="center"> 
        <td width="6%" class="hd1">Kod</td>
        <td width="60%" class="hd1">Nama</td>
        <td colspan="2" class="hd1">Proses</td>
      </tr>

      <%
		b = "select rowid, id "
		b = b&" from hasil.superadmin order by id "
		set sb = objConn.Execute(b)
	
	    ctrz = 0
		Do while not sb.EOF
		ctrz = ctrz + 1
		
		kod  = sb("id")
		bb = " Select nama from payroll.paymas where no_pekerja = '"& kod &"'"
		set sbb = objConn.Execute(bb)
		
		nama  = sbb("nama")
		
		warna = ctrz mod 2
%>
      <!--Bahagian Paparan Kod Sedia Ada-->
      <tr align="center" <%if warna ="1" then %>bgcolor="<%=color3%>" <%else%>bgcolor="<%=color4%>" <%end if%>> 
        <td><%=kod%></td>
        <td align="left">&nbsp;<%=nama%></td>
       
        <td> 
        <input type="submit" name="hap<%=ctrz%>" value="Hapus" onClick="return confirm('Anda pasti?')" class="button">
        <input type="hidden" name="kod<%=ctrz%>" value="<%=kod%>" >
        <input type="hidden" name="ket<%=ctrz%>" value="<%=ket%>" >
        <input type="hidden" name="nrg<%=ctrz%>" value="<%=nrg%>" >
        <input type="hidden" name="rowid<%=ctrz%>" value="<%=sb("rowid")%>"></td>
      </tr>
      <%sb.MoveNext
		Loop	%>
      <input type="hidden" name="bilrek" value="<%=ctrz%>">
      <%	end sub	%>
    </table>  
</form>
