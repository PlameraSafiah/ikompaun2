
<script type="text/javascript">
function dob(){
if (document.MyForm.kp1.value.length==6){
    var tahun=document.MyForm.kp1.value.substr(0,2);
    var bulan=document.MyForm.kp1.value.substr(2,2);
    var hari=document.MyForm.kp1.value.substr(4,2);

    document.MyForm.tkhlahir.value=hari+"/"+bulan+"/19"+tahun;
    }
}
</script>

<!-- cth : saya ringkaskan form awak
<form name="MyForm">
No ic<input name="no_ic" type="text" maxlength="12" onkeyup="dob()">
Tarikh lahir<input name="input" type="text" >
</form>-->