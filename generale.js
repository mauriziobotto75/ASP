function checkFields() {
missinginfo = "";
if (document.all.numcampi.value == "") 
{
    missinginfo = "inserire un valore";
}
if (isNaN(document.all.numcampi.value)) 
{
    missinginfo = "inserire un valore numerico";
}
if ((document.all.numcampi.value == 0) || (document.all.numcampi.value > 10)) 
  {
      missinginfo = "inserire un numero compreso tra 1 e 10";
  }

if (missinginfo != "") {

    alert(missinginfo);

return false;
}
else return true;
}
