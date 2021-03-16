<?php
$file = "C:/documenti/PROVA3SQL.TXT";
$name = basename($file);
if (file_exists($name))
 {
   $data_last=getdate(fileatime($name));
   $anno = $data_last["year"];
   $mese = $data_last["mon"];
   $giorno = $data_last["mday"];
   $ora    = $data_last["hours"];
   $minuti = $minuti["minutes"];
   $data_video = "$giorno/$mese/$anno";
   echo  "ULTIMO ACCESSo ". $data_video;
 } 
 	
?>


