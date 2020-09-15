<?php
//session_start();
//require_once '../../config/config.php';
//require_once '../../function/function.php';
//require_once '../../session_login.php';
//require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';
date_default_timezone_set("Asia/Bangkok");




 $objPHPExcel = PHPExcel_IOFactory::load("download/NOP_16-May-16_2016_06_20_11_11_40.xls");
 $objWorksheet = $objPHPExcel->getActiveSheet('2');
//$x=$objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"');

echo $x=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('B2')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
 //echo $objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"');
function getTextValue($num){
$nilai=str_replace(array('(', ')',','),'', (string)$num);
$num_length = strlen($nilai);
if ($num_length=='1')
{
     $txt_record="00000000".$nilai;
}   else if ($num_length=='2'){
     $txt_record="0000000".$nilai;
}   else if ($num_length=='3'){
     $txt_record="000000".$nilai;
}   else if ($num_length=='4'){
     $txt_record="00000".$nilai;
}   else if ($num_length=='5'){
     $txt_record="0000".$nilai;
}   else if ($num_length=='6'){
     $txt_record="000".$nilai;
}   else if ($num_length=='7'){
     $txt_record="00".$nilai;
}   else if ($num_length=='8'){
     $txt_record="0".$nilai;
}   else if ($num_length=='9'){
     $txt_record="".$nilai;
}     

 return $txt_record;
}





 if (!isset($x) || $x=="" || $x==NULL || $x==0){
$var_giro_aud="aaaaaa";
}else {
$var_giro_aud="15AUD".getTextValue($x).PHP_EOL;
//$jml_baris_txt++; 
}


echo $var_giro_aud;
 //echo $objWorksheet->getCellByColumnAndRow(5, 17)->getValue();

//$cell = $objWorksheet->getCellByColumnAndRow(3, 5);
//$cell_value = PHPExcel_Style_NumberFormat::toFormattedString($cell->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"'), 'hh:mm:ss');
//echo $cell_value;

//function getCell( PHPExcel_Worksheet $sheet, /* string */ $x = 'A', /* int */ $y = 1 ) {

 //   return $sheet->getCell( $x . $y );

//}

// eg:
//echo getCell( $sheet, 'E', 17 )->getValue();

?>
