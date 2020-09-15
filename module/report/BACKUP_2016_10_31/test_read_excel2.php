<?php
//session_start();
//require_once '../../config/config.php';
//require_once '../../function/function.php';
//require_once '../../session_login.php';
//require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';
date_default_timezone_set("Asia/Bangkok");
$file_eksport=date('Y_m_d_H_i_s');



 $objPHPExcel = PHPExcel_IOFactory::load("download/Report_UMKM__2016_07_20_13_12_19.xls");
 //$sheetnames = array('LAPORAN A','LAPORAN B')
 $objWorksheet = $objPHPExcel->getActiveSheet();
 $objPHPExcel->setActiveSheetIndex(0);
 //$m61_is_report_day_min1=$objPHPExcel->getActiveSheet()->getCell('M61')->getValue('#,##0,,;(#,##0,,);"-"');
//$sheetnames = array('LAPORAN A','LAPORAN B'); 
//$objReader->setLoadSheetsOnly($sheetnames); 
/**  Create a new Reader of the type defined in $inputFileType  **/ 
//$objReader = PHPExcel_IOFactory::createReader($inputFileType); 
/**  Advise the Reader of which WorkSheets we want to load  **/ 
//$objReader->setLoadSheetsOnly($sheetnames); 
/**  Load $inputFileName to a PHPExcel Object  **/ 
//$objPHPExcel = $objReader->load("download/Report_UMKM__2016_07_20_13_12_19.xls");

 //echo "nilai sebelumnya : ".$objPHPExcel->getActiveSheet()->getCell('B11')->getValue();
//echo " ".$objPHPExcel->getActiveSheet('LAPORAN A')->getCell("G15")->getFormattedValue('#,##0,,;(#,##0,,);"-"');;

//$objReader->setLoadAllSheets();
//$objPHPExcel = $objReader->load("download/Report_UMKM__2016_07_20_13_12_19.xls");
echo " ".$objPHPExcel->getActiveSheet('LAPORAN B')->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"');;
?>
