<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

/*
$objPHPExcel->getDefaultStyle()
    ->getBorders()
    ->getTop(A5)
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getDefaultStyle()
    ->getBorders()
    ->getBottom(A5)
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getDefaultStyle()
    ->getBorders()
    ->getLeft(A5)
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getDefaultStyle()
    ->getBorders()
    ->getRight(A5)
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);


*/

//$objPHPExcel->getDefaultStyle()
 //   ->getBorders()
 //   ->getRight(A5)
 //       ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

// Create a first sheet, representing sales data
//Setting for borders   
$styleArray = array('borders' => array('top' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B2:K2')->applyFromArray($styleArray);


// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');