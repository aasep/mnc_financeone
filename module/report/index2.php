<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Something');
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'A4');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'A5');

// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Name of Sheet 1');

// Create a new worksheet, after the default sheet
$objPHPExcel->createSheet();

// Add some data to the second sheet, resembling some different data types
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'More data');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'Ini B7');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'D1');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'F1');

// Rename 2nd sheet
$objPHPExcel->getActiveSheet()->setTitle('Second sheet');

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');