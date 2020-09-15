<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$objPHPExcel->getActiveSheet()->getStyle('A1:E1')->applyFromArray($styleArrayAlignment);//title
//DIMENSION D
//$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:E1');//merge title
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:B3');//merge title
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A4:B4');//merge title
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A5:B5');//merge title



// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'LAPORAN POSISI SIMPANAN');

$objPHPExcel->getActiveSheet()->setCellValue('A3', 'PER AKHIR BULAN :');
$objPHPExcel->getActiveSheet()->setCellValue('C3', 'Saturday, October 31, 2015 :');
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'TAHUN :');
$objPHPExcel->getActiveSheet()->setCellValue('C4', '2015');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'BANK :');
$objPHPExcel->getActiveSheet()->setCellValue('C5', 'Bank MNC Internasional, Tbk');
$objPHPExcel->getActiveSheet()->getStyle('A1:E5')->applyFromArray($styleArrayFont);




// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');