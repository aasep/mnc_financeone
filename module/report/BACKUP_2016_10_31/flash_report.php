<?php
//require_once 'config/config.php';
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';


$tanggal=$_POST['tanggal'];
$day=date('d',strtotime($tanggal));
$day_min1=date('j',strtotime($tanggal))-1;

if (strlen($day_min1)==1){
$day_min1="0".$day_min1;
}
//$day_min1=date('d',strtotime($day_min))-1;
//$day_min1=$day-1;
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$prev_date=date('t-M-y', strtotime(date('Y-m',strtotime($tanggal))." -1 month"));

$label_tgl=$day."-".$mon."-".$year;
$label_bln=$mon."-".$year;
//$label_tgl_min1=$day_min1."-".$mon."-".$year;

$label_tgl_min1=date('d-M-y', strtotime(date('Y-m-d',strtotime($tanggal))." -1 day"));
//$prev_mon = "2009-11-23";
//echo date("Y-m-t", strtotime($a_date));

/*
$query="SELECT SUM(NOMINAL) as angka
FROM (
select GL_02_Baru.NOP_Level_3,SUM(DM_Journal.Nominal) as NOMINAL,Referensi_NOP.NOP_Level_3_Description
from DM_Journal
join GL_02_Baru on DM_Journal.KodeGL =  GL_02_Baru.GLNO
join Referensi_GL_01 on Referensi_GL_01.PM_COA_Level_4 = GL_02_Baru.PM_COA_Level_4 
join Referensi_NOP on GL_02_Baru.NOP_Level_3 = Referensi_NOP.NOP_Level_3
WHERE DM_Journal.DataDate='2015-09-30' AND DM_Journal.JenisMataUang='AUD' AND GL_02_BarU.NOP_Level_3='NOP101000001'
group by DM_Journal.Nominal ,GL_02_Baru.NOP_Level_3,Referensi_NOP.NOP_Level_3_Description
) as table_alias";

$result=odbc_exec($connection, $query);

$row=odbc_fetch_array($result);
$angka=$row['angka'];
*/



// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
//$styleArraybackgroundRed = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignmentCenter = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignmentCenter2 = array('alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));
//BOLD
$objPHPExcel->getActiveSheet()->getStyle('B1:B3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B60')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B63')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B71')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B56')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B57')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B49')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B37')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B22')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B28')->applyFromArray($styleArrayFontBold);
//Bakgroud
//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArraybackgroundRed);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B60:C60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B63:C63')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B71:C71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');

//CENTER
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayAlignmentCenter2);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayAlignmentCenter2);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayAlignmentCenter);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayAlignmentCenter);
//$objPHPExcel->getActiveSheet()->getStyle('B5:B5')->applyFromArray($styleArrayAlignmentCenter);
//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'PT BANK MNC INTERNASIONAL TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'BALANCE SHEET');
$objPHPExcel->getActiveSheet()->setCellValue('B3', $label_tgl);

//GLOBAL


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:A1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K1:Z1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B73:C1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:J1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:Z1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B1:J1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B2:J2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B3:J3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B4:J4');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B5:B7');//Account of Assets
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C5:G6');//
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H5:J6');//
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A73:B73');//


$objPHPExcel->getActiveSheet()->setCellValue('C5', 'For The Month');
$objPHPExcel->getActiveSheet()->setCellValue('H5', $label_bln);

$objPHPExcel->getActiveSheet()->setCellValue('C7', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('D7', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('F7', $prev_date);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'Var MTD');
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'Actual');
$objPHPExcel->getActiveSheet()->setCellValue('I7', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'Var');




		
	
		
$objPHPExcel->getActiveSheet()->setCellValue('B5', 'Account of Assets');	

$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Cash');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Current account - Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Certificate of bank Indonesia (SBI & BI call money)	');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Interbank Placement');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Securities	');
$objPHPExcel->getActiveSheet()->setCellValue('B13', '-	Allowance For Securities');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Loans');
$objPHPExcel->getActiveSheet()->setCellValue('B15', '-	Performing Loan');
$objPHPExcel->getActiveSheet()->setCellValue('B16', '-	Non Performing Loan*)	');
$objPHPExcel->getActiveSheet()->setCellValue('B17', '-	Allowance For Loan	');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Acceptance receivables	');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Derivative receivables');
$objPHPExcel->getActiveSheet()->setCellValue('B20','Fixed assets (Property, Plant Equipment)');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Deferred taxes	');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Other assets');
$objPHPExcel->getActiveSheet()->setCellValue('B23', '-	Foreclosed properties');
$objPHPExcel->getActiveSheet()->setCellValue('B24', '- 	Allowance For Foreclosed properties	');
$objPHPExcel->getActiveSheet()->setCellValue('B25', '-	Account receivable	');
$objPHPExcel->getActiveSheet()->setCellValue('B26', '-	Others');
$objPHPExcel->getActiveSheet()->setCellValue('B27', '-	Allowances For Suspence Account	');
$objPHPExcel->getActiveSheet()->setCellValue('B28', 'TOTAL ASSETS');


$objPHPExcel->getActiveSheet()->setCellValue('C31', 'For The Month');
$objPHPExcel->getActiveSheet()->setCellValue('H31', $label_bln);

$objPHPExcel->getActiveSheet()->setCellValue('C33', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('D33', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E33', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('F33', $prev_date);
$objPHPExcel->getActiveSheet()->setCellValue('G33', 'Var MTD');
$objPHPExcel->getActiveSheet()->setCellValue('H33', 'Actual');
$objPHPExcel->getActiveSheet()->setCellValue('I33', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('J33', 'Var');




$objPHPExcel->getActiveSheet()->setCellValue('B31', 'Account of Liabilities & Equity');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B31:B33');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C31:G32');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H31:J32');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Current Account');
$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Saving Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Time Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Total deposits	');
$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Interbank');
$objPHPExcel->getActiveSheet()->setCellValue('B39', '-	Call Money');
$objPHPExcel->getActiveSheet()->setCellValue('B40', '-	Bank Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('B41', '-	Current account	');
$objPHPExcel->getActiveSheet()->setCellValue('B42', '-	Saving Account	');
$objPHPExcel->getActiveSheet()->setCellValue('B43', 'Derivative payable	');
$objPHPExcel->getActiveSheet()->setCellValue('B44', 'Acceptance payable	');
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'KLBI Payable');
$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Mandatory Convertible Bonds');
$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Securities sold with agreement to repurchase');
$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('B49', 'Total Other Liabilities');
$objPHPExcel->getActiveSheet()->setCellValue('B50', 'Paid in capital');
$objPHPExcel->getActiveSheet()->setCellValue('B51', 'Agio ( disagio)');
$objPHPExcel->getActiveSheet()->setCellValue('B52', 'General reserve');
$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Available for sale securities - net');
$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Retained earnings');
$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Profit/loss current year');
$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Total Equity');
$objPHPExcel->getActiveSheet()->setCellValue('B57', 'TOTAL LIABILITIES & EQUITY');

	
 	
 	
	
$objPHPExcel->getActiveSheet()->setCellValue('B60', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'New NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Penambah_OS_NPL');
$objPHPExcel->getActiveSheet()->setCellValue('B63', 'Total New NPL');
$objPHPExcel->getActiveSheet()->setCellValue('B64', '');
$objPHPExcel->getActiveSheet()->setCellValue('B65', 'NPL to PL (Reklass) ');
$objPHPExcel->getActiveSheet()->setCellValue('B66', 'NPL Paid Off');
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'Reverse Saldo NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'NPL Payment');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'NPL Exchange Rate');
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'NPL Credit Card');
$objPHPExcel->getActiveSheet()->setCellValue('B71', $label_tgl);

// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('Flash Report');

//=======BORDER
$styleArray = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$objPHPExcel->getActiveSheet()->getStyle('B5:J28')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B31:J57')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B60:C71')->applyFromArray($styleArray);
//=======END BORDER

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');