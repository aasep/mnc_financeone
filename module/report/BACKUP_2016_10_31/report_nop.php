<?php
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';


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




// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));

//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(2);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(2);
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'MASA LAPORAN');
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'BANK : PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('E9', 'VAR DATE');
$objPHPExcel->getActiveSheet()->getStyle('B9:E10')->applyFromArray($styleArrayFont);
$objPHPExcel->getActiveSheet()->getStyle('A13:K14')->applyFromArray($styleArrayAlignment);

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B13:C14');//No
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D13:D14'); //keterangan
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E13:J13'); //AUD sd USD
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K13:K14'); //Jumlah
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('D13', 'Keterangan');
$objPHPExcel->getActiveSheet()->setCellValue('E14', 'AUD');
$objPHPExcel->getActiveSheet()->setCellValue('F14', 'EUR');
$objPHPExcel->getActiveSheet()->setCellValue('G14', 'HKD');
$objPHPExcel->getActiveSheet()->setCellValue('H14', 'JPY');
$objPHPExcel->getActiveSheet()->setCellValue('I14', 'SGD');
$objPHPExcel->getActiveSheet()->setCellValue('J14', 'USD');
$objPHPExcel->getActiveSheet()->setCellValue('K13', 'Jumlah');
$objPHPExcel->getActiveSheet()->getStyle('B13:K14')->applyFromArray($styleArrayFont);
//A.NERACA
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'A.');
$objPHPExcel->getActiveSheet()->setCellValue('D15', 'Neraca');
$objPHPExcel->getActiveSheet()->getStyle('B15:Z15')->applyFromArray($styleArrayFont);
//1. Aktiva Valas
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C16', '1');
$objPHPExcel->getActiveSheet()->setCellValue('D16', 'Aktiva Valsa');
$objPHPExcel->getActiveSheet()->getStyle('B16:Z16')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D17', '- Aktiva Valas tidak termasuk giro pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('E17', $angka);
$objPHPExcel->getActiveSheet()->setCellValue('D18', '- Giro pada bank lain');
$objPHPExcel->getActiveSheet()->getStyle('A17:Z18')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C20', '2');
$objPHPExcel->getActiveSheet()->setCellValue('D20', 'Aktiva Valsa');

$objPHPExcel->getActiveSheet()->setCellValue('C21', '3');
$objPHPExcel->getActiveSheet()->setCellValue('D21', 'Selisih Aktiva dan Pasiva Valas (A.1 - A.2)');
$objPHPExcel->getActiveSheet()->getStyle('A20:Z21')->applyFromArray($styleArrayFont);

$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D22', 'Selisih Aktiva dan Pasiva Valas (Nilai Absolut)');
$objPHPExcel->getActiveSheet()->getStyle('A22:Z22')->applyFromArray($styleArrayFont);

//B. REKENING ADMINISTRATIF
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'B.');
$objPHPExcel->getActiveSheet()->setCellValue('D24', 'Rekening Administratif');
$objPHPExcel->getActiveSheet()->setCellValue('C25', '1');
$objPHPExcel->getActiveSheet()->setCellValue('D25', 'Rekening Administratif Tagihan Valas');
$objPHPExcel->getActiveSheet()->getStyle('A24:Z25')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D26', 'a. Kontrak pembelian forward');
$objPHPExcel->getActiveSheet()->setCellValue('D27', 'b. Kontrak pembelian futures');
$objPHPExcel->getActiveSheet()->setCellValue('D28', 'c. Kontrak penjualan put options (bank sebagai writter)');
$objPHPExcel->getActiveSheet()->setCellValue('D29', 'd. Kontrak pembelian call options (bank sebagai');
$objPHPExcel->getActiveSheet()->setCellValue('D30', '   holder, khusus back to back options)');
$objPHPExcel->getActiveSheet()->setCellValue('D31', 'e. Rekening Administratif Tagihan Valas diluar ');
$objPHPExcel->getActiveSheet()->setCellValue('D32', '    kontrak pemberian forward, futures, dan option');
$objPHPExcel->getActiveSheet()->getStyle('A26:D32')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C34', '2');
$objPHPExcel->getActiveSheet()->setCellValue('D34', 'Rekening Administratif Kewajiban Valas');
$objPHPExcel->getActiveSheet()->getStyle('A34:Z34')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D35', 'a. Kontrak penjualan forward');
$objPHPExcel->getActiveSheet()->setCellValue('D36', 'b. Kontrak penjualan futures');
$objPHPExcel->getActiveSheet()->setCellValue('D37', 'c.  Kontrak penjualan call options (bank sebagai writter)');
$objPHPExcel->getActiveSheet()->setCellValue('D38', 'd. Kontrak pembelian put options (bank sebagai');
$objPHPExcel->getActiveSheet()->setCellValue('D39', '   holder, khusus back to back option)');
$objPHPExcel->getActiveSheet()->setCellValue('D40', 'e. Rekening Administratif Kewajiban Valas diluar');
$objPHPExcel->getActiveSheet()->setCellValue('D41', '	kontrak penjualan forward, futures, dan option
');
$objPHPExcel->getActiveSheet()->getStyle('A35:D41')->applyFromArray($styleArrayFont);
//SELISIH REKENING ADM
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C43', '3');
$objPHPExcel->getActiveSheet()->setCellValue('D43', 'Selisih Rekening Administratif (B.1 - B.2)');
$objPHPExcel->getActiveSheet()->getStyle('A43:D43')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'C.');
$objPHPExcel->getActiveSheet()->setCellValue('D45', 'Posisi Devisa Netto per Valuta');
$objPHPExcel->getActiveSheet()->setCellValue('D46', '(A.3 + B.3)');
$objPHPExcel->getActiveSheet()->getStyle('A45:Z45')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B48', 'D.');
$objPHPExcel->getActiveSheet()->setCellValue('D48', 'Posisi Devisa Netto');
$objPHPExcel->getActiveSheet()->setCellValue('D49', '(Nilai Absolut C)');
$objPHPExcel->getActiveSheet()->getStyle('A48:Z49')->applyFromArray($styleArrayFont);

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 9,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B51', 'E.');
$objPHPExcel->getActiveSheet()->setCellValue('D51', 'Modal dalam Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B53', 'F.');
$objPHPExcel->getActiveSheet()->setCellValue('D53', '% PDN terhadap modal (A/E) Neraca');
$objPHPExcel->getActiveSheet()->setCellValue('B55', 'G.');
$objPHPExcel->getActiveSheet()->setCellValue('D55', '% PDN terhadap modal (D/E) Neraca & Rek. Adm.');

$objPHPExcel->getActiveSheet()->getStyle('A51:Z55')->applyFromArray($styleArrayFont);

//=======BORDER
$styleArray = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('E15:K57')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('D15:D57')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B15:C57')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$objPHPExcel->getActiveSheet()->getStyle('B13:C14')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('D13:D14')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('E13:J13')->applyFromArray($styleArray);
$styleArray = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('K13:K15')->applyFromArray($styleArray);
$styleArray = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('E14:J14')->applyFromArray($styleArray);
//=======END BORDER


$objPHPExcel->getActiveSheet()->setCellValue('H14', 'JPY');
$objPHPExcel->getActiveSheet()->setCellValue('I14', 'SGD');
$objPHPExcel->getActiveSheet()->setCellValue('J14', 'USD');
$objPHPExcel->getActiveSheet()->setCellValue('K13', 'Jumlah');



// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('FORMAT PDN BANK GARANSI');

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