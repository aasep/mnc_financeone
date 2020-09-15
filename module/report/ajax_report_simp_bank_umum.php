<?php
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';


date_default_timezone_set("Asia/Bangkok");

$file_eksport=date('Y_m_d_H_i_s');

error_reporting(1);
logActivity("generate skb efektif",date('Y_m_d_H_i_s'));

######## POST DATE ##############
$tanggal=$_POST['tanggal']; 

$curr_tgl=date('Y-m-d',strtotime($tanggal));
$end_curr_tgl=date('Y-m-t',strtotime($tanggal));

$label_tgl=date('d-M-y',strtotime($tanggal));


$label_txtfile=date('Ymd',strtotime($tanggal));
$tanggal_header=date('dmY',strtotime($tanggal));


$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih

$bulan=date('m',strtotime($tanggal));
$tahun=date('Y',strtotime($tanggal));

$var_tabel=date('Ymd',strtotime($tanggal));


#############################################################################################
$table_giro="DM_LiabilitasGiro_$var_tabel";
$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

$qcekTable=" select top 1* from  $table_giro ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger'> <br><b> Table $table_giro is not available.</b> <br><br></div>";
        die();
}

$qcekTable=" select top 1* from  $table_tabungan ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger'> <br><b> Table $table_tabungan is not available.</b> <br><br></div>";
        die();
}

$qcekTable=" select top 1* from  $table_deposito ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger'> <br><b> Table $table_deposito is not available.</b> <br><br></div>";
        die();
}

$qcekTable=" select top 1* from  $table_banklain ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger'> <br><b> Table $table_banklain is not available.</b> <br><br></div>";
        die();
}


########################### FROM LPS ###############################





#--Query Lampiran 1B--
#--Simpanan Pihak Ketiga Rupiah--


$query_kurs=" select Kurs_Tengah from DM_Kurs where Jenis_Mata_Uang='USD' and DataDate='$curr_tgl' ";
        $result=odbc_exec($connection2, $query_kurs);
        $row_kurs=odbc_fetch_array($result);

        $nilai_kurs=floatval($row_kurs['Kurs_Tengah']);





$range1=" and JumlahBulanLaporan >='0' and JumlahBulanLaporan <'100000001' ";
$range2=" and JumlahBulanLaporan >='100000001' and JumlahBulanLaporan <'200000001' ";
$range3=" and JumlahBulanLaporan >='200000001' and JumlahBulanLaporan <'500000001' ";
$range4=" and JumlahBulanLaporan >='500000001' and JumlahBulanLaporan <'1000000001' ";
$range5=" and JumlahBulanLaporan >='1000000001' and JumlahBulanLaporan <'2000000001' ";
$range6=" and JumlahBulanLaporan >='2000000001' and JumlahBulanLaporan <'5000000001' ";
$range7=" and JumlahBulanLaporan > '5000000000' ";


$array_range1=array("$range1","$range2","$range3","$range4","$range5","$range6","$range7");

$query =" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

and a.status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' and a.status_oncall='N'
 ";


//echo $query;
//die();
        $result2=odbc_exec($connection2, $query);

        $xcounter=1;
        //$row2=odbc_fetch_array($result2);
        while ($row2=odbc_fetch_array($result2)) {
           


        switch ($row2['FLAG']) {
            case 'GIRO':
                $d10=$row2['Nominal_Rupiah'];
                break;
            case 'TABUNGAN':
                $d11=$row2['Nominal_Rupiah'];
                break;
            case 'DEPOSIT ON CALL':
                $d12=$row2['Nominal_Rupiah'];
                break;
            case 'DEPOSITO':
                $d13=$row2['Nominal_Rupiah'];
                break;
           
        }

        //if ($xcounter=='1'){
        //    $d10=$row2['Nominal_Rupiah'];
        //}
            $xcounter++;
        }
        

//echo $d10;
//die();

#--Simpanan Pihak Ketiga Valuta Asing --


$query=" SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR' 

UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR' 

UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR'and a.status_oncall='Y' 

UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR'and a.status_oncall='N' 

)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas ";




        $result2=odbc_exec($connection2, $query);
        //$row2=odbc_fetch_array($result2);
        while ($row2=odbc_fetch_array($result2)) {
           
        switch ($row2['FLAG']) {
            case 'GIRO':
                $e10=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'TABUNGAN':
                $e11=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'DEPOSIT ON CALL':
                $e12=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'DEPOSITO':
                $e13=$row2['JumlahBulanLaporan_Valas'];
                break;
            
            default:
                # code...
                break;
        }
        }



#--Simpanan Bank Lain Rupiah--


$query=" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 
AND a.FlagDPK='Deposito' and a.status_oncall='y' 

UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang='IDR' 

AND a.FlagDPK='Deposito' ";


//echo $query."<br>";
//echo $query;
//die();
 $result2=odbc_exec($connection2, $query);
        //$row2=odbc_fetch_array($result2);
        while ($row2=odbc_fetch_array($result2)) {
           
        switch ($row2['FLAG']) {
            case 'GIRO':
                $d18=$row2['Nominal_Rupiah'];
                break;
            case 'TABUNGAN':
                $d19=$row2['Nominal_Rupiah'];
                break;
            case 'DEPOSIT ON CALL':
                $d20=$row2['Nominal_Rupiah'];
                break;
            case 'DEPOSITO':
                $d21=$row2['Nominal_Rupiah'];
                break;
            
            default:
                # code...
                break;
        }
        }

#--Simpanan Bank Lain Valuta Asing --
       


$query=" SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas

FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR' 

AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR' 

AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='Y' 

UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
WHERE   a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='N' 

)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas ";

//echo $query."<br>";

//die();

 $result2=odbc_exec($connection2, $query);
        //$row2=odbc_fetch_array($result2);
        while ($row2=odbc_fetch_array($result2)) {
           
        switch ($row2['FLAG']) {
            case 'GIRO':
                $e18=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'TABUNGAN':
                $e19=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'DEPOSIT ON CALL':
                $e20=$row2['JumlahBulanLaporan_Valas'];
                break;
            case 'DEPOSITO':
                $e21=$row2['JumlahBulanLaporan_Valas'];
                break;
            
            default:
                # code...
                break;
        }
        }





$label_nominal=array("0 < Nominal <= 100 Jt","100 Jt < Nominal <= 200 Jt ","200 Jt < Nominal <= 500 Jt ","500 Jt < Nominal <= 1M ","1M < Nominal <= 2M ","2M < Nominal <= 5M","Nominal > 5M");

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$objPHPExcel->getActiveSheet()->getStyle('A1:E6')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:E9')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A15:E17')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A23:E30')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A32:E32')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('D36:E43')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:E8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:E8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A15:C16')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A15:C16')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A23:B25')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A23:B25')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('A28:E29')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:E29')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('A32')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A32')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('D42:E43')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('D42:E43')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('A9:A25')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A9:A25')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A30:A32')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A30:A32')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('B34')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A7:E25')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A28:E32')->applyFromArray($styleArrayBorder1);



//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:Z6')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A26:Z27')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A33:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A7:E8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A28:E29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');

$objPHPExcel->getActiveSheet()->getStyle('E16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');

$objPHPExcel->getActiveSheet()->getStyle('E24:E25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);


$objPHPExcel->getActiveSheet()->getStyle("D10:E25")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:E1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A7:A8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B7:C8');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B15:C15');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B16:C16');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B23:C23');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B24:C24');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B25:C25');


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A28:A29');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B28:C29');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D36:E36');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D37:E37');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D42:E42');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D43:E43');



//$objPHPExcel->getActiveSheet()->getStyle('C13:E21')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('A1', 'LAPORAN POSISI SIMPANAN');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "PER AKHIR BULAN");
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'TAHUN');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'BANK');

$objPHPExcel->getActiveSheet()->setCellValue('C3', ": $label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('C4', ": $year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('C5', ': PT. BANK MNC INTERNASIONAL, TBK');


#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('A7', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'Bentuk Simpanan');
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Valuta Asing (Ekuivalen Rupiah)');

$objPHPExcel->getActiveSheet()->setCellValue('D8', "(i)");
$objPHPExcel->getActiveSheet()->setCellValue('E8', "(ii)");

$objPHPExcel->getActiveSheet()->setCellValue('A9', 'A.');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Simpanan Pihak Ketiga');

$objPHPExcel->getActiveSheet()->setCellValue('A10', '1');
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Giro');

$objPHPExcel->getActiveSheet()->setCellValue('A11', '2');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Tabungan');

$objPHPExcel->getActiveSheet()->setCellValue('A12', '3');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Deposito On Call (DOC)');

$objPHPExcel->getActiveSheet()->setCellValue('A13', '4');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Deposito');

$objPHPExcel->getActiveSheet()->setCellValue('A14', '5');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Setifikat Deposito');



$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Sub Total Simpanan Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B16', "A. Total Simpanan Pihak Ketiga Dalam Rupiah (i) + (ii)");


$objPHPExcel->getActiveSheet()->setCellValue('A17', "B.");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Simpanan Dari Bank Lain");

$objPHPExcel->getActiveSheet()->setCellValue('A18', '6');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Giro');

$objPHPExcel->getActiveSheet()->setCellValue('A19', '7');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Tabungan');

$objPHPExcel->getActiveSheet()->setCellValue('A20', '8');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Deposito On Call (DOC)');

$objPHPExcel->getActiveSheet()->setCellValue('A21', '9');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Deposito');

$objPHPExcel->getActiveSheet()->setCellValue('A22', '10');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Setifikat Deposito');

$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Sub Total Simpanan Dari Bank Lain (6 s/d 10)');
$objPHPExcel->getActiveSheet()->setCellValue('B24', "B. Total Simpanan Dari Bank lain Dalam Rupiah (i) + (ii)");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "Total Dalam Rupiah (A+B) ");


$objPHPExcel->getActiveSheet()->setCellValue('A27', "Cabang di Luar Negeri (Bagi Cabang yang memiliki kantor cabang di luar negeri)");
$objPHPExcel->getActiveSheet()->setCellValue('A28', "No.");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "Bentuk Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('D28', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('E28', "Valuta Asing (Ekuivalen Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('D29', "(i)");
$objPHPExcel->getActiveSheet()->setCellValue('E29', "(ii)");

$objPHPExcel->getActiveSheet()->setCellValue('A30', "C.");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "Simpanan Cabang di Luar Negeri");

$objPHPExcel->getActiveSheet()->setCellValue('A31', "11");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "Simpanan Cabang di Luar Negeri");

$objPHPExcel->getActiveSheet()->setCellValue('B32', "C. Total Simpanan Cabang Di Luar Negeri Dalam Rupiah");


$objPHPExcel->getActiveSheet()->setCellValue('B34', "Kurs USD1 Rp. ");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "*) Kewajiban dalam valuta asing selain USD dikonversikan ke dalam USD");

$objPHPExcel->getActiveSheet()->setCellValue('D36', "Jakarta , ");
$objPHPExcel->getActiveSheet()->setCellValue('D37', "PT BANK MNC INTERNASIONAL Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('D42', "Benny Purnomo");
$objPHPExcel->getActiveSheet()->setCellValue('D43', "Presiden Direktur");



$objPHPExcel->getActiveSheet()->setCellValue('c34', $nilai_kurs);

$objPHPExcel->getActiveSheet()->setCellValue('D10', floatval($d10));
$objPHPExcel->getActiveSheet()->setCellValue('D11', floatval($d11));
$objPHPExcel->getActiveSheet()->setCellValue('D12', floatval($d12));
$objPHPExcel->getActiveSheet()->setCellValue('D13', floatval($d13));
$objPHPExcel->getActiveSheet()->setCellValue('D14', floatval($d14));

$objPHPExcel->getActiveSheet()->setCellValue('E10', floatval($e10));
$objPHPExcel->getActiveSheet()->setCellValue('E11', floatval($e11));
$objPHPExcel->getActiveSheet()->setCellValue('E12', floatval($e12));
$objPHPExcel->getActiveSheet()->setCellValue('E13', floatval($e13));
$objPHPExcel->getActiveSheet()->setCellValue('E14', floatval($e14));




$objPHPExcel->getActiveSheet()->setCellValue('D18', floatval($d18));
$objPHPExcel->getActiveSheet()->setCellValue('D19', floatval($d19));
$objPHPExcel->getActiveSheet()->setCellValue('D20', floatval($d20));
$objPHPExcel->getActiveSheet()->setCellValue('D21', floatval($d21));
$objPHPExcel->getActiveSheet()->setCellValue('D22', floatval($d22));

$objPHPExcel->getActiveSheet()->setCellValue('E18', floatval($e18));
$objPHPExcel->getActiveSheet()->setCellValue('E19', floatval($e19));
$objPHPExcel->getActiveSheet()->setCellValue('E20', floatval($e20));
$objPHPExcel->getActiveSheet()->setCellValue('E21', floatval($e21));
$objPHPExcel->getActiveSheet()->setCellValue('E22', floatval($e22));


$objPHPExcel->getActiveSheet()->setCellValue('D15', "=SUM(D10:D14)");
$objPHPExcel->getActiveSheet()->setCellValue('E15', "=SUM(E10:E14)");

$objPHPExcel->getActiveSheet()->setCellValue('D23', "=SUM(D18:D22)");
$objPHPExcel->getActiveSheet()->setCellValue('E23', "=SUM(E18:E22)");

$objPHPExcel->getActiveSheet()->setCellValue('D16', "=D15+E15");
$objPHPExcel->getActiveSheet()->setCellValue('D24', "=D23+E23");

$objPHPExcel->getActiveSheet()->setCellValue('D25', "=D16+D24");






$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1B');



######################### END FROM LPS ################################








 



//echo $table_banklain;
//die();

#############################################################################################
/*
--Form 1-DPK--
--Rekening Giro Rupiah--
SELECT COUNT(a.NoRekening)
FROM DM_LiabilitasGiro a
JOIN Master_Range_Suku_Bunga b ON a.DataDate= b.DataDate
WHERE a.DataDate ='2015-10-31' AND a.JenisMataUang='IDR'
and a.Nominal between '0' and '100000000'
--and a.Nominal >'100000000' and a.Nominal<'200000001'
--and a.Nominal >'200000001' and a.Nominal<'500000001'
--and a.Nominal >'500000001' and a.Nominal<'1000000001'
--and a.Nominal >'1000000000' and a.Nominal<'2000000001'
--and a.Nominal>'2000000001' and a.Nominal<'5000000001'
--and a.Nominal > '5000000000'

and a.TingkatSukuBunga <= b.RangeAtas
and a.TingkatSukuBunga= b.RangeAtas and a.TingkatSukuBunga= b.RangeAtas + 1
and a.TingkatSukuBunga= b.RangeAtas + 1 and a.TingkatSukuBunga= b.RangeAtas + 2
and a.TingkatSukuBunga= b.RangeAtas + 2 and a.TingkatSukuBunga= b.RangeAtas + 3
and a.TingkatSukuBunga= b.RangeAtas + 3 and a.TingkatSukuBunga= b.RangeAtas + 4
and a.TingkatSukuBunga > b.RangeBawah
*/
###########################################################################################
#QUERY BATAS ATAS dan BATAS BAWAH
$query_batas=" SELECT * FROM Master_SKB_Efektif where Month(DataDate)='$bulan' and Year(DataDate)='$tahun'  ";
        $res_batas=odbc_exec($connection2, $query_batas);
        $row_batas=odbc_fetch_array($res_batas);
        $Range_Atas=$row_batas['Range_Atas'];
        $Range_Bawah=$row_batas['Range_Bawah'];
        $Range_Valas=$row_batas['Range_Valas'];
//echo $Range_Valas;
//die();
//".number_format((float)$row[Range_Atas], 2, '.', '')."
# Label Range EXCEL
$label_range1=" 0 <= ".number_format((float)$Range_Bawah, 2, '.', '')." ";
$label_range2= number_format((float)$Range_Bawah, 2, '.', '')." < x <= ".number_format((float)$Range_Bawah+1, 2, '.', '');
$label_range3= number_format((float)$Range_Bawah+1, 2, '.', '')." < x <= ".number_format((float)$Range_Bawah+2, 2, '.', '');
$label_range4= number_format((float)$Range_Bawah+2, 2, '.', '')." < x <= ".number_format((float)$Range_Bawah+3, 2, '.', '');
$label_range5= number_format((float)$Range_Bawah+3, 2, '.', '')." < x <= ".number_format((float)$Range_Bawah+4, 2, '.', '');
$label_range6=" x > ".number_format((float)$Range_Atas, 2, '.', '');


$label_valas1=" x <= ".number_format((float)$Range_Valas, 2, '.', '')." ";
$label_valas2=" x > ".number_format((float)$Range_Valas, 2, '.', '')." ";

//echo $label_valas1."<br>";
//echo $label_valas2;
//die();

$Range1=" and a.TingkatSukuBunga <= $Range_Bawah ";
$Range2=" and a.TingkatSukuBunga > $Range_Bawah and a.TingkatSukuBunga <= ".($Range_Bawah+1)." ";
$Range3=" and a.TingkatSukuBunga > ".($Range_Bawah+1)." and a.TingkatSukuBunga <= ".($Range_Bawah+2)." ";
$Range4=" and a.TingkatSukuBunga > ".($Range_Bawah+2)." and a.TingkatSukuBunga <= ".($Range_Bawah+3)." ";
$Range5=" and a.TingkatSukuBunga > ".($Range_Bawah+3)." and a.TingkatSukuBunga <= ".($Range_Bawah+4)." ";
$Range6=" and a.TingkatSukuBunga > $Range_Atas ";






###########################################################################################
//a.Nominal between '0' and '100000000'
###########################################################################################


$query_1a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range1; 

        $result2=odbc_exec($connection2, $query_1a);
        $row2=odbc_fetch_array($result2);
        $c10=$row2['Jml_Rekening'];

//echo $query_1a;
//die ();

$query_1b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and   a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range2;
        $result2=odbc_exec($connection2, $query_1b);
        $row2=odbc_fetch_array($result2);
        $c11=$row2['Jml_Rekening'];

$query_1c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and   a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range3;
        $result2=odbc_exec($connection2, $query_1c);
        $row2=odbc_fetch_array($result2);
        $c12=$row2['Jml_Rekening'];

$query_1d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and   a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range4;
        $result2=odbc_exec($connection2, $query_1d);
        $row2=odbc_fetch_array($result2);
        $c13=$row2['Jml_Rekening'];

$query_1e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and   a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range5;
        $result2=odbc_exec($connection2, $query_1e);
        $row2=odbc_fetch_array($result2);
        $c14=$row2['Jml_Rekening'];

$query_1f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and   a.JumlahBulanLaporan >= '0' and a.JumlahBulanLaporan < '100000001' ".$Range6;
        $result2=odbc_exec($connection2, $query_1f);
        $row2=odbc_fetch_array($result2);
        $c15=$row2['Jml_Rekening'];
##########################################################################################################
# and a.Nominal >'100000000' and a.Nominal<'200000001'
##########################################################################################################

$query_2a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range1; 
        $result2=odbc_exec($connection2, $query_2a);
        $row2=odbc_fetch_array($result2);
        $c16=$row2['Jml_Rekening'];

$query_2b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range2; 
        $result2=odbc_exec($connection2, $query_2b);
        $row2=odbc_fetch_array($result2);
        $c17=$row2['Jml_Rekening'];

$query_2c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range3; 
        $result2=odbc_exec($connection2, $query_2c);
        $row2=odbc_fetch_array($result2);
        $c18=$row2['Jml_Rekening'];

$query_2d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range4; 
        $result2=odbc_exec($connection2, $query_2d);
        $row2=odbc_fetch_array($result2);
        $c19=$row2['Jml_Rekening'];

$query_2e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range5; 
        $result2=odbc_exec($connection2, $query_2e);
        $row2=odbc_fetch_array($result2);
        $c20=$row2['Jml_Rekening'];

$query_2f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan <'200000001' ".$Range6; 
        $result2=odbc_exec($connection2, $query_2f);
        $row2=odbc_fetch_array($result2);
        $c21=$row2['Jml_Rekening'];

##########################################################################################################
# and  a.JumlahBulanLaporan  >'200000001' and  a.JumlahBulanLaporan <'500000001'
##########################################################################################################

$query_3a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range1; 
        $result2=odbc_exec($connection2, $query_3a);
        $row2=odbc_fetch_array($result2);
        $c22=$row2['Jml_Rekening'];

$query_3b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range2; 
        $result2=odbc_exec($connection2, $query_3b);
        $row2=odbc_fetch_array($result2);
        $c23=$row2['Jml_Rekening'];
//echo $query_3b;
//die();
$query_3c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range3; 
        $result2=odbc_exec($connection2, $query_3c);
        $row2=odbc_fetch_array($result2);
        $c24=$row2['Jml_Rekening'];

$query_3d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range4; 
        $result2=odbc_exec($connection2, $query_3d);
        $row2=odbc_fetch_array($result2);
        $c25=$row2['Jml_Rekening'];

$query_3e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range5; 
        $result2=odbc_exec($connection2, $query_3e);
        $row2=odbc_fetch_array($result2);
        $c26=$row2['Jml_Rekening'];

$query_3f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan <'500000001' ".$Range6; 
        $result2=odbc_exec($connection2, $query_3f);
        $row2=odbc_fetch_array($result2);
        $c27=$row2['Jml_Rekening'];        

###################################################################################################################
#--and  a.JumlahBulanLaporan  >'500000001' and  a.JumlahBulanLaporan <'1000000001'
###################################################################################################################
$query_3a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range1; 
        $result2=odbc_exec($connection2, $query_3a);
        $row2=odbc_fetch_array($result2);
        $c28=$row2['Jml_Rekening'];

$query_3b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range2; 
        $result2=odbc_exec($connection2, $query_3b);
        $row2=odbc_fetch_array($result2);
        $c29=$row2['Jml_Rekening'];

$query_3c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range3; 
        $result2=odbc_exec($connection2, $query_3c);
        $row2=odbc_fetch_array($result2);
        $c30=$row2['Jml_Rekening'];

$query_3d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range4; 
        $result2=odbc_exec($connection2, $query_3d);
        $row2=odbc_fetch_array($result2);
        $c31=$row2['Jml_Rekening'];

$query_3e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range5; 
        $result2=odbc_exec($connection2, $query_3e);
        $row2=odbc_fetch_array($result2);
        $c32=$row2['Jml_Rekening'];

$query_3f=" SELECT COUNT(a.NoRekening) as Jml_Rekening  
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan <'1000000001' ".$Range6; 
        $result2=odbc_exec($connection2, $query_3f);
        $row2=odbc_fetch_array($result2);
        $c33=$row2['Jml_Rekening'];
###################################################################################################################        
#--and  a.JumlahBulanLaporan  >'1000000000' and  a.JumlahBulanLaporan <'2000000001'
###################################################################################################################
$query_3a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range1; 
        $result2=odbc_exec($connection2, $query_3a);
        $row2=odbc_fetch_array($result2);
        $c34=$row2['Jml_Rekening'];

$query_3b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range2; 
        $result2=odbc_exec($connection2, $query_3b);
        $row2=odbc_fetch_array($result2);
        $c35=$row2['Jml_Rekening'];

$query_3c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range3; 
        $result2=odbc_exec($connection2, $query_3c);
        $row2=odbc_fetch_array($result2);
        $c36=$row2['Jml_Rekening'];

$query_3d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range4; 
        $result2=odbc_exec($connection2, $query_3d);
        $row2=odbc_fetch_array($result2);
        $c37=$row2['Jml_Rekening'];

$query_3e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range5; 
        $result2=odbc_exec($connection2, $query_3e);
        $row2=odbc_fetch_array($result2);
        $c38=$row2['Jml_Rekening'];

$query_3f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
 a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan <'2000000001' ".$Range6; 
        $result2=odbc_exec($connection2, $query_3f);
        $row2=odbc_fetch_array($result2);
        $c39=$row2['Jml_Rekening'];

###################################################################################################################
#--and  a.JumlahBulanLaporan  >'2000000001' and  a.JumlahBulanLaporan <'5000000001'
###################################################################################################################
$query_3a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
 a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range1; 
        $result2=odbc_exec($connection2, $query_3a);
        $row2=odbc_fetch_array($result2);
        $c40=$row2['Jml_Rekening'];

$query_3b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
 a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range2; 
        $result2=odbc_exec($connection2, $query_3b);
        $row2=odbc_fetch_array($result2);
        $c41=$row2['Jml_Rekening'];

$query_3c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range3; 
        $result2=odbc_exec($connection2, $query_3c);
        $row2=odbc_fetch_array($result2);
        $c42=$row2['Jml_Rekening'];

$query_3d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range4; 
        $result2=odbc_exec($connection2, $query_3d);
        $row2=odbc_fetch_array($result2);
        $c43=$row2['Jml_Rekening'];

$query_3e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range5; 
        $result2=odbc_exec($connection2, $query_3e);
        $row2=odbc_fetch_array($result2);
        $c44=$row2['Jml_Rekening'];

$query_3f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan <'5000000001' ".$Range6; 
        $result2=odbc_exec($connection2, $query_3f);
        $row2=odbc_fetch_array($result2);
        $c45=$row2['Jml_Rekening'];
###################################################################################################################        
#--and  a.JumlahBulanLaporan  >'5000000000'
###################################################################################################################

$query_3a=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range1; 
        $result2=odbc_exec($connection2, $query_3a);
        $row2=odbc_fetch_array($result2);
        $c46=$row2['Jml_Rekening'];

$query_3b=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range2; 
        $result2=odbc_exec($connection2, $query_3b);
        $row2=odbc_fetch_array($result2);
        $c47=$row2['Jml_Rekening'];

$query_3c=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range3; 
        $result2=odbc_exec($connection2, $query_3c);
        $row2=odbc_fetch_array($result2);
        $c48=$row2['Jml_Rekening'];

$query_3d=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range4; 
        $result2=odbc_exec($connection2, $query_3d);
        $row2=odbc_fetch_array($result2);
        $c49=$row2['Jml_Rekening'];

$query_3e=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range5; 
        $result2=odbc_exec($connection2, $query_3e);
        $row2=odbc_fetch_array($result2);
        $c50=$row2['Jml_Rekening'];

$query_3f=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE --b.SBELPS_Level_2 ='SBELPS101000000' AND 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' and  a.JumlahBulanLaporan  >'5000000000' ".$Range6; 
        $result2=odbc_exec($connection2, $query_3f);
        $row2=odbc_fetch_array($result2);
        $c51=$row2['Jml_Rekening'];
######################################  ARRAY VARIABLE ######################################################################

 /*       
$array_range=array("and a.Nominal between '0' and '100000000'","and a.Nominal >'100000000' and a.Nominal<'200000001'","and a.Nominal >'200000001' and a.Nominal<'500000001'","and a.Nominal >'500000001' and a.Nominal<'1000000001'","and a.Nominal >'1000000000' and a.Nominal<'2000000001'","and a.Nominal >'2000000001' and a.Nominal<'5000000001'","and a.Nominal > '5000000000'");
*/
$array_range=array("and  a.JumlahBulanLaporan  >='0' and  a.JumlahBulanLaporan  <'100000001'","and  a.JumlahBulanLaporan  >='100000001' and  a.JumlahBulanLaporan  <'200000001'","and  a.JumlahBulanLaporan  >='200000001' and  a.JumlahBulanLaporan  <'500000001'","and  a.JumlahBulanLaporan  >='500000001' and  a.JumlahBulanLaporan  <'1000000001'","and  a.JumlahBulanLaporan  >='1000000001' and  a.JumlahBulanLaporan  <'2000000001'","and  a.JumlahBulanLaporan  >='2000000001' and  a.JumlahBulanLaporan  <'5000000001'","and  a.JumlahBulanLaporan  > '5000000000'");

$array_range2=array("$Range1", "$Range2", "$Range3","$Range4","$Range5","$Range6");
$array_range3=array("and a.TingkatSukuBunga <= $Range_Valas ","and a.TingkatSukuBunga > $Range_Valas ");

$label_batas=array("$label_range1", "$label_range2", "$label_range3","$label_range4","$label_range5","$label_range6");

#############################################################################################################################
//---------------------------------------------------------------------------------------------------------------------------------------------------------------

//--Nominal Giro Rupiah--
$giro_nominal=array();
$query =" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  a.JenisMataUang='IDR'  ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {
//echo $query.$array_range["$counter"].$array_range2["$counter2"];
//die();
        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_nominal,$row2['Jml_Nominal']);


    $counter2++;   
}
    
    $counter++;

}





//var_dump($giro_nominal);
//die();
//$label_range1
#--Rekening Tabungan Rupiah--
$tabungan_rekening=array();
$query=" SELECT  COUNT(a.NoRekening) as Jml_Rekening
FROM $table_tabungan a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {
//echo $query."'".$label_batas["$counter2"]."'".$query_add.$array_range["$counter"].$array_range2["$counter2"]." GROUP BY Managed_GL_Code,Managed_GL_Prod_Code, KodeCabang   "."<br>";
//die();
         $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
         $row2=odbc_fetch_array($result2);
         array_push($tabungan_rekening,$row2['Jml_Rekening']);

    $counter2++;   
}
   
    $counter++;

}





/*
$tabungan_rekening=array();
$query=" SELECT  ";
$query_add="  as Batas, Managed_GL_Code,Managed_GL_Prod_Code, KodeCabang, COUNT(a.NoRekening) as Jml_Rekening
FROM $table_tabungan a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {
echo $query."'".$label_batas["$counter2"]."'".$query_add.$array_range["$counter"].$array_range2["$counter2"]." GROUP BY Managed_GL_Code,Managed_GL_Prod_Code, KodeCabang  UNION ALL "."<br>";
//die();
        // $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        // $row2=odbc_fetch_array($result2);
        // array_push($tabungan_rekening,$row2['Jml_Rekening']);

    $counter2++;   
}
   
    $counter++;

}
*/
//echo "============================================================================";
//die();
#--Nominal Tabungan Rupiah--
$tabungan_nominal=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_tabungan a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

//echo $query."'".$label_batas["$counter2"]."'".$array_range["$counter"].$array_range2["$counter2"]." GROUP BY Managed_GL_Code,Managed_GL_Prod_Code, KodeCabang  UNION ALL "."<br>";
//echo $query.$array_range["$counter"].$array_range2["$counter2"]."<br>";


        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_nominal,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}

//die();

//=============== deposito
#--Rekening Deposito Rupiah--
$deposito_rekening=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

    //echo $query.$array_range["$counter"].$array_range2["$counter2"]."<br>" ;

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rekening,$row2['Jml_Rekening']);

    $counter2++;   
}
   
    $counter++;

}
#--Nominal Deposito Rupiah--
$deposito_nominal=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

//echo $query.$array_range["$counter"].$array_range2["$counter2"]."<br>"; 
        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_nominal,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}

//=============== deposito on call
#--Rekening Deposito on call Rupiah--
$deposito_call_rekening=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' AND a.status_oncall='Y' ";

//echo $query;
//die();

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rekening,$row2['Jml_Rekening']);

    $counter2++;   
}
   
    $counter++;

}
#--Nominal Deposito on call Rupiah--
$deposito_call_nominal=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.status_oncall='Y'";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_nominal,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}




#####################################################  FORM2 ####################################################
#--Form 2-Liabilitas Bank Lain--
#--Rekening Giro Rupiah--
$giro_rekening2=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y' 
AND a.FlagDPK='Giro' ";


$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_rekening2,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}





#--Nominal Giro Rupiah--
$giro_nominal2=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.FlagDPK='Giro' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_nominal2,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}




#--Rekening Tabungan Rupiah--
$tabungan_rekening2=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y'
AND a.FlagDPK='Tabungan' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_rekening2,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}



#--Nominal Tabungan Rupiah--
$tabungan_nominal2=array();
$query=" SELECT SUM(a.JumlahBulanLaporan)  as Jml_Nominal
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR'
AND a.FlagDPK='Tabungan' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_nominal2,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}


#--Deposito Rupiah-- rekening
$deposito_rekening2=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND a.StatusLaporan_LBU='Y'
AND a.FlagDPK='Deposito' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rekening2,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}
//echo "<pre>";
//echo $deposito_rekening2;
//echo "</pre>";
//die();

#--Nominal Deposito Rupiah--
$deposito_nominal2=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' 
AND a.FlagDPK='Deposito' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_nominal2,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}
//echo $query.$array_range["$counter"].$array_range2["$counter2"];
//echo var_dump($deposito_nominal2);
//die();

#--Deposito on Call Rupiah-- rekening
$deposito_call_rekening2=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang='IDR' AND 
AND a.FlagDPK='Deposito' AND a.status_oncall='Y' AND a.StatusLaporan_LBU='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rekening2,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Deposito On call Rupiah--
$deposito_call_nominal2=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang='IDR' AND 
AND a.FlagDPK='Deposito' AND a.status_oncall='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range2 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range2["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_nominal2,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}


#####################################################  FORM3 ####################################################
#--Form 3
#--Rekening Giro Valas--
$giro_rekening3=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_rekening3,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Giro Valas--
$giro_nominal3=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_giro  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang<>'IDR'";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_nominal3,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}




#--Rekening Tabungan Valas--
$tabungan_rekening3=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_tabungan  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_rekening3,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}



#--Nominal Tabungan Valas--
$tabungan_nominal3=array();
$query=" SELECT SUM(a.JumlahBulanLaporan)  as Jml_Nominal
FROM $table_tabungan  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang<>'IDR' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_nominal3,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}


#--Deposito Valas-- rekening
$deposito_rekening3=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_deposito  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rekening3,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Deposito valas--
$deposito_nominal3=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_deposito  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_nominal3,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}

#--Deposito on Call  rekening
$deposito_call_rekening3=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_deposito  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y' AND a.status_oncall='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rekening3,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Deposito on call -
$deposito_call_nominal3=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_deposito  a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.status_oncall='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_nominal3,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}

//

#####################################################  FORM4 ####################################################
#--Form 4-Liabilitas Bank Lain--
#--Rekening Giro Valas--
$giro_rekening4=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y'
AND a.FlagDPK='Giro' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_rekening4,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}





#--Nominal Giro valas--
$giro_nominal4=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' 
AND a.FlagDPK='Giro' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_nominal4,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}




#--Rekening Tabungan valas--
$tabungan_rekening4=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR'  AND a.StatusLaporan_LBU='Y'
AND a.FlagDPK='Tabungan' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_rekening4,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}



#--Nominal Tabungan valas--
$tabungan_nominal4=array();
$query=" SELECT SUM(a.JumlahBulanLaporan)  as Jml_Nominal
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' 
AND a.FlagDPK='Tabungan' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_nominal4,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}


#--Deposito valas-- rekening
$deposito_rekening4=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE  
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y'
AND a.FlagDPK='Deposito' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rekening4,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Deposito valas--
$deposito_nominal4=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' AND a.status_oncall='N' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_nominal4,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}

#--Deposito on call  rekening
$deposito_call_rekening4=array();
$query=" SELECT COUNT(a.NoRekening) as Jml_Rekening
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' AND a.StatusLaporan_LBU='Y' 
AND a.FlagDPK='Deposito' AND a.status_oncall='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rekening4,$row2['Jml_Rekening']);

    $counter2++;   
}
    $counter++;

}


#--Nominal Deposito on call --
$deposito_call_nominal4=array();
$query=" SELECT SUM(a.JumlahBulanLaporan) as Jml_Nominal
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code 
JOIN Master_SKB_Efektif d ON a.DataDate= d.DataDate
WHERE 
a.JenisMataUang<>'IDR' 
AND a.FlagDPK='Deposito' AND a.status_oncall='Y' ";

$counter=0;
foreach ($array_range as $value) {
$counter2=0;
foreach ($array_range3 as $value2) {

        $result2=odbc_exec($connection2, $query.$array_range["$counter"].$array_range3["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_nominal4,$row2['Jml_Nominal']);

    $counter2++;   
}
    $counter++;

}


// AND a.status_oncall='Y'
//=============== sertifikat deposito



// SHEET KE 2 ======================================================================================
// Create new PHPExcel object
//$objPHPExcel = new PHPExcel();
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$objPHPExcel->getActiveSheet()->getStyle('A1:N3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A2:N3')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A8:N52')->applyFromArray($styleArrayBorder1);




//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:Z7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('O1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A53:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);


// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(1);

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A1:N1');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:N2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:N3');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A8:A9');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B8:B9');


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C8:D8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('E8:F8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('G8:H8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I8:J8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K8:L8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('M8:N8');



$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A10:A15');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A16:A21');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A22:A27');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A28:A33');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A34:A39');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A40:A45');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A46:A51');

$objPHPExcel->getActiveSheet()->getStyle('D10:D51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F10:F51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('H10:H51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J10:J51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L10:L51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:N51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C10:C51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E10:E51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G10:G51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I10:I51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('K10:K51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('M10:M51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->setCellValue('A1', 'SUKU BUNGA EFEKTIF');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI SIMPANAN PIHAK KETIGA ");
$objPHPExcel->getActiveSheet()->setCellValue('A3', 'RUPIAH');

$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Nama Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Kode Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Laporan Akhir Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B4', ': PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B5', ': 485');
$objPHPExcel->getActiveSheet()->setCellValue('B6', ": $label_tgl");

#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('A8', 'Tiering');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Suku Bunga Efektif');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Giro');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'Tabungan');
$objPHPExcel->getActiveSheet()->setCellValue('G8', 'Deposito');
$objPHPExcel->getActiveSheet()->setCellValue('I8', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('D9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('E9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('F9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('G9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('H9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('I9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('L9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('M9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "Jumlah Nominal");


$objPHPExcel->getActiveSheet()->setCellValue('A10', "0 < N <= 100 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "100 Jt < N <= 200 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "200 Jt < N <= 500 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A28', "500 Jt < N <= 1M");
$objPHPExcel->getActiveSheet()->setCellValue('A34', "1M < N <= 2M");
$objPHPExcel->getActiveSheet()->setCellValue('A40', "2M < N <= 5M");
$objPHPExcel->getActiveSheet()->setCellValue('A46', "> 5 Milyar ");

$objPHPExcel->getActiveSheet()->setCellValue('B10', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B12', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B14', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C10', $c10);
$objPHPExcel->getActiveSheet()->setCellValue('C11', $c11);
$objPHPExcel->getActiveSheet()->setCellValue('C12', $c12);
$objPHPExcel->getActiveSheet()->setCellValue('C13', $c13);
$objPHPExcel->getActiveSheet()->setCellValue('C14', $c14);
$objPHPExcel->getActiveSheet()->setCellValue('C15', $c15);

$objPHPExcel->getActiveSheet()->setCellValue('B16', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B18', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B20', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B21', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C16', $c16);
$objPHPExcel->getActiveSheet()->setCellValue('C17', $c17);
$objPHPExcel->getActiveSheet()->setCellValue('C18', $c18);
$objPHPExcel->getActiveSheet()->setCellValue('C19', $c19);
$objPHPExcel->getActiveSheet()->setCellValue('C20', $c20);
$objPHPExcel->getActiveSheet()->setCellValue('C21', $c21);

$objPHPExcel->getActiveSheet()->setCellValue('B22', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B23', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B24', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B25', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B26', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B27', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C22', $c22);
$objPHPExcel->getActiveSheet()->setCellValue('C23', $c23);
$objPHPExcel->getActiveSheet()->setCellValue('C24', $c24);
$objPHPExcel->getActiveSheet()->setCellValue('C25', $c25);
$objPHPExcel->getActiveSheet()->setCellValue('C26', $c26);
$objPHPExcel->getActiveSheet()->setCellValue('C27', $c27);

$objPHPExcel->getActiveSheet()->setCellValue('B28', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B29', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B30', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B31', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B32', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B33', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C28', $c28);
$objPHPExcel->getActiveSheet()->setCellValue('C29', $c29);
$objPHPExcel->getActiveSheet()->setCellValue('C30', $c30);
$objPHPExcel->getActiveSheet()->setCellValue('C31', $c31);
$objPHPExcel->getActiveSheet()->setCellValue('C32', $c32);
$objPHPExcel->getActiveSheet()->setCellValue('C33', $c33);

$objPHPExcel->getActiveSheet()->setCellValue('B34', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B35', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B36', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B37', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B38', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B39', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C34', $c34);
$objPHPExcel->getActiveSheet()->setCellValue('C35', $c35);
$objPHPExcel->getActiveSheet()->setCellValue('C36', $c36);
$objPHPExcel->getActiveSheet()->setCellValue('C37', $c37);
$objPHPExcel->getActiveSheet()->setCellValue('C38', $c38);
$objPHPExcel->getActiveSheet()->setCellValue('C39', $c39);

$objPHPExcel->getActiveSheet()->setCellValue('B40', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B41', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B42', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B43', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B44', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B45', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C40', $c40);
$objPHPExcel->getActiveSheet()->setCellValue('C41', $c41);
$objPHPExcel->getActiveSheet()->setCellValue('C42', $c42);
$objPHPExcel->getActiveSheet()->setCellValue('C43', $c43);
$objPHPExcel->getActiveSheet()->setCellValue('C44', $c44);
$objPHPExcel->getActiveSheet()->setCellValue('C45', $c45);

$objPHPExcel->getActiveSheet()->setCellValue('B46', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B47', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B48', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B49', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B50', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B51', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('C46', $c46);
$objPHPExcel->getActiveSheet()->setCellValue('C47', $c47);
$objPHPExcel->getActiveSheet()->setCellValue('C48', $c48);
$objPHPExcel->getActiveSheet()->setCellValue('C49', $c49);
$objPHPExcel->getActiveSheet()->setCellValue('C50', $c50);
$objPHPExcel->getActiveSheet()->setCellValue('C51', $c51);

# Giro Jumlah Nominal --> Kolom D
$index=10;
foreach ($giro_nominal as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$index", floatval($nilai) );

$index++;

}
# Tabungan Rekening --> Kolom E
$index2=10;
foreach ($tabungan_rekening as $nilai2 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("E$index2", floatval($nilai2) );

$index2++;

}
# Tabungan Nominal --> Kolom F
$index3=10;
foreach ($tabungan_nominal as $nilai3 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$index3", floatval($nilai3) );

$index3++;

}

# Deposito Rekening --> Kolom G
$index4=10;
foreach ($deposito_rekening as $nilai4 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$index4", floatval($nilai4) );

$index4++;

}
# Deposito Nominal --> Kolom H
$index5=10;
foreach ($deposito_nominal as $nilai5 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("H$index5", floatval($nilai5) );

$index5++;

}

# Deposito call Rekening --> Kolom K
$index4=10;
foreach ($deposito_call_rekening as $nilai4 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$index4", floatval($nilai4) );

$index4++;

}
# Deposito Nominal --> Kolom L 
$index5=10;
foreach ($deposito_call_nominal as $nilai5 ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$index5", floatval($nilai5) );

$index5++;

}





# TOTAL Rekening --> Kolom M

for ($i=10; $i < 52 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("M$i", "=C$i+E$i+G$i+I$i+K$i" );    
}

# TOTAL Nominal --> Kolom N
for ($i=10; $i < 52 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("N$i", "=D$i+F$i+H$i+J$i+L$i" );    
}




$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1D');


/*
var_dump($giro_nominal);
echo "<br><br>";
var_dump($tabungan_nominal);
die();
*/



// SHEET KE 3 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2); 
//$objPHPExcel->setActiveSheetIndex(0);

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$objPHPExcel->getActiveSheet()->getStyle('A1:N3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A2:N3')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


    
//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A8:N52')->applyFromArray($styleArrayBorder1);




//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:Z7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('O1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A53:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);


// Create a first sheet, representing sales data


$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A1:N1');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A2:N2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A3:N3');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A8:A9');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B8:B9');


$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C8:D8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('E8:F8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('G8:H8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('I8:J8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('K8:L8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('M8:N8');



$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A10:A15');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A16:A21');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A22:A27');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A28:A33');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A34:A39');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A40:A45');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A46:A51');

$objPHPExcel->getActiveSheet()->getStyle('D10:D51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F10:F51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('H10:H51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J10:J51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L10:L51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:N51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C10:C51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E10:E51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G10:G51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I10:I51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('K10:K51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('M10:M51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->setCellValue('A1', 'SUKU BUNGA EFEKTIF');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI SIMPANAN ANTAR BANK PASSIVA (SIMPANAN DARI BANK LAIN) ");
$objPHPExcel->getActiveSheet()->setCellValue('A3', 'RUPIAH');

$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Nama Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Kode Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Laporan Akhir Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B4', ': PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B5', ': 485');
$objPHPExcel->getActiveSheet()->setCellValue('B6', ": $label_tgl");

#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('A8', 'Tiering');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Suku Bunga Efektif');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Giro');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'Tabungan');
$objPHPExcel->getActiveSheet()->setCellValue('G8', 'Deposito');
$objPHPExcel->getActiveSheet()->setCellValue('I8', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('D9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('E9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('F9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('G9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('H9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('I9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('L9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('M9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "Jumlah Nominal");


$objPHPExcel->getActiveSheet()->setCellValue('A10', "0 < N <= 100 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "100 Jt < N <= 200 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "200 Jt < N <= 500 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A28', "500 Jt < N <= 1M");
$objPHPExcel->getActiveSheet()->setCellValue('A34', "1M < N <= 2M");
$objPHPExcel->getActiveSheet()->setCellValue('A40', "2M < N <= 5M");
$objPHPExcel->getActiveSheet()->setCellValue('A46', "> 5 Milyar ");

$objPHPExcel->getActiveSheet()->setCellValue('B10', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B12', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B14', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B16', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B18', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B20', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B21', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B22', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B23', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B24', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B25', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B26', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B27', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B28', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B29', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B30', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B31', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B32', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B33', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B34', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B35', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B36', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B37', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B38', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B39', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B40', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B41', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B42', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B43', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B44', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B45', $label_range6);

$objPHPExcel->getActiveSheet()->setCellValue('B46', $label_range1);
$objPHPExcel->getActiveSheet()->setCellValue('B47', $label_range2);
$objPHPExcel->getActiveSheet()->setCellValue('B48', $label_range3);
$objPHPExcel->getActiveSheet()->setCellValue('B49', $label_range4);
$objPHPExcel->getActiveSheet()->setCellValue('B50', $label_range5);
$objPHPExcel->getActiveSheet()->setCellValue('B51', $label_range6);



# Giro Jumlah Rekening --> Kolom C
$indexB1=10;
foreach ($giro_rekening2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("C$indexB1", floatval($nilai) );

$indexB1++;

}

# Giro Jumlah Nominal --> Kolom D
$indexB2=10;
foreach ($giro_nominal2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$indexB2", floatval($nilai) );

$indexB2++;

}

# TABUNGAN Jumlah Rekening --> Kolom E
$indexB3=10;
foreach ($tabungan_rekening2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("E$indexB3", floatval($nilai) );

$indexB3++;

}

# TABUNGAN Jumlah Nominal --> Kolom F
$indexB4=10;
foreach ($tabungan_nominal2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$indexB4", floatval($nilai) );

$indexB4++;

}

# deposito Jumlah Rekening --> Kolom G
$indexB5=10;
foreach ($deposito_rekening2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$indexB5", floatval($nilai) );

$indexB5++;

}

# deposito Jumlah Nominal --> Kolom H
$indexB6=10;
foreach ($deposito_nominal2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("H$indexB6", floatval($nilai) );

$indexB6++;

}


# deposito call Jumlah Rekening --> Kolom K
$indexB5=10;
foreach ($deposito_call_rekening2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$indexB5", floatval($nilai) );

$indexB5++;

}

# deposito call Jumlah Nominal --> Kolom L
$indexB6=10;
foreach ($deposito_call_nominal2 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$indexB6", floatval($nilai) );

$indexB6++;

}




# TOTAL Rekening --> Kolom M

for ($i=10; $i < 52 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("M$i", "=C$i+E$i+G$i+I$i+K$i" );    
}

# TOTAL Nominal --> Kolom N
for ($i=10; $i < 52 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("N$i", "=D$i+F$i+H$i+J$i+L$i" );    
}

$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1E');

#
#====================== 
// SHEET KE 4 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3); 
//$objPHPExcel->setActiveSheetIndex(0);

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$objPHPExcel->getActiveSheet()->getStyle('A1:N3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A2:N3')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A8:N24')->applyFromArray($styleArrayBorder1);




//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:Z7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('O1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A25:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A24:N24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);


// Create a first sheet, representing sales data


$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A1:N1');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A2:N2');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A3:N3');

$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A8:A9');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('B8:B9');


$objPHPExcel->setActiveSheetIndex(3)->mergeCells('C8:D8');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('E8:F8');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('G8:H8');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('I8:J8');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('K8:L8');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('M8:N8');



$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A10:A11');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A12:A13');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A14:A15');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A16:A17');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A18:A19');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A20:A21');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A22:A23');

$objPHPExcel->getActiveSheet()->getStyle('D10:D23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F10:F23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('H10:H23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J10:J23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L10:L23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:N23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C10:C23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E10:E23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G10:G23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I10:I23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('K10:K23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('M10:M23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('A1', 'SUKU BUNGA EFEKTIF');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI DANA PIHAK KETIGA");
$objPHPExcel->getActiveSheet()->setCellValue('A3', 'VALUTA ASING');

$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Nama Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Kode Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Laporan Akhir Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B4', ': PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B5', ': 485');
$objPHPExcel->getActiveSheet()->setCellValue('B6', ": $label_tgl");

#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('A8', 'Tiering');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Suku Bunga Efektif');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Giro');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'Tabungan');
$objPHPExcel->getActiveSheet()->setCellValue('G8', 'Deposito');
$objPHPExcel->getActiveSheet()->setCellValue('I8', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('D9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('E9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('F9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('G9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('H9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('I9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('L9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('M9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "Jumlah Nominal");



$objPHPExcel->getActiveSheet()->setCellValue('A10', "0 < N <= 100 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "100 Jt < N <= 200 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "200 Jt < N <= 500 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "500 Jt < N <= 1M");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "1M < N <= 2M");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "2M < N <= 5M");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "> 5 Milyar ");


$objPHPExcel->getActiveSheet()->setCellValue('B10', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B12', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B14', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B16', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B18', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B20', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B21', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B22', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B23', $label_valas2);

# Giro Jumlah Rekening --> Kolom C
$indexC1=10;
foreach ($giro_rekening3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("C$indexC1", floatval($nilai) );

$indexC1++;

}

# Giro Jumlah Nominal --> Kolom D
$indexC2=10;
foreach ($giro_nominal3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$indexC2", floatval($nilai) );

$indexC2++;

}

# TABUNGAN Jumlah Rekening --> Kolom E
$indexC3=10;
foreach ($tabungan_rekening3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("E$indexC3", floatval($nilai) );

$indexC3++;

}

# TABUNGAN Jumlah Nominal --> Kolom F
$indexC4=10;
foreach ($tabungan_nominal3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$indexC4", floatval($nilai) );

$indexC4++;

}

# deposito Jumlah Rekening --> Kolom G
$indexC5=10;
foreach ($deposito_rekening3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$indexC5", floatval($nilai) );

$indexC5++;

}

# deposito Jumlah Nominal --> Kolom H
$indexC6=10;
foreach ($deposito_nominal3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("H$indexC6", floatval($nilai) );

$indexC6++;

}

# deposito on call Jumlah Rekening --> Kolom K
$indexC5=10;
foreach ($deposito_call_rekening3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$indexC5", floatval($nilai) );

$indexC5++;

}

# deposito on call Jumlah Nominal --> Kolom L
$indexC6=10;
foreach ($deposito_call_nominal3 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$indexC6", floatval($nilai) );

$indexC6++;

}






# TOTAL Rekening --> Kolom M

for ($i=10; $i < 24 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("M$i", "=C$i+E$i+G$i+I$i+K$i" );    
}

# TOTAL Nominal --> Kolom N
for ($i=10; $i < 24 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("N$i", "=D$i+F$i+H$i+J$i+L$i" );    
}

$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1F');
#====================== 
// SHEET KE 5 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(4); 
//$objPHPExcel->setActiveSheetIndex(0);

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$objPHPExcel->getActiveSheet()->getStyle('A1:N3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A2:N3')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A10:A15')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A16:A21')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A22:A27')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:A33')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A34:A39')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A40:A45')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A46:A51')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A8:N24')->applyFromArray($styleArrayBorder1);




//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:Z7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('O1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A25:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A24:N24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);


// Create a first sheet, representing sales data


$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A1:N1');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A2:N2');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A3:N3');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A8:A9');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B8:B9');


$objPHPExcel->setActiveSheetIndex(4)->mergeCells('C8:D8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('E8:F8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('G8:H8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('I8:J8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('K8:L8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('M8:N8');



$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A10:A11');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A12:A13');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A14:A15');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A16:A17');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A18:A19');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A20:A21');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A22:A23');

$objPHPExcel->getActiveSheet()->getStyle('D10:D23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F10:F23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('H10:H23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J10:J23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L10:L23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:N23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C10:C23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E10:E23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G10:G23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I10:I23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('K10:K23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('M10:M23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('A1', 'SUKU BUNGA EFEKTIF');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI SIMPANAN ANTAR BANK PASSIVA (SIMPANAN DARI BANK LAIN) ");
$objPHPExcel->getActiveSheet()->setCellValue('A3', 'VALUTA ASING');

$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Nama Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Kode Bank');
$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Laporan Akhir Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B4', ': PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B5', ': 485');
$objPHPExcel->getActiveSheet()->setCellValue('B6', ": $label_tgl");

#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('A8', 'Tiering');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Suku Bunga Efektif');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Giro');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'Tabungan');
$objPHPExcel->getActiveSheet()->setCellValue('G8', 'Deposito');
$objPHPExcel->getActiveSheet()->setCellValue('I8', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('D9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('E9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('F9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('G9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('H9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('I9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K9', 'Jumlah Rekening');
$objPHPExcel->getActiveSheet()->setCellValue('L9', 'Jumlah Nominal');
$objPHPExcel->getActiveSheet()->setCellValue('M9', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "Jumlah Nominal");



$objPHPExcel->getActiveSheet()->setCellValue('A10', "0 < N <= 100 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "100 Jt < N <= 200 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "200 Jt < N <= 500 Jt");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "500 Jt < N <= 1M");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "1M < N <= 2M");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "2M < N <= 5M");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "> 5 Milyar ");





$objPHPExcel->getActiveSheet()->setCellValue('B10', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B12', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B14', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B16', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B18', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B20', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B21', $label_valas2);

$objPHPExcel->getActiveSheet()->setCellValue('B22', $label_valas1);
$objPHPExcel->getActiveSheet()->setCellValue('B23', $label_valas2);


# Giro Jumlah Rekening --> Kolom C
$indexD1=10;
foreach ($giro_rekening4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("C$indexD1", floatval($nilai) );

$indexD1++;

}

# Giro Jumlah Nominal --> Kolom D
$indexD2=10;
foreach ($giro_nominal4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$indexD2", floatval($nilai) );

$indexD2++;

}

# TABUNGAN Jumlah Rekening --> Kolom E
$indexD3=10;
foreach ($tabungan_rekening4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("E$indexD3", floatval($nilai) );

$indexD3++;

}

# TABUNGAN Jumlah Nominal --> Kolom F
$indexD4=10;
foreach ($tabungan_nominal4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$indexD4", floatval($nilai) );

$indexD4++;

}

# deposito Jumlah Rekening --> Kolom G
$indexD5=10;
foreach ($deposito_rekening4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$indexD5", floatval($nilai) );

$indexD5++;

}

# deposito Jumlah Nominal --> Kolom H
$indexD6=10;
foreach ($deposito_nominal4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("H$indexD6", floatval($nilai) );

$indexD6++;

}


# deposito on call Jumlah Rekening --> Kolom K
$indexD5=10;
foreach ($deposito_call_rekening4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$indexD5", floatval($nilai) );

$indexD5++;

}

# deposito on call Jumlah Nominal --> Kolom L
$indexD6=10;
foreach ($deposito_call_nominal4 as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$indexD6", floatval($nilai) );

$indexD6++;

}





# TOTAL Rekening --> Kolom M

for ($i=10; $i < 24 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("M$i", "=C$i+E$i+G$i+I$i+K$i" );    
}

# TOTAL Nominal --> Kolom N
for ($i=10; $i < 24 ; $i++) { 
    $objPHPExcel->getActiveSheet()->setCellValue("N$i", "=D$i+F$i+H$i+J$i+L$i" );    
}

$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1G');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/SKB_EFEKTIF_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/SKB_EFEKTIF_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);



//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save("download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls");





?>

<div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> List Laporan Posisi Simpanan Bank Umum
                            </div>
                           </div>
                        <div class="portlet-body">
                            <h4 ><b>PT Bank MNC Internasional, Tbk</b></h4>
                            
                            
                            <div class="tabbable-line">
                                <ul class="nav nav-tabs ">
                                   <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        Lampiran 1B </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        Lampiran 1D </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_3" data-toggle="tab">
                                        Lampiran 1E </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_4" data-toggle="tab">
                                        Lampiran 1F </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_5" data-toggle="tab">
                                        Lampiran 1G </a>
                                    </li>
                                </ul>
                                <div class="tab-content">

                                <!--  new -->
                                <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/SKB_EFEKTIF_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div></b> 

                                            
</br>
</br>

                                        <p>
                                        <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> LAMPIRAN 1 B</b> <br>
                                       <b> LAPORAN SIMPANAN <?php echo $label_tgl; ?> </b>
                                      
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="2" ><b>No</b></td>
                                                <td width="40%" align="center"  rowspan="2" ><b>Bentuk Simpanan </b></td>
                                                <td width="20%" align="center" ><b>Rupiah </b></td>
                                                <td width="25%" align="center" ><b> Valuta Asing (Ekuivalen Rupiah)  </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="20%" align="center" ><b> (i) </b></td>
                                                <td width="20%" align="center" ><b> (ii) </b></td>
                                                
                                               
                                                </tr>
                                                
                                                </thead>
                                                <tbody>

                                                <tr class="danger">
                                                <td align="center" > A.</td>
                                                <td align="left">Simpanan Pihak Ketiga</td>
                                                <td align="right"></td>
                                                <td align="right"></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 1</td>
                                                <td align="left">Giro</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 2</td>
                                                <td align="left">Tabungan</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 3</td>
                                                <td align="left">Deposit on Call (DOC)</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 4</td>
                                                <td align="left">Deposito</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 5</td>
                                                <td align="left">Sertifikat Deposito</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >Sub Total Simpanan Pihak Ketiga</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >A. Total Simpanan Pihak Ketiga Dalam Rupiah (i) + (ii)</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"></td>
                                                </tr>
                                                <tr class="danger">
                                                <td align="center"> B.</td>
                                                <td align="left">Simpanan dari bank Lain</td>
                                                <td align="right"></td>
                                                <td align="right"></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 6</td>
                                                <td align="left">Giro</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 7</td>
                                                <td align="left">Tabungan</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 8</td>
                                                <td align="left">Deposit on Call (DOC)</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 9</td>
                                                <td align="left">Deposito</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 10</td>
                                                <td align="left">Sertifikat Deposito</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >Sub Total Simpanan Dari bank Lain</td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >B. Total Simpanan dari bank Lain Dalam Rupiah (i) + (ii)</td>
                                                 <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >Total Dalam Rupiah</td>
                                                 <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        Cabang di Luar Negeri

                                        
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="2" ><b>No</b></td>
                                                <td width="40%" align="center"  rowspan="2" ><b>Bentuk Simpanan </b></td>
                                                <td width="20%" align="center" ><b>Rupiah </b></td>
                                                <td width="25%" align="center" ><b> Valuta Asing (Ekuivalen Rupiah)  </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="20%" align="center" ><b> (i) </b></td>
                                                <td width="20%" align="center" ><b> (ii) </b></td>
                                                
                                               
                                                </tr>
                                                
                                                </thead>
                                                <tbody>

                                                <tr class="danger">
                                                <td align="center" > C.</td>
                                                <td align="left">Simpanan Cabang Di Luar Negeri</td>
                                                <td align="right"></td>
                                                <td align="right"></td>
                                                </tr>
                                                <tr>
                                                <td align="center"> 11</td>
                                                <td align="left">Simpanan Cabang Di Luar Negeri</td>
                                                <td align="right"></td>
                                                <td align="right"></td>
                                                </tr>
                                                
                                                <tr>
                                                <td align="center"> </td>
                                                <td align="center" >C. Total Simpanan Cabang Di Luar Negeri Dalam Rupiah</td>
                                                <td align="right"></td>
                                                <td align="right"></td>
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </div>
                                <!--  end new -->


                                    <div class="tab-pane" id="tab_15_2">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/SKB_EFEKTIF_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div></b> 

                                            
</br>
</br>

                                        <p>
                                        <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 1</b> <br>
                                       <b> RINCIAN POSISI SIMPANAN PIHAK KETIGA </b> 
                                       <b> RUPIAH </b>
                                    </div>
                                    <?php $objPHPExcel->setActiveSheetIndex(1);?>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b>Giro </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Tabungan </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Sertifikat Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito on Call </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Total </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Tiering</b></td>
                                                 <td width="10%" align="center" ><b>Suku Bunga Efektif</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td align="center" rowspan="6" >0 < N <= 100 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >100 < N <= 200 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >200 < N <= 500 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >500 Jt < N <= 1M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >1M < N <= 2M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >2M < N <= 5M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >> 5 Milyar </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        </div>

                                        <div class="tab-pane" id="tab_15_3">
                                         <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 2</b><br>
                                       <b> RINCIAN POSISI SIMPANAN ANTAR BANK PASSIVA (SIMPANAN DARI BANK LAIN) </b> 
                                       <b> RUPIAH </b>
                                       <?php $objPHPExcel->setActiveSheetIndex(2);?>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b>Giro </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Tabungan </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Sertifikat Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito on Call </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Total </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Tiering</b></td>
                                                 <td width="10%" align="center" ><b>Suku Bunga Efektif</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td align="center" rowspan="6" >0 < N <= 100 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >100 < N <= 200 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >200 < N <= 500 Jt</td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N24")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N25")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N27")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >500 Jt < N <= 1M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N29")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N31")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N32")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >1M < N <= 2M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N34")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N35")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N37")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N38")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >2M < N <= 5M </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N40")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N41")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N42")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N43")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N44")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N45")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="6" >> 5 Milyar </td>
                                                <td align="left"><?php echo $label_range1;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N46")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range2;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N47")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range3;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N48")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range4;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $label_range5;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N50")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left">x ><?php echo $label_range6;?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N51")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        </div>
                                        <div class="tab-pane" id="tab_15_4">
                                          <div class="alert alert-info">
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 3 </b><br>
                                       <b> RINCIAN POSISI DANA PIHAK KETIGA (VALUTA ASING) </b><br>
                                       <?php $objPHPExcel->setActiveSheetIndex(3);?>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="1400">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b>Giro </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Tabungan </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Sertifikat Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito on Call </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Total </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Tiering</b></td>
                                                 <td width="10%" align="center" ><b>Suku Bunga Efektif</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td align="center" rowspan="2" >0 < N <= 100 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >100 Jt < N <= 200 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >200 Jt < N <= 500 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >500 Jt < N <= 1M </td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >1M < N <= 2M</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >2M < N <= 5M</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >> 5 Milyar</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </div>

                                         <div class="tab-pane" id="tab_15_5">
                                          <div class="alert alert-info">
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 4 </b><br>
                                       <b> RINCIAN POSISI SIMPANAN ANTAR BANK PASSIVA (SIMPANAN DARI BANK LAIN) (VALUTA ASING) </b><br>
                                        <?php $objPHPExcel->setActiveSheetIndex(4);?>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="1400">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="10%" align="center" ><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b>Giro </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Tabungan </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Sertifikat Deposito </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Deposito on Call </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Total </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Tiering</b></td>
                                                 <td width="10%" align="center" ><b>Suku Bunga Efektif</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Rekening</b></td>
                                                 <td width="10%" align="center" ><b>Jml. Nominal</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td align="center" rowspan="2" >0 < N <= 100 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N10")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N11")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >100 Jt < N <= 200 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N12")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N13")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >200 Jt < N <= 500 Jt</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N14")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N15")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >500 Jt < N <= 1M </td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N17")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >1M < N <= 2M</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N18")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N19")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >2M < N <= 5M</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N20")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="center" rowspan="2" >> 5 Milyar</td>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N22")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </div>


                                    </div>
                                  
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

