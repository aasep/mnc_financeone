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
logActivity("generate nop",date('Y_m_d_H_i_s'));
/*
$tanggal=$_POST['tanggal']; 
$curr_tgl=date('Y-m-d',strtotime($tanggal));

$label_txtfile=date('Ymd',strtotime($tanggal));
$tanggal_header=date('dmY',strtotime($tanggal));


$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih
*/
$tahun=$_POST['tahun'];
$kuartal=$_POST['kuartal'];
//echo $kuartal;
//die();
switch ($kuartal) {
        case '1':
        $tanggal=$tahun."-03-31";
        break;
        case '2':
        $tanggal=$tahun."-06-30";
        break;
        case '3':
        $tanggal=$tahun."-09-30";
        break;
        case '4':
        $tanggal=$tahun."-12-31";
        break;
     
}
$var_tgl=date('Y-m-d',strtotime($tanggal));

#--Laporan A--
#--Pihak Penyalur--
$query = " SELECT distinct b.NamaLengkapDebitur
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN DM_InfoDebitur b ON b.NomorNasabah = a.NomorNasabah 
JOIN Ref_BIGOID_JenisKreditPembiayaan c ON c.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur d ON d.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' 
AND a.GolonganDebitur='600' AND c.StatusUMKM='Y' AND d.StatusUMKM='Y' "; 


//echo $query;
       $nama_debitur=array();
       $res=odbc_exec($connection2, $query);
while ($row=odbc_fetch_array($res)) {
      array_push($nama_debitur,$row['NamaLengkapDebitur']);
}
//$data=json_encode($nama_debitur);

//var_dump($nama_debitur);

/*
$q_modal_kerja=" SELECT SUM(a.JumlahKreditPeriodeLaporan) as modal_kerja1
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' ";


//echo $q_modal_kerja;
//die();
       //$modal_kerja1=array();
       $res=odbc_exec($connection2, $q_modal_kerja);
       $row=odbc_fetch_array($res); 
       $modal_kerjac12=$row['modal_kerja1'];

*/


#--Laporan B--
#--Pihak Penyalur--
$query2 = "SELECT distinct b.NamaLengkapDebitur
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN DM_InfoDebitur b ON b.NomorNasabah = a.NomorNasabah 
JOIN Ref_BIGOID_JenisKreditPembiayaan c ON c.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur d ON d.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' 
AND a.GolonganDebitur='600' AND c.StatusUMKM='Y' AND d.StatusUMKM='Y' ";

  $nama_debitur2=array();
       $res2=odbc_exec($connection2, $query2);
while ($row2=odbc_fetch_array($res2)) {
      array_push($nama_debitur2,$row2['NamaLengkapDebitur']);
}

/*
# LANCAR  KURANG LANCAR   DIRAGUKAN   MACET
#--Lancar--
$q_lancar=" SELECT SUM(a.JumlahKreditPeriodeLaporan) as lancar
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas IN ('1','2')AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' ";
       $res=odbc_exec($connection2, $q_lancar);
       $row=odbc_fetch_array($res); 
       $lancar_s2_c7=$row['lancar'];

#--Kurang Lancar--
$q_krg_lancar=" SELECT SUM(JumlahKreditPeriodeLaporan) as kurang_lancar
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='3' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' ";
       $res=odbc_exec($connection2, $q_krg_lancar);
       $row=odbc_fetch_array($res); 
       $krg_lancar_s2_d7=$row['kurang_lancar'];
//echo $q_krg_lancar;
//die();
#--Diragukan--
$q_diragukan="SELECT SUM(JumlahKreditPeriodeLaporan) as diragukan
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='4' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' ";
       $res=odbc_exec($connection2, $q_diragukan);
       $row=odbc_fetch_array($res); 
       $diragukan_s2_e7=$row['diragukan'];
#--Macet--
$q_macet=" SELECT SUM(JumlahKreditPeriodeLaporan) as macet
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='5' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' ";
       $res=odbc_exec($connection2, $q_macet);
       $row=odbc_fetch_array($res); 
       $macet_s2_f7=$row['macet'];

*/
//die();
        //$e28=$row['jml_npl'];
#--Modal Kerja--
        /*

*/

/*
--Laporan B--
--Pihak Penyalur--
SELECT distinct b.NamaLengkapDebitur
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN DM_InfoDebitur b ON b.NomorNasabah = a.NomorNasabah 
JOIN Ref_BIGOID_JenisKreditPembiayaan c ON c.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur d ON d.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='2016-05-31' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' 
AND a.GolonganDebitur='600' AND c.StatusUMKM='Y' AND d.StatusUMKM='Y'

--Kualitas--
--Lancar--
SELECT SUM(a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='2016-05-31' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas IN ('1','2')AND b.StatusUMKM='Y' AND c.StatusUMKM='Y'
--Kurang Lancar--
SELECT SUM(JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='2016-05-31' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='3' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y'
--Diragukan--
SELECT SUM(JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='2016-05-31' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='4' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y'
--Macet--
SELECT SUM(JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
WHERE a.DataDate='2016-05-31' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='5' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y'



*/



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
$objPHPExcel->getActiveSheet()->getStyle('A4')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:F7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B9:F11')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A4:E6')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayAlignment1);
//$objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayAlignment2);
$objPHPExcel->getActiveSheet()->getStyle('E9')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C9')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('D11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('E11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B21')->applyFromArray($styleArrayAlignment1);

//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B9:E21')->applyFromArray($styleArrayBorder1);




//FILL COLOR
//$objPHPExcel->getActiveSheet()->getStyle('A10:E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');


$objPHPExcel->getActiveSheet()->getStyle('A1:Z8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A22:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:A1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(50);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A4:E4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A5:E5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A6:E6');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B9:B11');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C9:D10');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E9:E10');

$objPHPExcel->getActiveSheet()->getStyle('C13:E21')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('A4', 'LAPORAN REALISASI PEMBERIAN KREDIT ATAU PEMBIAYAAN UMKM MELALUI KERJASAMA POLA EXECUTING');
$objPHPExcel->getActiveSheet()->setCellValue('A5', "POSISI TRIWULAN  II  TAHUN $tahun ");
$objPHPExcel->getActiveSheet()->setCellValue('A7', 'A.');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'Laporan Baki Debet Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'PIHAK PENYALUR');

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jenis Penggunaan ( Dalam Rupiah )');
$objPHPExcel->getActiveSheet()->setCellValue('C11', 'Modal Kerja');
$objPHPExcel->getActiveSheet()->setCellValue('D11', 'Investasi');

$objPHPExcel->getActiveSheet()->setCellValue('E9', 'TOTAL KREDIT / PEMBIAYAAN UMKM');
$objPHPExcel->getActiveSheet()->setCellValue('E11', '(Dalam Rupiah)');

/*
echo $nama_debitur['0']."<br>";
echo $nama_debitur['1']."<br>";
echo $nama_debitur['2']."<br>";
echo $nama_debitur['3']."<br>";
echo $nama_debitur['4']."<br>";
*/

$objPHPExcel->getActiveSheet()->setCellValue('B12', "BPR");
/*
$objPHPExcel->getActiveSheet()->setCellValue('B13', $nama_debitur['0']);
$objPHPExcel->getActiveSheet()->setCellValue('B14', $nama_debitur['1']);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $nama_debitur['2']);
$objPHPExcel->getActiveSheet()->setCellValue('B16', $nama_debitur['3']);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $nama_debitur['4']);
$objPHPExcel->getActiveSheet()->setCellValue('B18', $nama_debitur['5']);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $nama_debitur['6']);
$objPHPExcel->getActiveSheet()->setCellValue('B20', $nama_debitur['7']);

$objPHPExcel->getActiveSheet()->setCellValue('C12', "");
$objPHPExcel->getActiveSheet()->setCellValue('C13', $modal_kerjac12);
$objPHPExcel->getActiveSheet()->setCellValue('C14', $modal_kerjac12);
$objPHPExcel->getActiveSheet()->setCellValue('C15', $modal_kerjac12);
$objPHPExcel->getActiveSheet()->setCellValue('C16', $modal_kerjac12);
$objPHPExcel->getActiveSheet()->setCellValue('C17', $modal_kerjac12);
$objPHPExcel->getActiveSheet()->setCellValue('C18', "");
$objPHPExcel->getActiveSheet()->setCellValue('C19', "");
$objPHPExcel->getActiveSheet()->setCellValue('C20', "");
*/




//$colors = array("red", "green", "blue", "yellow"); 
$number_row1=13;
foreach ($nama_debitur as $val) {
 //   echo "$value <br>";
$objPHPExcel->getActiveSheet()->setCellValue("B$number_row1", $val);
$q_modal_kerja=" SELECT SUM(a.JumlahKreditPeriodeLaporan) as modal_kerja
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
JOIN DM_InfoDebitur d ON d.NomorNasabah = a.NomorNasabah 
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' and d.NamaLengkapDebitur='$val'";

       $res=odbc_exec($connection2, $q_modal_kerja);
       $row=odbc_fetch_array($res); 
       $modal_kerja=$row['modal_kerja'];
$objPHPExcel->getActiveSheet()->setCellValue("C$number_row1", $modal_kerja);


$number_row1++;
}


$objPHPExcel->getActiveSheet()->setCellValue('E12', "");
$objPHPExcel->getActiveSheet()->setCellValue('E13', "=(C13+D13)");
$objPHPExcel->getActiveSheet()->setCellValue('E14', "=(C14+D14)");
$objPHPExcel->getActiveSheet()->setCellValue('E15', "=(C15+D15)");
$objPHPExcel->getActiveSheet()->setCellValue('E16', "=(C16+D16)");
$objPHPExcel->getActiveSheet()->setCellValue('E17', "=(C17+D17)");
$objPHPExcel->getActiveSheet()->setCellValue('E18', "=(C18+D18)");
$objPHPExcel->getActiveSheet()->setCellValue('E19', "=(C19+D19)");
$objPHPExcel->getActiveSheet()->setCellValue('E20', "=(C20+D20)");

$objPHPExcel->getActiveSheet()->setCellValue('B21', 'TOTAL');

# TOTAL 
$objPHPExcel->getActiveSheet()->setCellValue('C21', "=SUM(C13:C20)");
$objPHPExcel->getActiveSheet()->setCellValue('D21', "=SUM(D13:D20)");
$objPHPExcel->getActiveSheet()->setCellValue('E21', "=SUM(E13:E20)");

// TITLE 
$objPHPExcel->getActiveSheet()->setTitle('LAPORAN A');

for ($i=13;$i<22;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=13;$i<22;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}

for ($i=13;$i<22;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}







// SHEET KE 2 ======================================================================================
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
$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('C7:G15')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(30);

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B4:B5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C4:F4');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('G4:G5');

$objPHPExcel->getActiveSheet()->getStyle('A2:F5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayFontBold);




$objPHPExcel->getActiveSheet()->getStyle('A4:G5')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('B4:G15')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A2', 'B.');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Laporan Kualitas Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing.');


$objPHPExcel->getActiveSheet()->setCellValue('B4', 'PIHAK PENYALUR');
$objPHPExcel->getActiveSheet()->setCellValue('C4', 'KUALITAS ( Dalam Rupiah )');
$objPHPExcel->getActiveSheet()->setCellValue('G4', 'TOTAL');

$objPHPExcel->getActiveSheet()->setCellValue('C5', 'LANCAR ');
$objPHPExcel->getActiveSheet()->setCellValue('D5', 'KURANG LANCAR');
$objPHPExcel->getActiveSheet()->setCellValue('E5', 'DIRAGUKAN');
$objPHPExcel->getActiveSheet()->setCellValue('F5', 'MACET');





$objPHPExcel->getActiveSheet()->setCellValue('B6', "BPR");

$number_row2=7;
foreach ($nama_debitur2 as $val2) {

$objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", $val2);
$q_lancar=" SELECT SUM(a.JumlahKreditPeriodeLaporan) as lancar
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
JOIN DM_InfoDebitur d ON d.NomorNasabah = a.NomorNasabah 
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas IN ('1','2')AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' and d.NamaLengkapDebitur='$val2' ";

       $res=odbc_exec($connection2, $q_lancar);
       $row=odbc_fetch_array($res); 
       $lancar=$row['lancar']; 
#PRINT LANCAR
$objPHPExcel->getActiveSheet()->setCellValue("C$number_row2", $lancar);

$q_krg_lancar=" SELECT SUM(JumlahKreditPeriodeLaporan) as kurang_lancar
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
JOIN DM_InfoDebitur d ON d.NomorNasabah = a.NomorNasabah 
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas IN ('3')AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' and d.NamaLengkapDebitur='$val2' ";
       $res=odbc_exec($connection2, $q_krg_lancar);
       $row=odbc_fetch_array($res); 
       $krg_lancar=$row['kurang_lancar'];

#PRINT KURANG LANCAR
$objPHPExcel->getActiveSheet()->setCellValue("D$number_row2", $krg_lancar);

$q_diragukan="SELECT SUM(JumlahKreditPeriodeLaporan) as diragukan
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
JOIN DM_InfoDebitur d ON d.NomorNasabah = a.NomorNasabah 
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='4' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y' and d.NamaLengkapDebitur='$val2' ";

       $res=odbc_exec($connection2, $q_diragukan);
       $row=odbc_fetch_array($res); 
       $diragukan=$row['diragukan'];

#PRINT DIRAGUKAN
$objPHPExcel->getActiveSheet()->setCellValue("E$number_row2", $diragukan);

$q_macet=" SELECT SUM(JumlahKreditPeriodeLaporan) as macet
FROM DM_AsetKredit a WITH (NOLOCK)
JOIN Ref_BIGOID_JenisKreditPembiayaan b ON b.KodeInternal = a.Jenis
JOIN Ref_BIGOID_GolonganDebitur c ON c.KodeInternal = a.GolonganDebitur
JOIN DM_InfoDebitur d ON d.NomorNasabah = a.NomorNasabah 
WHERE a.DataDate='$var_tgl' AND a.Jenis ='25' AND a.KategoriUsahaDebitur <>'99' AND a.GolonganDebitur='600'
AND a.Kolektibilitas='5' AND b.StatusUMKM='Y' AND c.StatusUMKM='Y'  and d.NamaLengkapDebitur='$val2' ";
       $res=odbc_exec($connection2, $q_macet);
       $row=odbc_fetch_array($res); 
       $macet=$row['macet'];

#PRINT MACET
$objPHPExcel->getActiveSheet()->setCellValue("F$number_row2", $macet);

$number_row2++;
}

/*
$objPHPExcel->getActiveSheet()->setCellValue('B7', $nama_debitur2['0']);
$objPHPExcel->getActiveSheet()->setCellValue('B8', $nama_debitur2['1']);
$objPHPExcel->getActiveSheet()->setCellValue('B9', $nama_debitur2['2']);
$objPHPExcel->getActiveSheet()->setCellValue('B10', $nama_debitur2['3']);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $nama_debitur2['4']);
$objPHPExcel->getActiveSheet()->setCellValue('B12', $nama_debitur2['5']);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $nama_debitur2['6']);
$objPHPExcel->getActiveSheet()->setCellValue('B14', $nama_debitur2['7']);

$objPHPExcel->getActiveSheet()->setCellValue('C7', $lancar_s2_c7);
$objPHPExcel->getActiveSheet()->setCellValue('C8', $lancar_s2_c7);
$objPHPExcel->getActiveSheet()->setCellValue('C9', $lancar_s2_c7);
$objPHPExcel->getActiveSheet()->setCellValue('C10', $lancar_s2_c7);
$objPHPExcel->getActiveSheet()->setCellValue('C11', $lancar_s2_c7);
//$objPHPExcel->getActiveSheet()->setCellValue('C12', $lancar_s2_c7);
//$objPHPExcel->getActiveSheet()->setCellValue('C13', $lancar_s2_c7);
//$objPHPExcel->getActiveSheet()->setCellValue('C14', $lancar_s2_c7);


$objPHPExcel->getActiveSheet()->setCellValue('D7', $krg_lancar_s2_d7);
$objPHPExcel->getActiveSheet()->setCellValue('D8', $krg_lancar_s2_d7);
$objPHPExcel->getActiveSheet()->setCellValue('D9', $krg_lancar_s2_d7);
$objPHPExcel->getActiveSheet()->setCellValue('D10', $krg_lancar_s2_d7);
$objPHPExcel->getActiveSheet()->setCellValue('D11', $krg_lancar_s2_d7);
//$objPHPExcel->getActiveSheet()->setCellValue('D12', $krg_lancar_s2_d7);
//$objPHPExcel->getActiveSheet()->setCellValue('D13', $krg_lancar_s2_d7);
//$objPHPExcel->getActiveSheet()->setCellValue('D14', $krg_lancar_s2_d7);


$objPHPExcel->getActiveSheet()->setCellValue('E7', $diragukan_s2_e7);
$objPHPExcel->getActiveSheet()->setCellValue('E8', $diragukan_s2_e7);
$objPHPExcel->getActiveSheet()->setCellValue('E9', $diragukan_s2_e7);
$objPHPExcel->getActiveSheet()->setCellValue('E10', $diragukan_s2_e7);
$objPHPExcel->getActiveSheet()->setCellValue('E11', $diragukan_s2_e7);
//$objPHPExcel->getActiveSheet()->setCellValue('E12', $diragukan_s2_e7);
//$objPHPExcel->getActiveSheet()->setCellValue('E13', $diragukan_s2_e7);
//$objPHPExcel->getActiveSheet()->setCellValue('E14', $diragukan_s2_e7);


$objPHPExcel->getActiveSheet()->setCellValue('F7', $macet_s2_f7);
$objPHPExcel->getActiveSheet()->setCellValue('F8', $macet_s2_f7);
$objPHPExcel->getActiveSheet()->setCellValue('F9', $macet_s2_f7);
$objPHPExcel->getActiveSheet()->setCellValue('F10', $macet_s2_f7);
$objPHPExcel->getActiveSheet()->setCellValue('F11', $macet_s2_f7);
//$objPHPExcel->getActiveSheet()->setCellValue('F12', $macet_s2_f7);
//$objPHPExcel->getActiveSheet()->setCellValue('F13', $macet_s2_f7);
//$objPHPExcel->getActiveSheet()->setCellValue('F14', $macet_s2_f7);


*/



$objPHPExcel->getActiveSheet()->setCellValue('G7', "=C7+D7+E7+F7");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "=C8+D8+E8+F8");
$objPHPExcel->getActiveSheet()->setCellValue('G9', "=C9+D9+E9+F9");
$objPHPExcel->getActiveSheet()->setCellValue('G10', "=C10+D10+E10+F10");
$objPHPExcel->getActiveSheet()->setCellValue('G11', "=C11+D11+E11+F11");
$objPHPExcel->getActiveSheet()->setCellValue('G12', "=C12+D12+E12+F12");
$objPHPExcel->getActiveSheet()->setCellValue('G13', "=C13+D13+E13+F13");
$objPHPExcel->getActiveSheet()->setCellValue('G14', "=C14+D14+E14+F14");


$objPHPExcel->getActiveSheet()->setCellValue('C15', "=SUM(C7:C14)");
$objPHPExcel->getActiveSheet()->setCellValue('D15', "=SUM(D7:D14)");
$objPHPExcel->getActiveSheet()->setCellValue('E15', "=SUM(E7:E14)");
$objPHPExcel->getActiveSheet()->setCellValue('F15', "=SUM(F7:F14)");
$objPHPExcel->getActiveSheet()->setCellValue('G15', "=SUM(G7:G14)");

$objPHPExcel->getActiveSheet()->setCellValue('B15', 'TOTAL');

for ($i=7;$i<16;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=7;$i<16;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}

for ($i=7;$i<16;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}
for ($i=7;$i<16;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=7;$i<16;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}



$objPHPExcel->getActiveSheet()->setTitle('LAPORAN B');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_UMKM_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Report_UMKM_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);

//echo $objPHPExcel->getActiveSheet(1)->getCell('B13')->getValue(); 


//die();
?>
<!--
<b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_UMKM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br></div> </b>
<br>
<br>
<br>
-->


<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> LAPORAN UMKM 
                            </div>
                            <div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>

                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div>
                        </div>
                        <div class="portlet-body">
                            <h4 ><b>PT. Bank MNC Internasional .Tbk</b></h4>
                            <div class="tabbable-line">
                                <ul class="nav nav-tabs ">
                                    <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        LAPORAN A </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        LAPORAN B </a>
                                    </li>
                                  
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_UMKM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>
    <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> A. Laporan Baki Debet Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing</b>
                                    </div>                                  
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               <tr class="active">
                                                <td width="30%" align="left"><b></b></td>
                                                <td width="50%" align="center" colspan="2"><b>Jenis Penggunaan (Dalam Rupiah) </b></td>
                                                <td width="20%" align="center"><b>Total Kredit / Pembiayaan UMKM </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                <td width="30%" align="center"><b>PIHAK PENYALUR </b></td>
                                                <td width="25%" align="center" ><b>Modal Kerja </b></td>
                                                <td width="25%" align="center" ><b>Investasi </b></td>
                                                <td width="20%" align="center"><b> (Dalam Rupiah) </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  style="font-size:12px" align="left" >BPR </td>
                                                <td  style="font-size:12px" ></td>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px"></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B13')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B14')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B15')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B16')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B17')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B18')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B19')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B20')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell('B21')->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>
                                  
                                     <div class="tab-pane" id="tab_15_2">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> B. Laporan Kualitas Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing.</b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(1);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="25%" align="left"><b></b></td>
                                                <td width="60%" align="center" colspan="4"><b>KUALITAS (Dalam Rupiah) </b></td>
                                                <td width="15%" align="center"><b>Total  </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                <td width="25%" align="center"><b>PIHAK PENYALUR </b></td>
                                                <td width="15%" align="center" ><b>Lancar </b></td>
                                                <td width="15%" align="center" ><b>Kurang Lancar </b></td>
                                                <td width="15%" align="center" ><b>Diragukan </b></td>
                                                <td width="15%" align="center" ><b>Macet </b></td>
                                                <td width="15%" align="center"><b></b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="left">BPR </td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center"></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B7')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B8')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B9')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B10')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B11')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B12')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B13')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B14')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell('B15')->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('C15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('F15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('G15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>

                                </div>
                            </div>
                            
                        </div>
                </div>

