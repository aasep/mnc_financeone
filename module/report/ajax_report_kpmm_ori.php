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
logActivity("generate KPMM",date('Y_m_d_H_i_s'));

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




$var_tabel=date('Ymd',strtotime($tanggal));

/*

$tanggal=$_POST['tanggal']; 

$curr_tgl=date('Y-m',strtotime($tanggal));
$end_curr_tgl=date('Y-m',strtotime($tanggal));

$label_tgl=date('M-y',strtotime($tanggal));


$label_txtfile=date('Ym',strtotime($tanggal));
$tanggal_header=date('mY',strtotime($tanggal));


//$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

//$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih




$var_tabel=date('Ym',strtotime($tanggal));

#############################################################################################

$var_bulan=date('m',strtotime($tanggal));
$var_tahun=date('Y',strtotime($tanggal));
*/
/*
$query= " select a.Nominal from DM_Journal a 
left join Referensi_GL_02 b on a.KodeGL=b.GLNO
left join Referensi_KPMM c on b.KPMM_Level_3=c.KPMM_Level_3
left join Master_ATMR d on a.DataDate=d.DataDate 
WHERE a.DataDate='$curr_tgl'  ";
*/
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM (
SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_KPMM c ON c.KPMM_Level_3 = b.KPMM_Level_3
WHERE a.DataDate='$curr_tgl'  ";

$q_add2="  GROUP BY a.kodegl ,b.KPMM_Level_3 )AS tabel1 ";

// where Month(DataDate)='$curr_mon' and Year(DataDate)='$curr_year' 
 $curr_mon=date('n',strtotime($tanggal));
 $curr_year=date('Y',strtotime($tanggal));


$query_atmr= " select * from Master_ATMR WHERE where Month(DataDate)='$curr_mon' and Year(DataDate)='$curr_year'  ";
//echo $query_atmr;
//die();
$result_atmr=odbc_exec($connection2, $query_atmr);
$rowAtmr=odbc_fetch_array($result_atmr);
$d54=$rowAtmr['ATMR_Kredit'];
$d55=$rowAtmr['ATMR_Pasar'];
$d56=$rowAtmr['ATMR_Operasional'];

//KPMM101000001   PENEMPATAN PADA BANK
$q_add=" and b.KPMM_Level_3='KPMM101000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);

//KPMM205000001   Modal disetor (Setelah dikurangi Saham Treasury)
$q_add=" and b.KPMM_Level_3='KPMM205000001' ";
//echo $query.$q_add.$q_add2;
//die();
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$i11=$row['Jumlah_Nominal'];


//KPMM101000002   SURAT BERHARGA
$q_add=" and b.KPMM_Level_3='KPMM101000002' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM101000003   OTHER (AKSEPTASI/DERIVATIF)
$q_add=" and b.KPMM_Level_3='KPMM101000003' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM101000004   KREDIT
$q_add=" and b.KPMM_Level_3='KPMM101000004' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM102000001   AYDA
$q_add=" and b.KPMM_Level_3='KPMM102000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM102000002   TAGIHAN LAINNYA
$q_add=" and b.KPMM_Level_3='KPMM102000002' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM102000003   ASET TIDAK BERWUJUD LAINNYA
$q_add=" and b.KPMM_Level_3='KPMM102000003' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c31=$row['Jumlah_Nominal'];
//KPMM201000001   CADANGAN UMUM
$q_add=" and b.KPMM_Level_3='KPMM201000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c15=$row['Jumlah_Nominal'];
//KPMM202000001   AGIO / DISAGIO
$q_add=" and b.KPMM_Level_3='KPMM202000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c13=$row['Jumlah_Nominal'];
//KPMM203000001   PENDAPATAN KOMPREHENSIF
$q_add=" and b.KPMM_Level_3='KPMM203000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c22=$row['Jumlah_Nominal'];
//KPMM204000001   DANA SETORAN MODAL
$q_add=" and b.KPMM_Level_3='KPMM204000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c19=$row['Jumlah_Nominal'];
//KPMM205000001   MODAL DISETOR
$q_add=" and b.KPMM_Level_3='KPMM205000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM206000001   RUGI TAHUN TAHUN LALU YANG DAPAT DIPERHITUNGKAN
$q_add=" and b.KPMM_Level_3='KPMM206000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
/*$row=odbc_fetch_array($result);
$c16=$row['Jumlah_Nominal'];
*/
//KPMM301000001   LABA RUGI YANG DAPAT DIPERHITUNGKAN
$q_add=" and b.KPMM_Level_3='KPMM301000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM400000001   LCU
$q_add=" and b.KPMM_Level_3='KPMM400000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);



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
$objPHPExcel->getActiveSheet()->getStyle('A1:L4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('M1:O65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A63:O65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');


$objPHPExcel->getActiveSheet()->getStyle('A1:L7')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A9:L12')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A27:L28')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A37:L37')->applyFromArray($styleArrayFontBold);
//$objPHPExcel->getActiveSheet()->getStyle('A37:L49')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A41:L46')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A50:L52')->applyFromArray($styleArrayFontBold);



//$objPHPExcel->getActiveSheet()->getStyle('B14:H15')->applyFromArray($styleArrayFontBold);
//$objPHPExcel->getActiveSheet()->getStyle('B26:H27')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A1:L7')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('A50:L52')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A50:L52')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


//$objPHPExcel->getActiveSheet()->getStyle('A1:G7')->applyFromArray($styleArrayAlignment1);
//$objPHPExcel->getActiveSheet()->getStyle('B27:H28')->applyFromArray($styleArrayAlignment1);
//$objPHPExcel->getActiveSheet()->getStyle('B15:H16')->applyFromArray($styleArrayAlignment1);
//$objPHPExcel->getActiveSheet()->getStyle('C11')->applyFromArray($styleArrayAlignment1);
//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A5:L49')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A50:L62')->applyFromArray($styleArrayBorder1);


//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(90);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(65);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(15);
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:L1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:L2'); 
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:L3');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A5:H7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I5:J6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K5:L5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K6:L6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B9:C9');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B8:H8");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B9:H9");
for ($i=10; $i <=40  ; $i++) { 
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("C$i:H$i");
}
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B41:H41");
for ($i=42; $i <=48  ; $i++) { 
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells("C$i:H$i");
}


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A49:C49');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A50:C52');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H50:H52');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D50:E51');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F50:G50');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D51:E51');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F51:G51');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I50:J51');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K50:L50');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I51:J51');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K51:L51');


//PT
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'LAPORAN PERHITUNGAN KEWAJIBAN PENYEDIAAN MODAL MINIMUM TRIWULANAN BANK UMUM');
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT Bank MNC Internasional, Tbk');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Tanggal : $label_tgl ");

$objPHPExcel->getActiveSheet()->setCellValue('I5', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('K5', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('K6', 'Tahun Sebelumnya');

$objPHPExcel->getActiveSheet()->setCellValue('I7', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('K7', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('L7', 'Konsolidasi');
/*
D-->I
E-->J
F-->K
G-->L
*/

$objPHPExcel->getActiveSheet()->setCellValue('L4', '(dalam jutaan rupiah)');
$objPHPExcel->getActiveSheet()->setCellValue('A9', 'I');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Modal Inti (Tier 1)');
$objPHPExcel->getActiveSheet()->setCellValue('B10', '1');
$objPHPExcel->getActiveSheet()->setCellValue('C10', "Modal Inti Utama (CET 1)");
$objPHPExcel->getActiveSheet()->setCellValue('C11', '1.1     Modal disetor (Setelah dikurangi Saham Treasury)');
$objPHPExcel->getActiveSheet()->setCellValue('C12', '1.2     Cadangan Tambahan Modal ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('C13', '1.2.1   Agio / Disagio');
$objPHPExcel->getActiveSheet()->setCellValue('C14', '1.2.2   Modal sumbangan ');
$objPHPExcel->getActiveSheet()->setCellValue('C15', '1.2.3   Cadangan umum ');
$objPHPExcel->getActiveSheet()->setCellValue('C16', '1.2.4   Laba/Rugi tahun-tahun lalu yang dapat diperhitungkan');
$objPHPExcel->getActiveSheet()->setCellValue('C17', '1.2.5   Laba/Rugi tahun berjalan yang dapat diperhitungkan');
$objPHPExcel->getActiveSheet()->setCellValue('C18', '1.2.6   Selisih lebih karena penjabaran laporan keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('C19', '1.2.7   Dana setoran modal');
$objPHPExcel->getActiveSheet()->setCellValue('C20', '1.2.8   Waran yang diterbitkan ');
$objPHPExcel->getActiveSheet()->setCellValue('C21', '1.2.9   Opsi saham yang diterbitkan dalam rangka program kompensasi berbasis saham');
$objPHPExcel->getActiveSheet()->setCellValue('C22', '1.2.10  Pendapatan komprehensif lain ');
$objPHPExcel->getActiveSheet()->setCellValue('C23', '1.2.11  Saldo surplus revaluasi aset tetap ');
$objPHPExcel->getActiveSheet()->setCellValue('C24', '1.2.12  Selisih kurang antara PPA dan cadangan kerugian penurunan nilai atas aset produktif');
$objPHPExcel->getActiveSheet()->setCellValue('C25', '1.2.13  Penyisihan Penghapusan Aset (PPA) atas aset non produktif yang wajib dihitung');
$objPHPExcel->getActiveSheet()->setCellValue('C26', '1.2.14  Selisih kurang jumlah penyesuaian nilai wajar dari instrumen keuangan dalam <i>trading book</i>');
$objPHPExcel->getActiveSheet()->setCellValue('C27', '1.3     Kepentingan Non Pengendali yang dapat diperhitungkan');
$objPHPExcel->getActiveSheet()->setCellValue('C28', '1.4     Faktor Pengurang Modal Inti Utama ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('C29', '1.4.1   Perhitungan pajak tangguhan');
$objPHPExcel->getActiveSheet()->setCellValue('C30', '1.4.2   Goodwill');

$objPHPExcel->getActiveSheet()->setCellValue('C31', '1.4.3   Aset tidak berwujud lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('C32', '1.4.4   Penyertaan yang diperhitungkan sebagai faktor pengurang');
$objPHPExcel->getActiveSheet()->setCellValue('C33', '1.4.5   Kekurangan modal pada perusahaan anak asuransi');
$objPHPExcel->getActiveSheet()->setCellValue('C34', '1.4.6   Eksposur sekuritisasi');
$objPHPExcel->getActiveSheet()->setCellValue('C35', '1.4.7   Faktor Pengurang modal inti lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('C36', '1.4.8   Investasi pada instrumen AT1 dan Tier 2 pada bank lain ₂₎');

$objPHPExcel->getActiveSheet()->setCellValue('B37', '2');
$objPHPExcel->getActiveSheet()->setCellValue('C37', 'Modal Inti Tambahan (AT-1)  ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('C38', '2.1     Instrumen yang memenuhi persyaratan AT-1  ');
$objPHPExcel->getActiveSheet()->setCellValue('C39', '2.2     Agio / Disagio');
$objPHPExcel->getActiveSheet()->setCellValue('C40', '2.3     Faktor Pengurang: Investasi pada instrumen AT1 dan Tier 2 pada bank lain ₂₎ ');


$objPHPExcel->getActiveSheet()->setCellValue('A41', 'II');
$objPHPExcel->getActiveSheet()->setCellValue('B41', 'Modal Pelengkap (Tier 2)');
$objPHPExcel->getActiveSheet()->setCellValue('B42', '1');
$objPHPExcel->getActiveSheet()->setCellValue('C42', 'Instrumen modal dalam bentuk saham atau lainnya yang memenuhi persyaratan');
$objPHPExcel->getActiveSheet()->setCellValue('B43', '2');
$objPHPExcel->getActiveSheet()->setCellValue('C43', 'Agio atau disagio yang berasal dari penerbitan instrumen modal inti tambahan');
$objPHPExcel->getActiveSheet()->setCellValue('B44', '3');
$objPHPExcel->getActiveSheet()->setCellValue('C44', 'Cadangan umum aset produktif PPA yang wajib dibentuk (maks 1,25% ATMR Risiko Kredit)');
$objPHPExcel->getActiveSheet()->setCellValue('B45', '4');
$objPHPExcel->getActiveSheet()->setCellValue('C45', 'Cadangan tujuan');
$objPHPExcel->getActiveSheet()->setCellValue('B46', '5');
$objPHPExcel->getActiveSheet()->setCellValue('C46', 'Faktor Pengurang Modal Pelengkap ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('C47', '5.1     Sinking Fund ');
$objPHPExcel->getActiveSheet()->setCellValue('C48', '5.2     Investasi pada instrumen Tier 2 pada bank lain ₂₎');

$objPHPExcel->getActiveSheet()->setCellValue('A49', ' TOTAL MODAL ');



$objPHPExcel->getActiveSheet()->setCellValue('D50', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('F50', 'Posisi Tgl Laporan Thn Lalu');
//$objPHPExcel->getActiveSheet()->setCellValue('F51', 'Tahun Lalu');

$objPHPExcel->getActiveSheet()->setCellValue('D52', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('E52', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('F52', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('G52', 'Konsolidasi');

$objPHPExcel->getActiveSheet()->setCellValue('I50', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('K50', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('K51', 'Tahun Lalu');

$objPHPExcel->getActiveSheet()->setCellValue('I52', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('J52', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('K52', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('L52', 'Konsolidasi');

$objPHPExcel->getActiveSheet()->setCellValue('H50', 'KETERANGAN');


$objPHPExcel->getActiveSheet()->setCellValue('C53', 'ASET TERTIMBANG MENURUT RISIKO');
$objPHPExcel->getActiveSheet()->setCellValue('C54', 'ATMR RISIKO KREDIT ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('C55', 'ATMR RISIKO PASAR');
$objPHPExcel->getActiveSheet()->setCellValue('C56', 'ATMR RISIKO OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('C57', 'TOTAL ATMR');
$objPHPExcel->getActiveSheet()->setCellValue('C58', 'RASIO KPMM SESUAI PROFIL RISIKO');

#fill ATMR dan jumlah dari ATMR
$objPHPExcel->getActiveSheet()->setCellValue("D54", $d54);
$objPHPExcel->getActiveSheet()->setCellValue("D55", $d55);
$objPHPExcel->getActiveSheet()->setCellValue("D56", $d56);
$objPHPExcel->getActiveSheet()->setCellValue("D57", "=SUM(D54:D56)");
$objPHPExcel->getActiveSheet()->setCellValue("D58", "=ABS(+IF(I49=0,0,(I49/D57)))");


$objPHPExcel->getActiveSheet()->setCellValue('C59', 'ALOKASI PEMENUHAN KPMM ');
$objPHPExcel->getActiveSheet()->setCellValue('C60', 'Dari CET1');
$objPHPExcel->getActiveSheet()->setCellValue('C61', 'Dari AT1');
$objPHPExcel->getActiveSheet()->setCellValue('C62', 'Dari Tier 2');


$objPHPExcel->getActiveSheet()->setCellValue('H53', 'RASIO KPMM');
$objPHPExcel->getActiveSheet()->setCellValue('H54', 'Rasio CET1');
$objPHPExcel->getActiveSheet()->setCellValue('H55', 'Rasio Tier 1');
$objPHPExcel->getActiveSheet()->setCellValue('H56', 'Rasio Tier 2');
$objPHPExcel->getActiveSheet()->setCellValue('H57', 'Rasio total');
$objPHPExcel->getActiveSheet()->setCellValue('H58', 'CET 1 UNTUK  BUFFER ');
$objPHPExcel->getActiveSheet()->setCellValue('H59', 'PERSENTASE BUFFER YANG WAJIB DIPENUHI OLEH BANK   ');
$objPHPExcel->getActiveSheet()->setCellValue('H60', 'Capital Conservation Buffer');
$objPHPExcel->getActiveSheet()->setCellValue('H61', 'Countercyclical Buffer');
$objPHPExcel->getActiveSheet()->setCellValue('H62', 'Capital Surcharge untuk D-SIB');
//===================================
//Kalkulasi
$objPHPExcel->getActiveSheet()->setCellValue('I9', "=I10+I37");
$objPHPExcel->getActiveSheet()->setCellValue('I10', "=I11+I12+I27-I28");
$objPHPExcel->getActiveSheet()->setCellValue('I12', "=SUM(I13:I26)");
$objPHPExcel->getActiveSheet()->setCellValue('I28', "=SUM(I29:I36)");
//$objPHPExcel->getActiveSheet()->setCellValue('I28', "=SUM(N29:N36)");
$objPHPExcel->getActiveSheet()->setCellValue('I41', "=I42+I43+I44+I45+I46");
$objPHPExcel->getActiveSheet()->setCellValue('I49', "=I41+I9");

//===================================
$objPHPExcel->getActiveSheet()->setCellValue('I11', abs($i11));
$objPHPExcel->getActiveSheet()->setCellValue('I31', $c31);
$objPHPExcel->getActiveSheet()->setCellValue('I15', abs($c15));
$objPHPExcel->getActiveSheet()->setCellValue('I13', abs($c13));
$objPHPExcel->getActiveSheet()->setCellValue('I22', $c22);
$objPHPExcel->getActiveSheet()->setCellValue('I19', $c29);
$objPHPExcel->getActiveSheet()->setCellValue('I16', $c16);


$objPHPExcel->getActiveSheet()->getStyle('D58')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));      
$objPHPExcel->getActiveSheet()->getStyle('A9:L9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('483D8B');
$objPHPExcel->getActiveSheet()->getStyle('A41:L41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('483D8B');
$objPHPExcel->getActiveSheet()->getStyle('A5:L7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('A50:L52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');

//$objPHPExcel->getActiveSheet()->getStyle('C17:H22')->getNumberFormat()->setFormatCode('0.00');
//$objPHPExcel->getActiveSheet()->getStyle('C29:H34')->getNumberFormat()->setFormatCode('0.00');
    
$objPHPExcel->getActiveSheet()->getStyle('I9:L49')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('D54:D57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

for ($i=9;$i<=49;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=9;$i<=49;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=9;$i<=49;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}
for ($i=9;$i<=49;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('L'.$i, 0);
    }
}

$objPHPExcel->getActiveSheet()->setTitle('KPMM');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/KPMM_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/KPMM_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>

<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> KPMM 
                            </div>
                            <div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>

                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div>
                        </div>
                        <div class="portlet-body">
                            <h4 ><b>PT Bank MNC Internasional, Tbk</b></h4>
                            
                            
                            <div class="tabbable-line">
                             <!-- <ul class="nav nav-tabs ">
                                    <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        KPMM </a>
                                    </li>
                                </ul>
							-->
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/KPMM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div>
<br>
<br>
<br>
<br>
<div class="pull-right" >(dalam jutaan rupiah)</div>
</b> 

                                            

</br>

                                        <p>
                                       
                                         
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="60%" align="center" rowspan="2" colspan="3"><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b> Posisi Tgl Laporan </b></td>
                                                <td width="15%" align="center" colspan="2"><b> Posisi Tgl Laporan Thn Sebelumnya </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="10%" align="center"><b>Bank</b></td>
                                                <td width="10%" align="center"><b>Konsolidasi</b></td>
                                                <td width="10%" align="center"><b>Bank</b></td>
                                                <td width="10%" align="center"><b>Konsolidasi</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                               <?php 

                                               for ($i=9; $i <= 49 ; $i++) { 

                                                if ($i=='9' || $i=='41' || $i=='49') {

                                                    

                                                 ?>     
                                                

                                                <tr>
                                                <td align="left" > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td align="left" colspan="2"> <b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <?php
                                        } else {
                                                ?>

                                                <tr>
                                                <td align="left" > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>								
												<td align="left" > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                <?php
                                                if ($i=='10' ||$i=='11' ||$i=='12') {

                                                ?>
                                                <td align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <?php
                                                } else  {


                                                ?>

                                                <td align="left" ><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <?php
                                            }

                                                ?>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>



                                                 <?php
                                         }
                                               }
                                               ?>
                                                 
                                                </tbody>
                                            </table>
                                        </div>


                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="20%" align="center" rowspan="2" ><b></b></td>
                                                <td width="17%" align="center" colspan="2"><br><b> Posisi Tgl Laporan </b></br></td>
                                                <td width="20%" align="center" colspan="2"><br><b> Posisi Laporan Thn Sebelumnya </b></br></td>
                                                <td width="15%" align="center" rowspan="2"><br><br><b> Keterangan </br></br></b></td>
                                                <td width="17%" align="center" colspan="2"><br><b> Posisi Tgl Laporan </b></br></td>
                                                <td width="20%" align="center" colspan="2"><br><b> Posisi Laporan Thn Sebelumnya </b></br></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="7%" align="center"><b>Bank</b></td>
                                                <td width="7%" align="center"><b>Konsolidasi</b></td>
                                                <td width="10%" align="center"><b>Bank</b></td>
                                                <td width="10%" align="center"><b>Konsolidasi</b></td>
                                                <td width="7%" align="center"><b>Bank</b></td>
                                                <td width="7%" align="center"><b>Konsolidasi</b></td>
                                                <td width="8%" align="center"><b>Bank</b></td>
                                                <td width="12 %" align="center"><b>Konsolidasi</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>

                                               <?php 

                                               for ($i=53; $i <= 62 ; $i++) { 

                                                if ($i=='9' || $i=='41' || $i=='49') {
                                                 ?>     
                                                

                                                <tr>
                                                <td align="left" > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td align="left" colspan="2"> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <?php
                                        } else {
                                                ?>

                                                <tr>
                                                <td align="left" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="left" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="left" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="left" > <?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>



                                                 <?php
                                         }
                                               }
                                               ?>
                                                 
                                                </tbody>
                                            </table>
                                        </div>        







                                    </div>
                                  
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

