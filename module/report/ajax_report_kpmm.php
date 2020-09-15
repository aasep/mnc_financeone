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

##### QUERY KPMM ############## 
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM (
SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_KPMM c ON c.KPMM_Level_3 = b.KPMM_Level_3
WHERE a.DataDate='$curr_tgl'  ";

$q_add2="  GROUP BY a.kodegl ,b.KPMM_Level_3 )AS tabel1 ";

 $curr_mon=date('n',strtotime($tanggal));
 $curr_year=date('Y',strtotime($tanggal));
 
#### Query to Master Modal ####
$q_modal=" select Nominal_Modal as modal_master from Master_Modal where Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal' ";
$res_modal=odbc_exec($connection2, $q_modal);
$row_modal=odbc_fetch_array($res_modal);
$found_modal=odbc_num_rows($res_modal);
$modal_nilai_fix=$row_modal['modal_master'];


if ($found_modal ==0 || !isset($found_modal)){
    echo "<div class='alert alert-danger alert-dismissable'><b>Data on month $mon_modal - $year_modal is not available.</b> </div>";
    die();
    }

# +++++++++++ QUERY ATMR ++++++++++++++
$query_atmr= " select * from Master_ATMR WHERE  Month(DataDate)='$curr_mon' and Year(DataDate)='$curr_year'  ";
//echo $query_atmr;
//die();

$result_atmr=odbc_exec($connection2, $query_atmr);
$rowAtmr=odbc_fetch_array($result_atmr);
$atmr_kredit=$rowAtmr['ATMR_Kredit'];
$atmr_pasar=$rowAtmr['ATMR_Pasar'];
$atmr_operasional=$rowAtmr['ATMR_Operasional'];

// remark 2016-12-22
// KPMM101000001   PENEMPATAN PADA BANK
//$q_add=" and b.KPMM_Level_3='KPMM101000001' ";
//$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//$row=odbc_fetch_array($result);

//1.1 +++++  KPMM205000001   Modal disetor (Setelah dikurangi Saham Treasury)
$q_add=" and b.KPMM_Level_3='KPMM205000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m11=abs($row['Jumlah_Nominal']);

#--1.4.1 Perhitungan Pajak Tangguhan  KPMM102000004
//KPMM102000004

$q_add=" and b.KPMM_Level_3='KPMM102000004' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m39=$row['Jumlah_Nominal'];

//echo $query.$q_add.$q_add2;
//die();
/*
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
*/

// 1.2.1.2.2  ++++  KPMM201000001   CADANGAN UMUM
$q_add=" and b.KPMM_Level_3='KPMM201000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m20=abs($row['Jumlah_Nominal']);
// 1.2.1.2.1  +++++  KPMM202000001   AGIO / DISAGIO
$q_add=" and b.KPMM_Level_3='KPMM202000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m19=abs($row['Jumlah_Nominal']);


#--1.4.3 Seluruh aset tidak berwujud lainnya
$q_add=" and b.KPMM_Level_3='KPMM102000003' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m41=$row['Jumlah_Nominal'];
//echo $query.$q_add.$q_add2;
// 1.2.1.2.5  +++++   KPMM204000001   DANA SETORAN MODAL
$q_add=" and b.KPMM_Level_3='KPMM204000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m23=abs($row['Jumlah_Nominal']);


/*
//KPMM203000001   PENDAPATAN KOMPREHENSIF
$q_add=" and b.KPMM_Level_3='KPMM203000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$c22=$row['Jumlah_Nominal'];
*/

#--1.2.1.2.4 : Laba Tahun Berjalan (*Note : dibuat positive)
$query = " select Tanggal,sum(Jumlah)/1000000 as Jumlah_Nominal from (
select DataDate as Tanggal,sum (Nominal) as Jumlah from DM_Journal WHERE DataDate='$curr_tgl' 
AND LEFT(KODEGL,1) in ('4','5') and kodegl not in ('40501001','40501002','40504001','40504002','40504006')
group by datadate
UNION ALL
select DataDate as Tanggal,sum (Nominal) as Jumlah from DM_Journal WHERE DataDate='$curr_tgl' 
AND KodeGL in('40501001','40501002','40504001','40504002','40504006') AND Nominal>='0'
group by datadate
)as tabel1
group by tanggal ";

$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m22=abs($row['Jumlah_Nominal']);




#--1.2.2.1.2 : Potensi kerugian dari penurunan nilai wajar aset keuangan dalam kelompok tersedia untuk dijual
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM 
(
 SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
 JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
 WHERE a.DataDate='$curr_tgl' AND b.GLNO in ('30102003','30102005','30102006')
 GROUP BY a.kodegl ,b.GLNO ) AS tabel1 ";

$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m28=abs($row['Jumlah_Nominal']);



#--- 1.2.1.1.2-----

$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM 
(
 SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
 JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
 WHERE a.DataDate='$curr_tgl' AND b.GLNO='30102009'
 GROUP BY a.kodegl ,b.GLNO
) AS tabel1 ";
$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m16=abs($row['Jumlah_Nominal']);
//echo $query;
//die();



/*
//KPMM205000001   MODAL DISETOR
$q_add=" and b.KPMM_Level_3='KPMM205000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM206000001   RUGI TAHUN TAHUN LALU YANG DAPAT DIPERHITUNGKAN
$q_add=" and b.KPMM_Level_3='KPMM206000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//$row=odbc_fetch_array($result);
//$c16=$row['Jumlah_Nominal'];

//KPMM301000001   LABA RUGI YANG DAPAT DIPERHITUNGKAN
$q_add=" and b.KPMM_Level_3='KPMM301000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
//KPMM400000001   LCU
$q_add=" and b.KPMM_Level_3='KPMM400000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
 // end remark 2016-12-22
*/








#--II : 3 Cadangan Umum PPA atas aset produktif yang wajib dibentuk
/*
$query=" SELECT SUM (NILAI)/1000000 AS Jumlah_Nominal FROM 
(
 SELECT SUM(PPA) AS NILAI FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
 WHERE JENIS='PPA ASET PRODUKTIF' AND KOL='1' AND DATADATE='$curr_tgl'
) AS TABEL1 ";
*/
$query=" SELECT SUM (PPA)/1000000 as Jumlah_Nominal FROM DM_KPMM_PPALBU WITH (NOLOCK)
    WHERE Jenis='01'AND datadate='$curr_tgl' and kualitas='1' and SourceForm in ('F05','F06','F07','F07a','F10','F11','F22','F43','F44')
 ";


$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m57=$row['Jumlah_Nominal'];




##--1.2.2.2.2 Rugi tahun-tahun lalu
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM 
(
 SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
 JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
 WHERE a.DataDate='$curr_tgl' AND b.GLNO ='30103002'
 GROUP BY a.kodegl ,b.GLNO
)AS tabel1 ";
$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m31=$row['Jumlah_Nominal'];


#--1.2.2.2.4 Selisih kurang antara Penyisihan Penghapusan Aset (PPA) dan Cadangan Kerugian Penurunan Nilai (CKPN) atas aset produktif (*Note : dibuat positive)
/*
$query=" SELECT SUM (Total_Nilai1)/1000000 AS Total_Nilai FROM 
(
   SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM 
  (
   SELECT SUM (Nilai4) AS Jumlah_Nilai1,SUM (Nilai4A) AS Jumlah_Nilai2 FROM 
   (
    SELECT SUM (CKPN) AS Nilai4,SUM (PPA) AS Nilai4A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
    WHERE SourceData='F11' and Jenis='PPA Aset Produktif'AND datadate='$curr_tgl'
   ) AS TABEL1
  )as total1
UNION
  SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM 
  (
   SELECT SUM (Nilai1) AS Jumlah_Nilai1,SUM (Nilai1A) AS Jumlah_Nilai2 FROM 
   (
   SELECT SUM (CKPN) AS Nilai1,SUM (PPA) AS Nilai1A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
   WHERE SourceData='F05' and Jenis='PPA Aset Produktif'AND Datadate='$curr_tgl'
   ) AS TABEL2
  )as total2
 UNION 
 SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM (
   SELECT SUM (Nilai2) AS Jumlah_Nilai1,SUM (Nilai2A) AS Jumlah_Nilai2 FROM (
   SELECT SUM (CKPN) AS Nilai2,SUM (PPA) AS Nilai2A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
   WHERE SourceData='F06&F10' and Jenis='PPA Aset Produktif'AND datadate='$curr_tgl'
  )AS TABEL3
    )as total3
 UNION
  SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM 
  (
   SELECT SUM (Nilai3) AS Jumlah_Nilai1,SUM (Nilai3A) AS Jumlah_Nilai2 FROM 
   (
    SELECT SUM (CKPN) AS Nilai3,SUM (PPA) AS Nilai3A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
    WHERE SourceData='F07&F08&F09' and Jenis='PPA Aset Produktif'AND datadate='$curr_tgl'
   )AS TABEL4
  )as total4
) AS TOTAL ";
*/

$query=" SELECT SUM(NILAI1-NILAI2)/1000000 AS Total_Nilai FROM 
  (
    SELECT SUM (CKPN) AS NILAI1,SUM(PPA) AS NILAI2 FROM DM_KPMM_PPALBU WITH (NOLOCK)
    WHERE Jenis='01'AND datadate='$curr_tgl'
  ) as TOTAL ";



$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m33=abs($row['Total_Nilai']);

#--1.2.2.2.6 PPA aset non produktif yang wajib dibentuk (*Note : dibuat positive)

/*
$query=" SELECT SUM (Total_Nilai1)/1000000 AS Total_Nilai FROM 
(
   SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM 
  (
   SELECT SUM (Nilai4) AS Jumlah_Nilai1,SUM (Nilai4A) AS Jumlah_Nilai2 FROM (
   SELECT SUM (CKPN) AS Nilai4,SUM (PPA) AS Nilai4A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
   WHERE SourceData='F17' and Jenis='PPA Aset Tidak Produktif'AND datadate='$curr_tgl'
  ) AS TABEL1
  )as total1
UNION
   SELECT SUM(Jumlah_Nilai1-Jumlah_Nilai2) AS Total_Nilai1 FROM (
   SELECT SUM (Nilai1) AS Jumlah_Nilai1,SUM (Nilai1A) AS Jumlah_Nilai2 FROM (
   SELECT SUM (CKPN) AS Nilai1,SUM (PPA) AS Nilai1A FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
   WHERE SourceData='F18' and Jenis='PPA Aset Tidak Produktif'AND Datadate='$curr_tgl'
  ) AS TABEL2
   )as total2
) AS TOTAL ";
*/
$query=" SELECT SUM(NILAI1-NILAI2)/1000000 AS Total_Nilai FROM 
  (
    SELECT SUM (CKPN) AS NILAI1,SUM(PPA) AS NILAI2 FROM DM_KPMM_PPALBU WITH (NOLOCK)
    WHERE Jenis='02'AND datadate='$curr_tgl'
  ) as TOTAL ";

$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m35=abs($row['Total_Nilai']);


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



$objPHPExcel->getActiveSheet()->getStyle('A1:P80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A5:P7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('A9A9A9');
$objPHPExcel->getActiveSheet()->getStyle('A64:P65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('A9A9A9');
$objPHPExcel->getActiveSheet()->getStyle('G66:J66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
$objPHPExcel->getActiveSheet()->getStyle('G72:J72')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
$objPHPExcel->getActiveSheet()->getStyle('M66:P66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
$objPHPExcel->getActiveSheet()->getStyle('M72:N72')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');

$objPHPExcel->getActiveSheet()->getStyle('A9:L9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('9370DB');
$objPHPExcel->getActiveSheet()->getStyle('A54:L54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('9370DB');


$objPHPExcel->getActiveSheet()->getStyle('A1:L7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A5:P7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A9:P12')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A37:P38')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A48:P48')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A54:P58')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A62:P62')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A64:P75')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('A1:P7')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('M4:P7')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('A64:P65')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A64:P65')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('M4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
//$objPHPExcel->getActiveSheet()->getStyle('M4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_LEFT);
//$objPHPExcel->getActiveSheet()->getStyle('A64:P65')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
//$objPHPExcel->getActiveSheet()->getStyle('A64:P65')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A5:P7')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A64:N65')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('M8:P75')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('G66:L75')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A10:A61')->applyFromArray($styleArrayBorder1);
//$objPHPExcel->getActiveSheet()->getStyle('E13:F36')->applyFromArray($styleArrayBorder1);
//objPHPExcel->getActiveSheet()->getStyle('E46:E47')->applyFromArray($styleArrayBorder1);
for ($i=9; $i<=75 ; $i++) { 
$objPHPExcel->getActiveSheet()->getStyle("A$i:L$i")->applyFromArray($styleArrayBorder2);
}


//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(3);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(6);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(7);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(25);
//$objPHPExcel->getActiveSheet()->getRowDimension(33)->setRowHeight(30);


// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:P1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:P2'); 
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:P3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('M4:P4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A5:L7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('M5:N6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('O5:P6');


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A64:F65');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K64:L65');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('M64:N64');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G64:H64');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I64:J64');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('O64:P64');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K66:L66');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K72:L72');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A77:G77');

$objPHPExcel->getActiveSheet()->setCellValue('A1', 'LAPORAN PERHITUNGAN KEWAJIBAN PENYEDIAAN MODAL MINIMUM TRIWULANAN BANK UMUM KONVENSIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'Bank : MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Tanggal : $label_tgl ");

$objPHPExcel->getActiveSheet()->setCellValue('M5', 'Posisi Tgl Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('O5', 'Posisi Tgl Laporan Tahun Sebelumnya ⁴⁾');

$objPHPExcel->getActiveSheet()->setCellValue('M7', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('N7', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('O7', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('P7', 'Konsolidasi');

$objPHPExcel->getActiveSheet()->setCellValue('M4', '(dalam jutaan rupiah)');
$objPHPExcel->getActiveSheet()->setCellValue('A9', 'I');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Modal Inti (Tier 1)');
$objPHPExcel->getActiveSheet()->setCellValue('B10', '1');
$objPHPExcel->getActiveSheet()->setCellValue('C10', "Modal Inti Utama / Common Equity Tier (CET 1)");
$objPHPExcel->getActiveSheet()->setCellValue('C11', '1.1');
$objPHPExcel->getActiveSheet()->setCellValue('D11', 'Modal disetor (Setelah dikurangi treasury stock)');
$objPHPExcel->getActiveSheet()->setCellValue('C12', '1.2');
$objPHPExcel->getActiveSheet()->setCellValue('D12', 'Cadangan Tambahan Modal');
$objPHPExcel->getActiveSheet()->setCellValue('D13', '1.2.1');
$objPHPExcel->getActiveSheet()->setCellValue('E13', 'Faktor  Penambah');
$objPHPExcel->getActiveSheet()->setCellValue('E14', '1.2.1.1');
$objPHPExcel->getActiveSheet()->setCellValue('F14', 'Pendapatan komprehensif lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('F15', '1.2.1.1.1');
$objPHPExcel->getActiveSheet()->setCellValue('G15', 'Selisih lebih penjabaran laporan keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('F16', '1.2.1.1.2');
$objPHPExcel->getActiveSheet()->setCellValue('G16', 'Potensi keuntungan dari peningkatan nilai wajar aset keuangan dalam kelompok tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('F17', '1.2.1.1.3');
$objPHPExcel->getActiveSheet()->setCellValue('G17', 'Saldo surplus revaluasi aset tetap');

$objPHPExcel->getActiveSheet()->setCellValue('E18', '1.2.1.2');
$objPHPExcel->getActiveSheet()->setCellValue('F18', 'Cadangan tambahan modal lainnya (other disclosed reserves)');
$objPHPExcel->getActiveSheet()->setCellValue('F19', '1.2.1.2.1');
$objPHPExcel->getActiveSheet()->setCellValue('G19', 'Agio');
$objPHPExcel->getActiveSheet()->setCellValue('F20', '1.2.1.2.2');
$objPHPExcel->getActiveSheet()->setCellValue('G20', 'Cadangan umum');
$objPHPExcel->getActiveSheet()->setCellValue('F21', '1.2.1.2.3');
$objPHPExcel->getActiveSheet()->setCellValue('G21', 'Laba tahun-tahun lalu');
$objPHPExcel->getActiveSheet()->setCellValue('F22', '1.2.1.2.4');
$objPHPExcel->getActiveSheet()->setCellValue('G22', 'Laba tahun berjalan');
$objPHPExcel->getActiveSheet()->setCellValue('F23', '1.2.1.2.5');
$objPHPExcel->getActiveSheet()->setCellValue('G23', 'Dana setoran modal');
$objPHPExcel->getActiveSheet()->setCellValue('F24', '1.2.1.2.6');
$objPHPExcel->getActiveSheet()->setCellValue('G24', 'Lainnya');

$objPHPExcel->getActiveSheet()->setCellValue('D25', '1.2.2');
$objPHPExcel->getActiveSheet()->setCellValue('E25', 'Faktor  Pengurang  ');
$objPHPExcel->getActiveSheet()->setCellValue('E26', '1.2.2.1');
$objPHPExcel->getActiveSheet()->setCellValue('F26', 'Pendapatan komprehensif lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('F27', '1.2.2.1.1');
$objPHPExcel->getActiveSheet()->setCellValue('G27', 'Selisih kurang penjabaran laporan keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('F28', '1.2.2.1.2');
$objPHPExcel->getActiveSheet()->setCellValue('G28', 'Potensi kerugian dari penurunan nilai wajar aset keuangan dalam kelompok tersedia untuk dijual');

$objPHPExcel->getActiveSheet()->setCellValue('E29', '1.2.2.2');
$objPHPExcel->getActiveSheet()->setCellValue('F29', 'Cadangan tambahan modal lainnya (other disclosed reserves)');
$objPHPExcel->getActiveSheet()->setCellValue('F30', '1.2.2.2.1');
$objPHPExcel->getActiveSheet()->setCellValue('G30', 'Disagio');
$objPHPExcel->getActiveSheet()->setCellValue('F31', '1.2.2.2.2');
$objPHPExcel->getActiveSheet()->setCellValue('G31', 'Rugi tahun-tahun lalu');
$objPHPExcel->getActiveSheet()->setCellValue('F32', '1.2.2.2.3');
$objPHPExcel->getActiveSheet()->setCellValue('G32', 'Rugi tahun berjalan');
$objPHPExcel->getActiveSheet()->setCellValue('F33', '1.2.2.2.4');
$objPHPExcel->getActiveSheet()->setCellValue('G33', 'Selisih kurang antara Penyisihan Penghapusan Aset (PPA) dan Cadangan Kerugian Penurunan Nilai (CKPN) atas aset produktif');
$objPHPExcel->getActiveSheet()->setCellValue('F34', '1.2.2.2.5');
$objPHPExcel->getActiveSheet()->setCellValue('G34', 'Selisih kurang jumlah penyesuaian nilai wajar dari instrumen keuangan dalam Trading Book');
$objPHPExcel->getActiveSheet()->setCellValue('F35', '1.2.2.2.6');
$objPHPExcel->getActiveSheet()->setCellValue('G35', 'PPA aset non produktif yang wajib dibentuk');
$objPHPExcel->getActiveSheet()->setCellValue('F36', '1.2.2.2.7');
$objPHPExcel->getActiveSheet()->setCellValue('G36', 'Lainnya');

$objPHPExcel->getActiveSheet()->setCellValue('C37', '1.3');
$objPHPExcel->getActiveSheet()->setCellValue('D37', 'Kepentingan Non Pengendali yang dapat diperhitungkan');
$objPHPExcel->getActiveSheet()->setCellValue('C38', '1.4');
$objPHPExcel->getActiveSheet()->setCellValue('D38', 'Faktor Pengurang Modal Inti Utama');

$objPHPExcel->getActiveSheet()->setCellValue('D39', '1.4.1');
$objPHPExcel->getActiveSheet()->setCellValue('E39', 'Perhitungan pajak tangguhan');
$objPHPExcel->getActiveSheet()->setCellValue('D40', '1.4.2');
$objPHPExcel->getActiveSheet()->setCellValue('E40', 'Goodwill');
$objPHPExcel->getActiveSheet()->setCellValue('D41', '1.4.3');
$objPHPExcel->getActiveSheet()->setCellValue('E41', 'Seluruh aset tidak berwujud lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('D42', '1.4.4');
$objPHPExcel->getActiveSheet()->setCellValue('E42', 'Penyertaan yang diperhitungkan sebagai faktor pengurang');
$objPHPExcel->getActiveSheet()->setCellValue('D43', '1.4.5');
$objPHPExcel->getActiveSheet()->setCellValue('E43', 'Kekurangan modal pada perusahaan anak asuransi');
$objPHPExcel->getActiveSheet()->setCellValue('D44', '1.4.6');
$objPHPExcel->getActiveSheet()->setCellValue('E44', 'Eksposur sekuritisasi');
$objPHPExcel->getActiveSheet()->setCellValue('D45', '1.4.7');
$objPHPExcel->getActiveSheet()->setCellValue('E45', 'Faktor Pengurang modal inti utama lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('E46', '1.4.7.1');
$objPHPExcel->getActiveSheet()->setCellValue('F46', 'Pendapatan dana pada instrumen AT 1 dan/atau Tier 2 pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('E47', '1.4.7.2');
$objPHPExcel->getActiveSheet()->setCellValue('F47', 'Kepemilikan silang pada entitas lain yang diperoleh berdasarkan peralihan karena hukum, hibah, atau hibah wasiat');

$objPHPExcel->getActiveSheet()->setCellValue('B48', '2');
$objPHPExcel->getActiveSheet()->setCellValue('C48', "Modal Inti Tambahan / Additional Tier (AT-1)");
$objPHPExcel->getActiveSheet()->setCellValue('C49', '2.1');
$objPHPExcel->getActiveSheet()->setCellValue('D49', 'Instrumen yang memenuhi persyaratan AT-1');
$objPHPExcel->getActiveSheet()->setCellValue('C50', '2.2');
$objPHPExcel->getActiveSheet()->setCellValue('D50', 'Agio / Disagio');
$objPHPExcel->getActiveSheet()->setCellValue('C51', '2.3');
$objPHPExcel->getActiveSheet()->setCellValue('D51', 'Faktor Pengurang Modal Inti Tambahan');
$objPHPExcel->getActiveSheet()->setCellValue('D52', '2.3.1');
$objPHPExcel->getActiveSheet()->setCellValue('E52', 'Penempatan dana pada instrumen AT 1 dan/atau Tier 2 pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('D53', '2.3.2');
$objPHPExcel->getActiveSheet()->setCellValue('E53', 'Kepemilikan silang pada entitas lain yang diperoleh berdasarkan peralihan karena hukum, hibah, atau hibah wasiat');

$objPHPExcel->getActiveSheet()->setCellValue('A54', 'II');
$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Modal Pelengkap (Tier 2)');
$objPHPExcel->getActiveSheet()->setCellValue('B55', '1');
$objPHPExcel->getActiveSheet()->setCellValue('C55', "Instrumen modal dalam bentuk saham atau lainnya yang memenuhi persyaratan Tier 2");
$objPHPExcel->getActiveSheet()->setCellValue('B56', '2');
$objPHPExcel->getActiveSheet()->setCellValue('C56', "Agio / disagio");
$objPHPExcel->getActiveSheet()->setCellValue('B57', '3');
$objPHPExcel->getActiveSheet()->setCellValue('C57', "Cadangan umum PPA atas aset produktif yang wajib dibentuk (paling tinggi 1,25 % ATMR Risiko Kredit )");
$objPHPExcel->getActiveSheet()->setCellValue('B58', '4');
$objPHPExcel->getActiveSheet()->setCellValue('C58', "Faktor Pengurang Modal Pelengkap");
$objPHPExcel->getActiveSheet()->setCellValue('C59', '4.1');
$objPHPExcel->getActiveSheet()->setCellValue('D59', 'Sinking Fund');
$objPHPExcel->getActiveSheet()->setCellValue('C60', '4.2');
$objPHPExcel->getActiveSheet()->setCellValue('D60', 'Penempatan dana pada instrumen Tier 2 pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('C61', '4.3');
$objPHPExcel->getActiveSheet()->setCellValue('D61', 'Kepemilikan silang pada entitas lain yang diperoleh berdasarkan peralihan karena hukum, hibah, atau hibah wasiat');

$objPHPExcel->getActiveSheet()->setCellValue('A62', 'Total Modal (I+II)');


$objPHPExcel->getActiveSheet()->setCellValue('G64', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('I64', 'Posisi Tanggal Laporan Tahun Sebelumnya');
$objPHPExcel->getActiveSheet()->setCellValue('G65', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('H65', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('I65', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('J65', 'Konsolidasi');
$objPHPExcel->getActiveSheet()->setCellValue('K64', 'KETERANGAN');
$objPHPExcel->getActiveSheet()->setCellValue('M64', 'Posisi Tanggal Laporan');
$objPHPExcel->getActiveSheet()->setCellValue('M65', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('N65', 'Konsolidasi');

$objPHPExcel->getActiveSheet()->setCellValue('O64', 'Posisi Tanggal Laporan Tahun Sebelumnya');
$objPHPExcel->getActiveSheet()->setCellValue('O65', 'Bank');
$objPHPExcel->getActiveSheet()->setCellValue('P65', 'Konsolidasi');

$objPHPExcel->getActiveSheet()->setCellValue('K66', 'RASIO KPMM');
$objPHPExcel->getActiveSheet()->setCellValue('L67', 'Rasio CET1 (%)');
$objPHPExcel->getActiveSheet()->setCellValue('L68', 'Rasio Tier 1 (%)');
$objPHPExcel->getActiveSheet()->setCellValue('L69', 'Rasio Tier 2 (%)');
$objPHPExcel->getActiveSheet()->setCellValue('L70', 'Rasio KPMM (%)');
$objPHPExcel->getActiveSheet()->setCellValue('K71', 'CET 1 UNTUK BUFFER (%)');

$objPHPExcel->getActiveSheet()->setCellValue('A66', 'ASET TERTIMBANG MENURUT RISIKO');
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'ATMR RISIKO KREDIT ₁₎');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'ATMR RISIKO PASAR');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'ATMR RISIKO OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'TOTAL ATMR');
$objPHPExcel->getActiveSheet()->setCellValue('A71', 'RASIO KPMM SESUAI PROFIL RISIKO (%) ');

$objPHPExcel->getActiveSheet()->setCellValue('A72', 'ALOKASI PEMENUHAN KPMM');
$objPHPExcel->getActiveSheet()->setCellValue('K72', 'PERSENTASE BUFFER YANG WAJIB DIPENUHI OLEH BANK (%)');
$objPHPExcel->getActiveSheet()->setCellValue('B73', 'Dari CET1');
$objPHPExcel->getActiveSheet()->setCellValue('L73', 'Capital Conservation Buffer (%)');
$objPHPExcel->getActiveSheet()->setCellValue('B74', 'Dari AT1');
$objPHPExcel->getActiveSheet()->setCellValue('L74', 'Countercyclical Buffer (%)');
$objPHPExcel->getActiveSheet()->setCellValue('B75', 'Dari Tier 1');
$objPHPExcel->getActiveSheet()->setCellValue('L75', 'Capital Surcharge untuk Bank Sistemik (%)');
$objPHPExcel->getActiveSheet()->setCellValue('A77', '*** Sesuai dengan profil resiko Bank MNC Internasional');


#fill ATMR dan jumlah dari ATMR

/* remak 2016-12-22
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
 */ // remak 2016-12-22  
$objPHPExcel->getActiveSheet()->setCellValue('M11', $m11);
$objPHPExcel->getActiveSheet()->setCellValue('M16', $m16);
$objPHPExcel->getActiveSheet()->setCellValue('M19', $m19);
$objPHPExcel->getActiveSheet()->setCellValue('M20', $m20);
$objPHPExcel->getActiveSheet()->setCellValue('M22', $m22);
$objPHPExcel->getActiveSheet()->setCellValue('M23', $m23);
$objPHPExcel->getActiveSheet()->setCellValue('M28', $m28);
$objPHPExcel->getActiveSheet()->setCellValue('M31', $m31);
$objPHPExcel->getActiveSheet()->setCellValue('M33', $m33);
$objPHPExcel->getActiveSheet()->setCellValue('M35', $m35);
$objPHPExcel->getActiveSheet()->setCellValue('M39', $m39);
$objPHPExcel->getActiveSheet()->setCellValue('M41', $m41);
$objPHPExcel->getActiveSheet()->setCellValue('M57', $m57);


$objPHPExcel->getActiveSheet()->setCellValue('M9', "=M10+M48");
$objPHPExcel->getActiveSheet()->setCellValue('M10', "=M11+M12+M37-M38");
$objPHPExcel->getActiveSheet()->setCellValue('M12', "=M13-M25");
$objPHPExcel->getActiveSheet()->setCellValue('M13', "=M14+M18");


$objPHPExcel->getActiveSheet()->setCellValue('M14', "=SUM(M15:M17)");
$objPHPExcel->getActiveSheet()->setCellValue('M18', "=SUM(M19:M24)");
$objPHPExcel->getActiveSheet()->setCellValue('M25', "=M26+M29");
$objPHPExcel->getActiveSheet()->setCellValue('M26', "=SUM(M27:M28)");
$objPHPExcel->getActiveSheet()->setCellValue('M29', "=SUM(M30:M36)");
$objPHPExcel->getActiveSheet()->setCellValue('M38', "=SUM(M39:M45)");
$objPHPExcel->getActiveSheet()->setCellValue('M48', "=SUM(M49:M51)");
$objPHPExcel->getActiveSheet()->setCellValue('M54', "=SUM(M55:M58)");

$objPHPExcel->getActiveSheet()->setCellValue('M58', "=SUM(M59:M61)");
$objPHPExcel->getActiveSheet()->setCellValue('M62', "=M54+M9");

$objPHPExcel->getActiveSheet()->setCellValue('G71', "=ABS(+IF(M62=0,0,(M62/G70)))");







$objPHPExcel->getActiveSheet()->setCellValue('G67', $atmr_kredit);
$objPHPExcel->getActiveSheet()->setCellValue('G68', $atmr_pasar);
$objPHPExcel->getActiveSheet()->setCellValue('G69', $atmr_operasional);
$objPHPExcel->getActiveSheet()->setCellValue('G70', "=SUM(G67:G69)");
//$objPHPExcel->getActiveSheet()->setCellValue('G73', $g73);
//$objPHPExcel->getActiveSheet()->setCellValue('G74', $g74);
//$objPHPExcel->getActiveSheet()->setCellValue('G75', $c75);


$objPHPExcel->getActiveSheet()->setCellValue('M67', "=M10/G70");
$objPHPExcel->getActiveSheet()->setCellValue('M68', "=M9/G70");
$objPHPExcel->getActiveSheet()->setCellValue('M69', "=M54/G70");
$objPHPExcel->getActiveSheet()->setCellValue('M70', "=M62/G70");
$objPHPExcel->getActiveSheet()->setCellValue('M71', "=ABS(M67-G73)");

$objPHPExcel->getActiveSheet()->setCellValue('G73', "=H77-G75");
$objPHPExcel->getActiveSheet()->setCellValue('G75', "=M69");
$objPHPExcel->getActiveSheet()->setCellValue('H77', 0.1);

$objPHPExcel->getActiveSheet()->getStyle('G71:J71')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));  
$objPHPExcel->getActiveSheet()->getStyle('M67:M75')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 )); 
$objPHPExcel->getActiveSheet()->getStyle('G73:J75')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));  
$objPHPExcel->getActiveSheet()->getStyle('H77')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));     
//$objPHPExcel->getActiveSheet()->getStyle('A9:L9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('483D8B');
//$objPHPExcel->getActiveSheet()->getStyle('A41:L41')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('483D8B');
//$objPHPExcel->getActiveSheet()->getStyle('A5:L7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
//$objPHPExcel->getActiveSheet()->getStyle('A50:L52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');

//$objPHPExcel->getActiveSheet()->getStyle('C17:H22')->getNumberFormat()->setFormatCode('0.00');
//$objPHPExcel->getActiveSheet()->getStyle('C29:H34')->getNumberFormat()->setFormatCode('0.00');
 
$objPHPExcel->getActiveSheet()->getStyle('M9:P62')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('M67:N75')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G74')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('G67:J70')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('G73:J75')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

for ($i=9;$i<=62;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, 0);
    }
}
for ($i=9;$i<=62;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('N'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('N'.$i, 0);
    }
}

for ($i=9;$i<=62;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('O'.$i, 0);
    }
}
for ($i=9;$i<=62;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('P'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('P'.$i, 0);
    }
}



for ($i=67;$i<=75;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, 0);
    }
}
/*
for ($i=67;$i<=75;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('N'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('N'.$i, 0);
    }
}
*/
for ($i=73;$i<=75;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}


$objPHPExcel->getActiveSheet()->setTitle('KPMM');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/KPMM_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/KPMM_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>

<div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> List KPMM 
                            </div>
                            <!--<div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>

                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div>-->
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
                                            <b>
												<div class="pull-right" style="font-size:12px">
													<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/KPMM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> 
													</a> 
												</div>
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
                                                <td width="80%" align="center" rowspan="2" colspan="3"><b></b></td>
                                                <td width="20%" align="center" colspan="2"><b> Posisi Tgl Laporan </b></td>
                                                <td width="20%" align="center" colspan="2"><b> Posisi Tgl Laporan Tahun Sebelumnya ⁴⁾</b></td>
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

                                               for ($i=9; $i <= 63 ; $i++) { 
                                               	 if ($i=='10' || $i=='48' || $i=='55' || $i=='56' || $i=='57' || $i=='58' ){

                                               	 $class_tr=" ";

                                               	} else {
                                               	 $class_tr=" ";

                                               	}

                                                 ?>
                                                 <?php
												if ($i=='9' || $i=='54' || $i=='62'){
												?>
												<tr class="success">
                                                <td  align="left" width="80%" colspan="3"> <b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("O$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("P$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                  <?php
                                            		} else {
                                                ?>
                                                <tr <?php echo $class_tr;?>>
                                                <td  align="left" width="5%"> <b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td  align="left" width="3%"> <b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                <td  align="left" width="72%"> <b>
                                                <?php 
                                                echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); 
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                ?></b></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("O$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("P$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
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
                                                <td width="17%" align="center" colspan="2"><br><b> <?php echo $label_tgl;?> </b></br></td>
                                                <td width="20%" align="center" colspan="2"><br><b> Posisi Tanggal Laporan Tahun Sebelumnya </b></br></td>
                                                <td width="15%" align="center" rowspan="2"><br><br><b> Keterangan </br></br></b></td>
                                                <td width="17%" align="center" colspan="2"><br><b> Posisi Tgl Laporan </b></br></td>
                                                <td width="17%" align="center" colspan="2"><br><b> Posisi Tgl Laporan Tahun Sebelumnya</b></br></td>
                                                
                                                </tr>
                                                <tr class="active">
                                                <td width="7%" align="center"><b>Bank</b></td>
                                                <td width="7%" align="center"><b>Konsolidasi</b></td>
                                                <td width="10%" align="center"><b>Bank</b></td>
                                                <td width="10%" align="center"><b>Konsolidasi</b></td>
                                                <td width="7%" align="center"><b>Bank</b></td>
                                                <td width="7%" align="center"><b>Konsolidasi</b></td>
                                                <td width="7%" align="center"><b>Bank</b></td>
                                                <td width="7%" align="center"><b>Konsolidasi</b></td>
                                                
                                                </tr>
                                                </thead>



                                                <tbody>
                                                <?php
                                                for ($i=66; $i <=75 ; $i++) { 

                                                	if ($i=='66'){
                                                		$varclass="bgcolor='#2F353B'";
                                                    $varclass2="bgcolor='#2F353B'";
                                                    $varclass3="bgcolor='#2F353B'";

                                                	} else if ($i=='72') {

                                                		$varclass="bgcolor='#2F353B'";
                                                    $varclass2="bgcolor='#2F353B'";
                                                    $varclass3="";
                                                	} else {
                                                    $varclass="";
                                                    $varclass2="";
                                                    $varclass3="";

                                                  }
                                                
                                                ?>
                                                <tr >
                                                <td width="25%" align="left"><b>
                                                 <?php 
                                                echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); 
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                
                                                ?>

                                                </b></td>
                                                <td width="8%" align="right" <?php echo $varclass;?> ><b>  <?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass;?> ><b><?php echo "<div ".$varclass." >";?> <?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="23%" align="left"><b><i><?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');
                                                echo " ".$objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></i></b></td>
                                                <td width="8%" align="right" <?php echo $varclass2;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass2;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass3;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("O$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td width="8%" align="right" <?php echo $varclass3;?>><b><?php echo $objPHPExcel->getActiveSheet()->getCell("P$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                
                                                </tr>
                                               <?php
                                           }
                                               ?>
                                                 
                                                </tbody>
                                            </table>
                                        </div>        
                                        *** Sesuai dengan profil resiko Bank MNC Internasional 10% 






                                    </div>
                                  
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

