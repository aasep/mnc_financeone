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
logActivity("generate ltv",date('Y_m_d_H_i_s'));
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




$var_tabel=date('Ymd',strtotime($tanggal));

#############################################################################################
$table_asetkredit="DM_AsetKredit_$var_tabel";
//$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
//$table_deposito="DM_LiabilitasDeposito_$var_tabel";
//$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

$qcekTable=" select top 1* from  $table_asetkredit ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_asetkredit Tidak tersedia ... ! </b> <br><br></div>";
        die();
}
/*
$qcekTable=" select top 1* from  $table_tabungan ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_tabungan Tidak tersedia ... ! </b> <br><br></div>";
        die();
}

$qcekTable=" select top 1* from  $table_deposito ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_deposito Tidak tersedia ... ! </b> <br><br></div>";
        die();
}

$qcekTable=" select top 1* from  $table_banklain ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_banklain Tidak tersedia ... ! </b> <br><br></div>";
        die();
}

*/

//echo $table_banklain;
//die();

#############################################################################################


##############################################################################################################################################################
#-- Query LTV Malida(Update 13/07/2016)--

//--Rumah Tinggal Tipe < 21--
//--Jumlah--

$var_add_query =" and a.status NOT IN ('2','8') ";

$query_jml = "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='01' AND  a.DataDate='$curr_tgl'  $var_add_query ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d11=$row_jml['jml'];




//--NPL--
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='01' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl' $var_add_query ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e11=$row_npl['jml_npl'];

//--Rumah Tinggal Tipe 22-70--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='02' AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d12=$row_jml['jml'];
//--NPL--
$query_npl=" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='02' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e12=$row_npl['jml_npl'];

//--Rumah Tinggal Tipe > 70--
//--Jumlah--
$query_jml = "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='03' AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d13=$row_jml['jml'];
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='03' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e13=$row_npl['jml_npl'];

//--Rumah Susun Tipe < 21--
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='04' AND a.DataDate='$curr_tgl' $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d15=$row_jml['jml'];
//--NPL--
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='04' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e15=$row_npl['jml_npl'];

//echo $query_npl;
//die();

//--Rumah Susun Tipe 22-70--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='05' AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d16=$row_jml['jml'];
//--NPL--
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='05' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e16=$row_npl['jml_npl'];


//--Rumah Susun Tipe > 70--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='06' AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d17=$row_jml['jml'];
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='06' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e17=$row_npl['jml_npl'];

//--Ruko/Rukan--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='07' AND a.DataDate='$curr_tgl'  $var_add_query  ";
$res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d18=$row_jml['jml'];
//--NPL--
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='07' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
$res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e18=$row_npl['jml_npl'];




//--KKBP (Kredit Konsumsi Beragun Properti)--
//--Jumlah-- Rumah Tinggal
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE C.KodeLTV ='08' AND c.KodeLTV2='001' AND d.KodeLTV2='001' AND a.DataDate='$curr_tgl' $var_add_query ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d21=$row_jml['jml'];

//--NPL--

$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE C.KodeLTV ='08' AND c.KodeLTV2='001' AND d.KodeLTV2='001' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl' AND a.status not in ('2','8') ";

        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e21=$row_npl['jml_npl'];
//echo $e21;
//die();
#Apartemen

$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE C.KodeLTV ='08' AND c.KodeLTV2='002' AND d.KodeLTV2='002' AND a.DataDate='$curr_tgl' $var_add_query ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d22=$row_jml['jml'];


//echo $query_jml;
//die();
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE C.KodeLTV ='08' AND c.KodeLTV2='002' AND d.KodeLTV2='002' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl' AND a.status not in ('2','8')  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e22=$row_npl['jml_npl'];



#rumah toko


$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE C.KodeLTV ='08' AND c.KodeLTV2='003' AND d.KodeLTV2='003' AND a.DataDate='$curr_tgl'AND a.status not in ('2','8') $var_add_query  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d23=$row_jml['jml'];



//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE c.KodeLTV ='08' AND c.KodeLTV2='003' AND d.KodeLTV2='003' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='curr_tgl' $var_add_query ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e23=$row_npl['jml_npl'];









//--Total KPR/KKBP dengan nominal s.d Rp5 Juta--
//--Jumlah--
$query_jml =" SELECT SUM(Total) as jml FROM (
SELECT SUM (a.JumlahKreditPeriodeLaporan)AS Total
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV IN('01','02','03','04','05','06','07') AND a.DataDate='$curr_tgl'AND a.JumlahKreditPeriodeLaporan <='5000000' AND a.status not in ('2','8')
UNION ALL
SELECT SUM (a.JumlahKreditPeriodeLaporan)AS Total
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE c.KodeLTV ='08' AND c.KodeLTV2 IN ('001','002','003') AND d.KodeLTV2 IN('001','002','003') AND a.JumlahKreditPeriodeLaporan <='5000000' AND a.DataDate='$curr_tgl' AND a.status not in ('2','8')
)AS Table1 ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d25=$row_jml['jml'];



//--NPL--
$query_npl= " SELECT SUM(Total) as jml_npl FROM (
SELECT SUM (a.JumlahKreditPeriodeLaporan)AS Total
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV IN('01','02','03','04','05','06','07') AND a.DataDate='$curr_tgl' AND a.JumlahKreditPeriodeLaporan <='5000000' AND a.Kolektibilitas IN ('3','4','5') AND a.status not in ('2','8')
UNION ALL
SELECT SUM (a.JumlahKreditPeriodeLaporan)AS Total
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV 
JOIN Ref_BIGOID_JenisAgunanJaminan d ON d.KodeInternal = a.JenisAgunan AND d.KodeLTV2 = c.KodeLTV2
WHERE c.KodeLTV ='08' AND c.KodeLTV2 IN ('001','002','003') AND d.KodeLTV2 IN('001','002','003') AND a.JumlahKreditPeriodeLaporan <='5000000' AND a.DataDate='$curr_tgl' AND a.Kolektibilitas IN ('3','4','5') AND a.status not in ('2','8')
)AS Table1 ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e25=$row_npl['jml_npl'];
      //  echo $query_jml ."<br>";
       // echo $query_npl;
      // die();


 /*       
//--KKB(Rumah Tangga untuk Pemilikan Mobil Roda Empat)--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='09' AND a.DataDate='$curr_tgl'  $var_add_query  ";
$res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d26=$row_jml['jml'];
//--NPL-- 
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='09' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
$res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e26=$row_npl['jml_npl'];


//--KKB(Rumah Tangga untuk Pemilikan Sepeda Bermotor)--
//--Jumlah--
$query_jml ="SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='10' AND a.DataDate='$curr_tgl'  $var_add_query   ";
$res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d27=$row_jml['jml'];
//--NPL--
$query_npl= "SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='10' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  $var_add_query  ";
$res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e27=$row_npl['jml_npl'];
//--Query LTV Baru--
*/
##########################################################################################################################################################################
/* 
28       -KKB(Rumah Tangga untuk Pemilikan Mobil Roda Empat)--
--Jumlah--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002100' AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
--NPL--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002100' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'

29
--KKB(Rumah Tangga untuk Pemilikan Sepeda Bermotor)--
--Jumlah--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002200' AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
--NPL--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002200' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'

30
--KKB (Rumah Tangga untuk Pemilikan Truk dan Kendaraan Bermotor Roda Enam atau Lebih)--
--Jumlah--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002300' AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
--NPL--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002300' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'

31
--KKB (Rumah Tangga untuk Pemilikan Kendaraan Bermotor Lainnya)--
--Jumlah--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002900' AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
--NPL--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi='002900' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'

32
--Total KKB dengan nominal s.d Rp5 Juta--
--Jumlah--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi IN ('002100','002200','002300','002900') AND a.JumlahKreditPeriodeLaporan <= '5000000' AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
--NPL--
SELECT SUM (a.JumlahKreditPeriodeLaporan)
FROM DM_AsetKredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
WHERE a.SektorEkonomi IN ('002100','002200','002300','002900') AND a.JumlahKreditPeriodeLaporan <= '5000000' AND a.Kolektibilitas IN ('3','4','5')AND a.DataDate='2016-05-31' AND b.StatusLTV='Y'
*/
##########################################################################################################################################################################


//--KKB(Rumah Tangga untuk Pemilikan Mobil Roda Empat)--
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='09' AND a.DataDate='$curr_tgl' AND a.status not in ('2','8')  ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d28=$row_jml['jml'];
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='09' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl' ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e28=$row_npl['jml_npl'];


 //echo  $query_npl;
 //die();      
//--KKB(Rumah Tangga untuk Pemilikan Sepeda Bermotor)--
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='10' AND a.DataDate='$curr_tgl' AND a.status not in ('2','8') ";
$res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d29=$row_jml['jml'];
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='10' AND a.Kolektibilitas IN ('3','4','5')AND a.DataDate='$curr_tgl'";
$res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e29=$row_npl['jml_npl'];


//----KKB (Rumah Tangga untuk Pemilikan Truk dan Kendaraan Bermotor Roda Enam atau Lebih)---
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='11' AND a.DataDate='$curr_tgl' AND a.status not in ('2','8') ";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d30=$row_jml['jml'];
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='11' AND a.Kolektibilitas IN ('3','4','5') AND a.DataDate='$curr_tgl'  ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e30=$row_npl['jml_npl'];

//31
//--KKB (Rumah Tangga untuk Pemilikan Kendaraan Bermotor Lainnya)--
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='12' AND a.DataDate='$curr_tgl' AND a.status not in ('2','8')";
        $res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d31=$row_jml['jml'];
//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV ='12' AND a.Kolektibilitas IN ('3','4','5')AND a.DataDate='$curr_tgl' ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e31=$row_npl['jml_npl'];
//32
//--Total KKB dengan nominal s.d Rp5 Juta--
//--Jumlah--
$query_jml =" SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV IN('09','10','11','12') AND a.JumlahKreditPeriodeLaporan <= '5000000' AND a.DataDate='$curr_tgl'  ";
$res_jml=odbc_exec($connection2, $query_jml);
        $row_jml=odbc_fetch_array($res_jml);
        $d32=$row_jml['jml'];

//--NPL--
$query_npl= " SELECT SUM (a.JumlahKreditPeriodeLaporan) as jml_npl
FROM $table_asetkredit a
JOIN Ref_BIGOID_SektorEkonomi b ON b.KodeInternal= a.SektorEkonomi 
JOIN Referensi_LTV c ON c.KodeLTV = b.KodeLTV
WHERE c.KodeLTV IN('09','10','11','12') AND a.JumlahKreditPeriodeLaporan <= '5000000' AND a.Kolektibilitas IN ('3','4','5')AND a.DataDate='$curr_tgl' ";
        $res_npl=odbc_exec($connection2, $query_npl);
        $row_npl=odbc_fetch_array($res_npl);
        $e32=$row_npl['jml_npl'];

##############################################################################################################################################################

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
$objPHPExcel->getActiveSheet()->getStyle('D7')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('D8')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('E8')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C1:E5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A8:E8')->applyFromArray($styleArrayFontBold);

//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('A8:E33')->applyFromArray($styleArrayBorder1);




//FILL COLOR
$objPHPExcel->getActiveSheet()->getStyle('A10:E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');
$objPHPExcel->getActiveSheet()->getStyle('A14:E14')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');
$objPHPExcel->getActiveSheet()->getStyle('A18:E18')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');
$objPHPExcel->getActiveSheet()->getStyle('A27:E27')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');

//$objPHPExcel->getActiveSheet()->getStyle('A23:E23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');

$objPHPExcel->getActiveSheet()->getStyle('A1:Z7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A34:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
/*
$objPHPExcel->getActiveSheet()->getStyle('A1:Z12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A58:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('L13:Z57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F9:Z10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A13:A57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A9:A10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
*/
//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(80);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(25);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D7:E7');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B8:C8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B9:C9');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B10:C10');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B14:C14');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B18:C18');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B19:C19');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B20:C20');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B25:C25');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B26:C26');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B27:C27');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B32:C32');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B33:C33');



$objPHPExcel->getActiveSheet()->setCellValue('C1', 'LAPORAN KREDIT PROPERTI (KP) DAN KREDIT KENDARAAN BERMOTOR (KKB)');
$objPHPExcel->getActiveSheet()->setCellValue('C2', 'SANDI BANK :');
$objPHPExcel->getActiveSheet()->setCellValue('C3', 'KEGIATAN BANK :');
$objPHPExcel->getActiveSheet()->setCellValue('C4', 'TAHUN :');
$objPHPExcel->getActiveSheet()->setCellValue('C5', 'BULAN :');
$objPHPExcel->getActiveSheet()->setCellValue('D4', 'TAHUN :');
$objPHPExcel->getActiveSheet()->setCellValue('D5', 'BULAN :');


$objPHPExcel->getActiveSheet()->setCellValue('D2', '485-Bank MNC Internasional');
$objPHPExcel->getActiveSheet()->setCellValue('D3', 'Bank Umum Konvensional');
$objPHPExcel->getActiveSheet()->setCellValue('D4', $year_modal);
$objPHPExcel->getActiveSheet()->setCellValue('D5', $mon_modal);
$objPHPExcel->getActiveSheet()->setCellValue('D7', '(dlm Rp)');

$objPHPExcel->getActiveSheet()->setCellValue('A8', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'JENIS KREDIT');
$objPHPExcel->getActiveSheet()->setCellValue('D8', 'Jumlah');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'NPL');

$objPHPExcel->getActiveSheet()->setCellValue('A9', '1');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'KPR ( Kredit Pemilikan Rumah )');

$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Rumah Tinggal');
$objPHPExcel->getActiveSheet()->setCellValue('C11', 'Tipe < 21');
$objPHPExcel->getActiveSheet()->setCellValue('C12', 'Tipe 22-70');
$objPHPExcel->getActiveSheet()->setCellValue('C13', 'Tipe > 70');

$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Rumah Susun');
$objPHPExcel->getActiveSheet()->setCellValue('C15', 'Tipe < 21');
$objPHPExcel->getActiveSheet()->setCellValue('C16', 'Tipe 22-70');
$objPHPExcel->getActiveSheet()->setCellValue('C17', 'Tipe > 70');

$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Ruko / Rukan');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Total KPR');
$objPHPExcel->getActiveSheet()->setCellValue('A20', '2');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'KKBP (Kredit Konsumsi Beragun Properti)');



$objPHPExcel->getActiveSheet()->setCellValue('C21', 'Rumah Tinggal');
$objPHPExcel->getActiveSheet()->setCellValue('C22', 'Apartemen / Rumah Susun');
$objPHPExcel->getActiveSheet()->setCellValue('C23', 'Rumah Toko / Rumah Kantor)');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'Total KKBP');


$objPHPExcel->getActiveSheet()->setCellValue('A25', '3');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'Total KPR/KKBP dengan nominal s.d Rp5 Juta');
$objPHPExcel->getActiveSheet()->setCellValue('A26', '4');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'Total KP (Kredit Properti)');

$objPHPExcel->getActiveSheet()->setCellValue('A27', '5');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'KKB (Kredit Kendaraan Bermotor)');

$objPHPExcel->getActiveSheet()->setCellValue('C28', 'Rumah Tangga untuk Pemilikan Mobil Roda Empat');
$objPHPExcel->getActiveSheet()->setCellValue('C29', 'Rumah Tangga untuk Pemilikan Sepeda Bermotor');
$objPHPExcel->getActiveSheet()->setCellValue('C30', 'Rumah Tangga untuk Pemilikan Truk dan Kendaraan Bermotor Roda Enam atau Lebih');
$objPHPExcel->getActiveSheet()->setCellValue('C31', 'Rumah Tangga untuk Pemilikan Kendaraan Bermotor Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('A32', '6');
$objPHPExcel->getActiveSheet()->setCellValue('B32', 'Total KKB dengan nominal s.d Rp5 Juta');
$objPHPExcel->getActiveSheet()->setCellValue('A33', '7');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'Total KKB');



############################## 1 Bulan RP ###################################
$objPHPExcel->getActiveSheet()->setCellValue('D11', floatval($d11));
$objPHPExcel->getActiveSheet()->setCellValue('E11', floatval($e11));
$objPHPExcel->getActiveSheet()->setCellValue('D12', floatval($d12));
$objPHPExcel->getActiveSheet()->setCellValue('E12', floatval($e12));
$objPHPExcel->getActiveSheet()->setCellValue('D13', floatval($d13));
$objPHPExcel->getActiveSheet()->setCellValue('E13', floatval($e13));
$objPHPExcel->getActiveSheet()->setCellValue('D15', floatval($d15));
$objPHPExcel->getActiveSheet()->setCellValue('E15', floatval($e15));
$objPHPExcel->getActiveSheet()->setCellValue('D16', floatval($d16));
$objPHPExcel->getActiveSheet()->setCellValue('E16', floatval($e16));
$objPHPExcel->getActiveSheet()->setCellValue('D17', floatval($d17));
$objPHPExcel->getActiveSheet()->setCellValue('E17', floatval($e17));

$objPHPExcel->getActiveSheet()->setCellValue('D18', floatval($d18));
$objPHPExcel->getActiveSheet()->setCellValue('E18', floatval($e18));
$objPHPExcel->getActiveSheet()->setCellValue('D19', "=SUM(D11:D18)");
$objPHPExcel->getActiveSheet()->setCellValue('E19', "=SUM(E11:E18)");
//$objPHPExcel->getActiveSheet()->setCellValue('D20', floatval($d20));
//$objPHPExcel->getActiveSheet()->setCellValue('E20', floatval($e20));

$objPHPExcel->getActiveSheet()->setCellValue('D21', floatval($d21));
$objPHPExcel->getActiveSheet()->setCellValue('E21', floatval($e21));
$objPHPExcel->getActiveSheet()->setCellValue('D22', floatval($d22));
$objPHPExcel->getActiveSheet()->setCellValue('E22', floatval($e22));
$objPHPExcel->getActiveSheet()->setCellValue('D23', floatval($d23));
$objPHPExcel->getActiveSheet()->setCellValue('E23', floatval($e23));






$objPHPExcel->getActiveSheet()->setCellValue('D24', "=SUM(D21:D23)");
$objPHPExcel->getActiveSheet()->setCellValue('E24', "=SUM(E21:E23)");

$objPHPExcel->getActiveSheet()->setCellValue('D25', floatval($d25));
$objPHPExcel->getActiveSheet()->setCellValue('E25', floatval($e25));
$objPHPExcel->getActiveSheet()->setCellValue('D26', "=(D19+D24)");
$objPHPExcel->getActiveSheet()->setCellValue('E26', "=(E19+E24)");
//$objPHPExcel->getActiveSheet()->setCellValue('D27', floatval($d27));
//$objPHPExcel->getActiveSheet()->setCellValue('E27', floatval($e27));
$objPHPExcel->getActiveSheet()->setCellValue('D28', floatval($d28));
$objPHPExcel->getActiveSheet()->setCellValue('E28', floatval($e28));
$objPHPExcel->getActiveSheet()->setCellValue('D29', floatval($d29));
$objPHPExcel->getActiveSheet()->setCellValue('E29', floatval($e29));
$objPHPExcel->getActiveSheet()->setCellValue('D30', floatval($d30));
$objPHPExcel->getActiveSheet()->setCellValue('E30', floatval($e30));

$objPHPExcel->getActiveSheet()->setCellValue('D31', floatval($d31));
$objPHPExcel->getActiveSheet()->setCellValue('E31', floatval($e31));
$objPHPExcel->getActiveSheet()->setCellValue('D32', floatval($d32));
$objPHPExcel->getActiveSheet()->setCellValue('E32', floatval($e32));


//$objPHPExcel->getActiveSheet()->setCellValue('D26', "=SUM(D19:D23)");
//$objPHPExcel->getActiveSheet()->setCellValue('E26', "=SUM(E19:E23)");

$objPHPExcel->getActiveSheet()->setCellValue('D33', "=SUM(D28:D31)");
$objPHPExcel->getActiveSheet()->setCellValue('E33', "=SUM(E28:E31)");





//$objPHPExcel->getActiveSheet()->getStyle('B9:E10')->applyFromArray($styleArrayFont);
//$objPHPExcel->getActiveSheet()->getStyle('A13:K14')->applyFromArray($styleArrayAlignment);

$objPHPExcel->getActiveSheet()->getStyle('D11:E13')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('D15:E17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('D18:E20')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('D21:E27')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('D28:E33')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');



// TITLE 
$objPHPExcel->getActiveSheet()->setTitle('REPORT LTV');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_LTV_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Report_LTV_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>


<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> Laporan LTV
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
                                        Laporan LTV</a>
                                    </li>
                                  
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_LTV_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>

                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               <tr class="active">
                                                <td width="5%" align="center"><b>No </b></td>
                                                <td width="50%" align="center"><b>Jenis Kredit</b></td>
                                                <td width="25%" align="center"><b> Jumlah </b></td>
                                                <td width="20%" align="center"><b> NPL </b></td>
                                               
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  style="font-size:12px">1 </td>
                                                <td  style="font-size:12px">KPR ( Kredit Pemilikan Rumah ) </td>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px"> </td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Rumah Tinggal </td>
                                                <td  style="font-size:12px" align="right"> </td>
                                                <td  style="font-size:12px" align="right"> </td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe < 21 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe 22 - 70 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe > 70 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Rumah Susun</td>
                                                <td  style="font-size:12px" align="right"> </td>
                                                <td  style="font-size:12px" align="right"> </td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe < 21 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe 22 - 70 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Tipe > 70 </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Ruko / Rukan</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"> </td>
                                                <td  style="font-size:12px">Total KPR</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px">2</td>
                                                <td  style="font-size:12px">KKBP ( Kredit Konsumsi Beragun Properti )</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Tinggal</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Apartemen / Rumah Susun</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Toko / Rumah Kantor)</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Total KKBP</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D24')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E24')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px">3</td>
                                                <td  style="font-size:12px">Total KPR / KKBP dengan Nominal s.d Rp 5 Juta </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px">4</td>
                                                <td  style="font-size:12px">Total KP (Kredit Properti) </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px">5</td>
                                                <td  style="font-size:12px">KKB ( Kredit Kendaraan Bermotor ) </td>
                                                <td  style="font-size:12px" align="right"> <?php echo $objPHPExcel->getActiveSheet()->getCell('D27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"> <?php echo $objPHPExcel->getActiveSheet()->getCell('D27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Tangga untuk Pemilikan Mobil Roda Empat </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Tangga untuk Pemilikan Sepeda Bermotort </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Tangga untuk Pemilikan Truk dan Kendaraan Bermotor Roda Enam atau Lebih</td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px">Rumah Tangga untuk Pemilikan Kendaraan Bermotor Lainnya </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px">6</td>
                                                <td  style="font-size:12px">Total KKB dengan Nominal s.d Rp 5 Juta </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td  style="font-size:12px">7</td>
                                                <td  style="font-size:12px">Total KKB  </td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('D33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell('E33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
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

