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
logActivity("generate premi lps",date('Y_m_d_H_i_s'));



$tahun=$_POST['tahun'];
$semester=$_POST['semester'];







switch ($semester) {
        case '1':
        $name="Semester 1";
        $tanggal=$tahun."-06-30";
        $input=$tahun."-06-01";
        $array_bulan=array("Januari","Februari","Maret","April","Mei","Juni");
        $array_bulan_sebelumnya=array("Juli","Agustus","September","Oktober","November","Desember");
        $tahun_sebelumnya=date('Y',strtotime(date('Y-m-d',strtotime($input))." -1 year"));
        //$tanggal_akhir=$tahun."-06-30";
        $tanggal_start=$tahun."-01-01";

       
        break;
        case '2':
        $name="Semester 2";
        $tanggal=$tahun."-12-31";
        $input=$tahun."-12-01";
        $array_bulan=array("Juli","Agustus","September","Oktober","November","Desember");
        $array_bulan_sebelumnya=array("Januari","Februari","Maret","April","Mei","Juni");
        $tahun_sebelumnya=date('Y',strtotime(date('Y-m-d',strtotime($input))." 0 year"));
        $tanggal_start=$tahun."-07-31";
        //$tanggal_akhir=$tahun."-12-31";
        break;
        
     
}

$array_tgl=array();
$start=5;
for ($i=1; $i <=6 ; $i++) { 
$tgl=date('Y-m-t',strtotime(date('Y-m-d',strtotime($input))." -$start month"));
array_push($array_tgl,$tgl);
$start=$start-1;
}



// tanggal sebelumnya

$array_tgl_sebelumnya=array();
$start_sebelumnya=11;
for ($i=1; $i <=6 ; $i++) { 
$tgl_sebelumnya=date('Y-m-t',strtotime(date('Y-m-d',strtotime($input))." -$start_sebelumnya month"));
array_push($array_tgl_sebelumnya,$tgl_sebelumnya);
$start_sebelumnya=$start_sebelumnya-1;
}



//var_dump($array_tgl);
//die();

//$tanggal=$tahun.$semester;




$var_tgl=date('Y-m-d',strtotime($tanggal));


#############################  QUERY  MASTER PREMI ################################

$mon1=date('n',strtotime($tanggal_start));
$year1=date('Y',strtotime($tanggal_start));
$mon2=date('n',strtotime($tanggal));
$year2=date('Y',strtotime($tanggal));

$query_premi =" select * from Master_Saldo_Premi_LPS   ";
$query_premi.=" where Month(Periode_Awal)='$mon1' and Year(Periode_Awal)='$year1' and Month(Periode_Akhir)='$mon2' and Year(Periode_Akhir)='$year2' ";



        $result_premi=odbc_exec($connection2, $query_premi);
        $row=odbc_fetch_array($result_premi);
        $ver_premi=floatval($row['Jumlah_Premi_Verifikasi_LPS']);
        $saldo_premi_sblmnya=floatval($row['Saldo_Premi_Periode_Lalu']);








$sim_pihak3=array();
#################################################   QUERY   #############################################################

foreach ($array_tgl as $key => $value_tgl) {
/*
$query=" SELECT Tanggal,SUM (Nilai_DPK)as Total_DPK FROM (
SELECT SUM (Sub_Total_DPK)AS Nilai_DPK,Tanggal FROM(
SELECT SUM(Nominal_Rupiah)as Sub_Total_DPK,Tanggal FROM(
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM DM_LiabilitasGiro a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
join Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS201000000' 
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM DM_LiabilitasTabungan a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
join Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS202000000' 
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM DM_LiabilitasDeposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
join Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS203000000' 
AND a.status_oncall='Y'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM DM_LiabilitasDeposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
join Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS203000000' 
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Rupiah
GROUP BY Tanggal

UNION ALL
SELECT SUM(Nominal_Valas)as Sub_Total_DPK,Tanggal FROM(
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal  
FROM DM_LiabilitasGiro a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS201000000'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal  
FROM DM_LiabilitasTabungan a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS202000000' 
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal 
FROM DM_LiabilitasDeposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS203000000' 
AND a.status_oncall='Y'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal   
FROM DM_LiabilitasDeposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS203000000' 
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Valas 
GROUP BY Tanggal
)AS DPK
GROUP BY Tanggal
)AS Total_DPK
GROUP BY Tanggal ";
*/
$var_tabel=date('Ymd',strtotime($value_tgl));

############################################### CEK TABEL##############################################
$table_giro="DM_LiabilitasGiro_$var_tabel";
$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
//$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";


$qcekTable=" select top 1* from  $table_giro ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_giro Tidak tersedia ... ! </b> <br><br></div>";
       // die();
}

$qcekTable=" select top 1* from  $table_tabungan ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_tabungan Tidak tersedia ... ! </b> <br><br></div>";
       // die();
}

$qcekTable=" select top 1* from  $table_deposito ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_deposito Tidak tersedia ... ! </b> <br><br></div>";
        //die();
}



################################################  END CEK TABLE ####################################

/*

$query="  SELECT Tanggal,SUM (Nilai_DPK)as Total_DPK FROM (
SELECT FLAG,SUM (Sub_Total_DPK)AS Nilai_DPK,Tanggal FROM(
SELECT FLAG,SUM(Nominal_Rupiah)as Sub_Total_DPK,Tanggal FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM $table_giro a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal
FROM $table_tabungan a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' 
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal 
FROM $table_deposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' 
AND a.status_oncall='y'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal  
FROM $table_deposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' 
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Rupiah
GROUP BY FLAG,Tanggal
UNION ALL
SELECT FLAG,SUM(Nominal_Valas)as Sub_Total_DPK,Tanggal FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal  
FROM $table_giro a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal  
FROM $table_tabungan a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' 
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal 
FROM $table_deposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' 
AND a.status_oncall='y'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal   
FROM $table_deposito a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' 
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Valas 
GROUP BY FLAG,Tanggal
)AS DPK
GROUP BY FLAG,Tanggal
)AS Total_DPK
GROUP BY Tanggal ";

*/

$query=" SELECT SUM (Nilai_DPK)as Total_DPK FROM (
SELECT FLAG,SUM (Sub_Total_DPK)AS Nilai_DPK FROM(
SELECT FLAG,SUM( JumlahBulanLaporan_Rupiah)as Sub_Total_DPK FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS101000000'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS103000000'
and a.status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR'and b.PREMILPS_Level_2='PREMILPS103000000' 
and a.status_oncall='n'
)AS Rupiah
GROUP BY FLAG
UNION ALL
SELECT FLAG,SUM(JumlahBulanLaporan_Valas)as Sub_Total_DPK FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS101000000' 
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'and a.status_oncall='Y' and b.PREMILPS_Level_2='PREMILPS103000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'and a.status_oncall='N' and b.PREMILPS_Level_2='PREMILPS103000000' and a.status_oncall='n'
)AS Valas 
GROUP BY FLAG
)AS DPK
GROUP BY FLAG
)AS Total_DPK ";





//echo $query;
//die();

        $result1=odbc_exec($connection2, $query);
        $row1=odbc_fetch_array($result1);
        array_push($sim_pihak3,$row1['Total_DPK']);


}


$sim_bank_lain=array();
#################################  QUERY DARI BANK LAIN #######################################
foreach ($array_tgl as $key => $value_tgl) {


$var_tabel=date('Ymd',strtotime($value_tgl));


#############################################################################################

$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";


$qcekTable=" select top 1* from  $table_banklain ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_banklain Tidak tersedia ... ! </b> <br><br></div>";
       // die();
}




/*
$query2=" SELECT Tanggal,SUM(Nilai_Bank_Lain)as Total_Bank_Lain FROM (
SELECT SUM (Sub_Total_Bank_Lain)AS Nilai_Bank_Lain,Tanggal FROM(
SELECT SUM(Nominal_Rupiah)as Sub_Total_Bank_Lain,Tanggal FROM(
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal  
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS101000000'
AND a.flagDPK='Giro'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal  
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR' AND d.PREMILPS_Level_2 ='PREMILPS102000000' 
AND a.flagDPK='Tabungan'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Rupiah,a.DataDate as Tanggal   
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang='IDR'  AND d.PREMILPS_Level_2 ='PREMILPS103000000' 
AND a.flagDPK='Deposito'
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Rupiah
GROUP BY Tanggal
UNION ALL
SELECT SUM(Nominal_Valas)as Sub_Total_Bank_Lain,Tanggal FROM(
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal 
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS101000000'
AND a.flagDPK='Giro'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal  
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS102000000'  
AND a.flagDPK='Tabungan'
GROUP BY a.DataDate,b.Kurs_Tengah
UNION
SELECT SUM(a.jumlahbulanlaporanorgamount)* b.Kurs_Tengah AS Nominal_Valas,a.DataDate as Tanggal   
FROM $table_banklain a WITH (NOLOCK)
JOIN DM_Kurs b ON a.DataDate = b.DataDate AND b.Jenis_Mata_Uang = a.JenisMataUang
JOIN Referensi_GL_02 c ON c.GLNO = a.managed_gl_code AND c.PRODNO = a.managed_gl_prod_code
JOIN Referensi_PremiLPS d ON d.PREMILPS_Level_2 = c.PREMILPS_Level_2
WHERE a.DataDate ='$value_tgl' AND a.JenisMataUang<>'IDR' AND d.PREMILPS_Level_2 ='PREMILPS103000000' 
AND a.flagDPK='Deposito'
GROUP BY a.DataDate,b.Kurs_Tengah
)AS Valas
GROUP BY Tanggal
)AS Bank_Lain
GROUP BY Tanggal
)AS Total_Bank_Lain
GROUP BY Tanggal ";
*/
$query2=" SELECT SUM(Nilai_Bank_Lain)as Total_Bank_Lain FROM (
SELECT FLAG,SUM (Sub_Total_Bank_Lain)AS Nilai_Bank_Lain FROM(
SELECT FLAG,SUM(JumlahBulanLaporan_Rupiah)as Sub_Total_Bank_Lain FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' 
AND a.FlagDPK='Deposito' and a.status_oncall='y' and b.PREMILPS_Level_2='PREMILPS203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS203000000' and a.status_oncall='n' 
AND a.FlagDPK='Deposito'
)AS Rupiah
GROUP BY FLAG
UNION ALL
SELECT FLAG,SUM(JumlahBulanLaporan_Valas)as Sub_Total_Bank_Lain FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='Y' and b.PREMILPS_Level_2='PREMILPS203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='N' and b.PREMILPS_Level_2='PREMILPS203000000' and a.status_oncall='n'
)AS Valas
GROUP BY FLAG
)AS Bank_Lain
GROUP BY FLAG
)AS Total_Bank_Lain ";







//echo $query2;
//die();
        $result2=odbc_exec($connection2, $query2);
        $row2=odbc_fetch_array($result2);
        array_push($sim_bank_lain,$row2['Total_Bank_Lain']);


}


#################  PERIODE SEBELUMNYA ###################################################
$sim_pihak3_sebelumnya=array();
#################################################   QUERY   #############################################################

foreach ($array_tgl_sebelumnya as $key => $value_tgl) {

$var_tabel=date('Ymd',strtotime($value_tgl));

############################################### CEK TABEL##############################################
$table_giro="DM_LiabilitasGiro_$var_tabel";
$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
//$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

$qcekTable=" select top 1* from  $table_giro ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_giro Tidak tersedia ... ! </b> <br><br></div>";
        //die();
}

$qcekTable=" select top 1* from  $table_tabungan ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_tabungan Tidak tersedia ... ! </b> <br><br></div>";
        //die();
}

$qcekTable=" select top 1* from  $table_deposito ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_deposito Tidak tersedia ... ! </b> <br><br></div>";
        //die();
}

$query=" SELECT SUM (Nilai_DPK)as Total_DPK FROM (
SELECT FLAG,SUM (Sub_Total_DPK)AS Nilai_DPK FROM(
SELECT FLAG,SUM( JumlahBulanLaporan_Rupiah)as Sub_Total_DPK FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS101000000'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS103000000'
and a.status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR'and b.PREMILPS_Level_2='PREMILPS103000000' 
and a.status_oncall='n'
)AS Rupiah
GROUP BY FLAG
UNION ALL
SELECT FLAG,SUM(JumlahBulanLaporan_Valas)as Sub_Total_DPK FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS101000000' 
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'and a.status_oncall='Y' and b.PREMILPS_Level_2='PREMILPS103000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'and a.status_oncall='N' and b.PREMILPS_Level_2='PREMILPS103000000' and a.status_oncall='n'
)AS Valas 
GROUP BY FLAG
)AS DPK
GROUP BY FLAG
)AS Total_DPK ";


        $result1=odbc_exec($connection2, $query);
        $row1=odbc_fetch_array($result1);
        array_push($sim_pihak3_sebelumnya,$row1['Total_DPK']);

}


$sim_bank_lain_sebelumnya=array();
#################################  QUERY DARI BANK LAIN #######################################
foreach ($array_tgl_sebelumnya as $key => $value_tgl) {


$var_tabel=date('Ymd',strtotime($value_tgl));


#############################################################################################

$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";


$qcekTable=" select top 1* from  $table_banklain ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_banklain Tidak tersedia ... ! </b> <br><br></div>";
        //die();
}


$query2=" SELECT SUM(Nilai_Bank_Lain)as Total_Bank_Lain FROM (
SELECT FLAG,SUM (Sub_Total_Bank_Lain)AS Nilai_Bank_Lain FROM(
SELECT FLAG,SUM(JumlahBulanLaporan_Rupiah)as Sub_Total_Bank_Lain FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' 
AND a.FlagDPK='Deposito' and a.status_oncall='y' and b.PREMILPS_Level_2='PREMILPS203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang='IDR' and b.PREMILPS_Level_2='PREMILPS203000000' and a.status_oncall='n' 
AND a.FlagDPK='Deposito'
)AS Rupiah
GROUP BY FLAG
UNION ALL
SELECT FLAG,SUM(JumlahBulanLaporan_Valas)as Sub_Total_Bank_Lain FROM(
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR' and b.PREMILPS_Level_2='PREMILPS202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='Y' and b.PREMILPS_Level_2='PREMILPS203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_PremiLPS c ON c.PREMILPS_Level_2 = b.PREMILPS_Level_2
WHERE a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='N' and b.PREMILPS_Level_2='PREMILPS203000000' and a.status_oncall='n'
)AS Valas
GROUP BY FLAG
)AS Bank_Lain
GROUP BY FLAG
)AS Total_Bank_Lain ";
//echo $query2;
//die();
        $result2=odbc_exec($connection2, $query2);
        $row2=odbc_fetch_array($result2);
        array_push($sim_bank_lain_sebelumnya,$row2['Total_Bank_Lain']);

}





#########################################################################################




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

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder3 = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);



$objPHPExcel->getActiveSheet()->getStyle('A3:I14')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->getStyle('F21')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F23')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F26')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F28')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F30')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F33')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F36')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('E16')->applyFromArray($styleArrayBorder3);
//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A15:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:Z2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('J1:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

//$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//$objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A2:I6')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A13:I14')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A3:I6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A3:I6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A13:H14')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A13:H14')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A19:A36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A19:A36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('E16:E17')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B19:B36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('F19:F36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(25);

$objPHPExcel->getActiveSheet()->getRowDimension(21)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(23)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(26)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(28)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(30)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(33)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(36)->setRowHeight(40);

$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C3:E3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F3:H3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I3:I5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B3:B6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:A6');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C4:C5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D4:D5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E4:E5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F4:F5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('G4:G5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H4:H5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A13:H14');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F16:F17');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('I13:I14');

$objPHPExcel->getActiveSheet()->setCellValue('A2', "POSISI SIMPANAN BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('C2', "$array_bulan[0] S/d $array_bulan[5] $tahun ");

$objPHPExcel->getActiveSheet()->setCellValue('A3', "No.");
$objPHPExcel->getActiveSheet()->setCellValue('B3', "Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('C3', "Simpanan Konvensional");
$objPHPExcel->getActiveSheet()->setCellValue('F3', "Simpanan Syariah/UUS");
$objPHPExcel->getActiveSheet()->setCellValue('I3', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C4', "Simpanan Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue('D4', "Simpanan Dari Bank Lain ");
$objPHPExcel->getActiveSheet()->setCellValue('E4', "Sub Total 1 ");
$objPHPExcel->getActiveSheet()->setCellValue('F4', "Simpanan Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue('G4', "Simpanan Dari Bank Lain ");
$objPHPExcel->getActiveSheet()->setCellValue('H4', "Sub Total 1 ");

$objPHPExcel->getActiveSheet()->setCellValue('C6', "(1)");
$objPHPExcel->getActiveSheet()->setCellValue('D6', "(2)");
$objPHPExcel->getActiveSheet()->setCellValue('E6', "(3)");
$objPHPExcel->getActiveSheet()->setCellValue('F6', "(4)");
$objPHPExcel->getActiveSheet()->setCellValue('G6', "(5)");
$objPHPExcel->getActiveSheet()->setCellValue('H6', "(6)");
$objPHPExcel->getActiveSheet()->setCellValue('I6', "(7)");

#######  SIMP PIHAK KE3
$counter2=7;
$index2=0;

foreach ($sim_pihak3 as $key => $valrow) {
 
$objPHPExcel->getActiveSheet()->setCellValue("C$counter2", $sim_pihak3["$index2"]);



$counter2++;
$index2++;
}
#######  DARI BANK LAIN 
$counter3=7;
$index3=0;

foreach ($sim_bank_lain as $key => $valrow) {
 
$objPHPExcel->getActiveSheet()->setCellValue("D$counter3", $sim_bank_lain["$index3"]);



$counter3++;
$index3++;
}


$counter=7;
$index=0;
$no=1;

 foreach ($array_bulan as $value) {
  $objPHPExcel->getActiveSheet()->setCellValue("A$counter", $no);
  $objPHPExcel->getActiveSheet()->setCellValue("B$counter", $array_bulan["$index"]);

  $counter++;
  $index++;
  $no++;
 }



$objPHPExcel->getActiveSheet()->setCellValue('A13', "Total Simpanan");

$objPHPExcel->getActiveSheet()->setCellValue('B16', "Dasar Perhitungan Premi");
$objPHPExcel->getActiveSheet()->setCellValue('C16', "Total Simpanan (A) ");
$objPHPExcel->getActiveSheet()->setCellValue('C17', "Jumlah Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('D16', "=");
$objPHPExcel->getActiveSheet()->setCellValue('E16', "=I13");
$objPHPExcel->getActiveSheet()->setCellValue('E17', 6);
$objPHPExcel->getActiveSheet()->setCellValue('F16', "");

$objPHPExcel->getActiveSheet()->setCellValue('A19', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "PENYESUAIAN PREMI PERIODE ");
$objPHPExcel->getActiveSheet()->setCellValue('C19', "$array_bulan[0] S/d $array_bulan[5] $tahun");
$objPHPExcel->getActiveSheet()->setCellValue('A21', "1.a");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "REALISASI PREMI (0,1% x Dasar Perhitungan Premi)");
$objPHPExcel->getActiveSheet()->setCellValue('A23', "1.b");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "PREMI AWAL (0,1% x Dasar Perhitungan Premi Periode Sebelumnya) $array_bulan_sebelumnya[0] S/d $array_bulan_sebelumnya[5] $tahun_sebelumnya ");

$objPHPExcel->getActiveSheet()->setCellValue('A26', "1.c");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "PENYESUAIAN PREMI = C - D");
$objPHPExcel->getActiveSheet()->setCellValue('A28', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "PREMI AWAL PERIODE $array_bulan[0] S/d $array_bulan[5] $tahun ");
$objPHPExcel->getActiveSheet()->setCellValue('A30', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "JUMLAH PREMI HASIL VERIFIKASI: PERIODE $array_bulan[0] s/d $array_bulan[5] ");
$objPHPExcel->getActiveSheet()->setCellValue('A33', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "JUMLAH PREMI PERIODE $array_bulan[0] s/d $array_bulan[5] Tahun $tahun = (E + F + G) ");
$objPHPExcel->getActiveSheet()->setCellValue('A36', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "SALDO PREMI PERIODE SEBELUMNYA");
$objPHPExcel->getActiveSheet()->setCellValue('A39', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "JUMLAH PREMI YANG HARUS DIBAYARKAN = H+I");



for ($i=7; $i <= 12 ; $i++) { 
  # code...
  $objPHPExcel->getActiveSheet()->setCellValue("E$i", "=(C$i+D$i)");
  $objPHPExcel->getActiveSheet()->setCellValue("H$i", "=(F$i+G$i)");
  $objPHPExcel->getActiveSheet()->setCellValue("I$i", "=(E$i+H$i)");
}

$objPHPExcel->getActiveSheet()->setCellValue("I13", "=SUM(I7:I12)");

// BAGIAN BAWAH
$objPHPExcel->getActiveSheet()->setCellValue("F16", "=E16/E17");
$objPHPExcel->getActiveSheet()->setCellValue("F21", "=F16*0.1%");
$objPHPExcel->getActiveSheet()->setCellValue("F23", "=PERIODE_SEBELUMNYA!I13");
$objPHPExcel->getActiveSheet()->setCellValue("F26", "=F21-F23");
$objPHPExcel->getActiveSheet()->setCellValue("F28", "=F21");
$objPHPExcel->getActiveSheet()->setCellValue("F30", $ver_premi);
$objPHPExcel->getActiveSheet()->setCellValue("F33", "=F26+F28-F30");
$objPHPExcel->getActiveSheet()->setCellValue("F36", $saldo_premi_sblmnya);
$objPHPExcel->getActiveSheet()->setCellValue("F39", "=+F33-F36");



$objPHPExcel->getActiveSheet()->getStyle('C7:I12')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('I13')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E16')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F16')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F21')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F26')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F30')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F33')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F36')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->setTitle('LAPORAN PREMI LPS');

#####  SHEET 2 BAYANGAN U PERITUNGAN  BULAN SEBELUMNYA
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
$styleArrayBorder3 = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);



$objPHPExcel->getActiveSheet()->getStyle('A3:I14')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->getStyle('F21')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F23')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F26')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F28')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F30')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F33')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F36')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('E16')->applyFromArray($styleArrayBorder3);
//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A15:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:Z2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('J1:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

//$objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//$objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
$objPHPExcel->getActiveSheet()->getStyle('A2:I6')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A13:I14')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A3:I6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A3:I6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A13:H14')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A13:H14')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A19:A36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A19:A36')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('E16:E17')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B19:B36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('F19:F36')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('F16:F17')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(25);

$objPHPExcel->getActiveSheet()->getRowDimension(21)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(23)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(26)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(28)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(30)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(33)->setRowHeight(40);
$objPHPExcel->getActiveSheet()->getRowDimension(36)->setRowHeight(40);

$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C3:E3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('F3:H3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I3:I5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B3:B6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:A6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C4:C5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('D4:D5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('E4:E5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('F4:F5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('G4:G5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('H4:H5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A13:H14');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('F16:F17');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I13:I14');

$objPHPExcel->getActiveSheet()->setCellValue('A2', "POSISI SIMPANAN BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('C2', "$array_bulan[0] S/d $array_bulan[5] $tahun ");

$objPHPExcel->getActiveSheet()->setCellValue('A3', "No.");
$objPHPExcel->getActiveSheet()->setCellValue('B3', "Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('C3', "Simpanan Konvensional");
$objPHPExcel->getActiveSheet()->setCellValue('F3', "Simpanan Syariah/UUS");
$objPHPExcel->getActiveSheet()->setCellValue('I3', "Total");

$objPHPExcel->getActiveSheet()->setCellValue('C4', "Simpanan Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue('D4', "Simpanan Dari Bank Lain ");
$objPHPExcel->getActiveSheet()->setCellValue('E4', "Sub Total 1 ");
$objPHPExcel->getActiveSheet()->setCellValue('F4', "Simpanan Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue('G4', "Simpanan Dari Bank Lain ");
$objPHPExcel->getActiveSheet()->setCellValue('H4', "Sub Total 1 ");

$objPHPExcel->getActiveSheet()->setCellValue('C6', "(1)");
$objPHPExcel->getActiveSheet()->setCellValue('D6', "(2)");
$objPHPExcel->getActiveSheet()->setCellValue('E6', "(3)");
$objPHPExcel->getActiveSheet()->setCellValue('F6', "(4)");
$objPHPExcel->getActiveSheet()->setCellValue('G6', "(5)");
$objPHPExcel->getActiveSheet()->setCellValue('H6', "(6)");
$objPHPExcel->getActiveSheet()->setCellValue('I6', "(7)");

#######  SIMP PIHAK KE3
$counter2=7;
$index2=0;

foreach ($sim_pihak3_sebelumnya as $key => $valrow) {
 
$objPHPExcel->getActiveSheet()->setCellValue("C$counter2", $sim_pihak3_sebelumnya["$index2"]);



$counter2++;
$index2++;
}
#######  DARI BANK LAIN 
$counter3=7;
$index3=0;

foreach ($sim_bank_lain_sebelumnya as $key => $valrow) {
 
$objPHPExcel->getActiveSheet()->setCellValue("D$counter3", $sim_bank_lain_sebelumnya["$index3"]);



$counter3++;
$index3++;
}


$counter=7;
$index=0;
$no=1;

 foreach ($array_bulan as $value) {
  $objPHPExcel->getActiveSheet()->setCellValue("A$counter", $no);
  $objPHPExcel->getActiveSheet()->setCellValue("B$counter", $array_bulan_sebelumnya["$index"]);

  $counter++;
  $index++;
  $no++;
 }



$objPHPExcel->getActiveSheet()->setCellValue('A13', "Total Simpanan");

$objPHPExcel->getActiveSheet()->setCellValue('B16', "Dasar Perhitungan Premi");
$objPHPExcel->getActiveSheet()->setCellValue('C16', "Total Simpanan (A) ");
$objPHPExcel->getActiveSheet()->setCellValue('C17', "Jumlah Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('D16', "=");
$objPHPExcel->getActiveSheet()->setCellValue('E16', "=I13");
$objPHPExcel->getActiveSheet()->setCellValue('E17', 6);
$objPHPExcel->getActiveSheet()->setCellValue('F16', "");

$objPHPExcel->getActiveSheet()->setCellValue('A19', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "PENYESUAIAN PREMI PERIODE ");
$objPHPExcel->getActiveSheet()->setCellValue('C19', "$array_bulan[0] S/d $array_bulan[5] $tahun");
$objPHPExcel->getActiveSheet()->setCellValue('A21', "1.a");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "REALISASI PREMI (0,1% x Dasar Perhitungan Premi)");
$objPHPExcel->getActiveSheet()->setCellValue('A23', "1.b");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "PREMI AWAL (0,1% x Dasar Perhitungan Premi Periode Sebelumnya) $array_bulan_sebelumnya[0] s/d $array_bulan_sebelumnya[5] $tahun");

$objPHPExcel->getActiveSheet()->setCellValue('A26', "1.c");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "PENYESUAIAN PREMI = C - D");
$objPHPExcel->getActiveSheet()->setCellValue('A28', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "PREMI AWAL PERIODE $array_bulan[0] S/d $array_bulan[5] $tahun ");
$objPHPExcel->getActiveSheet()->setCellValue('A30', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "JUMLAH PREMI HASIL VERIFIKASI: PERIODE JULI s/d Desember ");
$objPHPExcel->getActiveSheet()->setCellValue('A33', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "JUMLAH PREMI PERIODE JANUARI s/d JUNI Tahun 2016 = (E + F + G) ");
$objPHPExcel->getActiveSheet()->setCellValue('A36', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "SALDO PREMI PERIODE SEBELUMNYA");
$objPHPExcel->getActiveSheet()->setCellValue('A39', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "JUMLAH PREMI YANG HARUS DIBAYARKAN = H+I");



for ($i=7; $i <= 12 ; $i++) { 
  # code...
  $objPHPExcel->getActiveSheet()->setCellValue("E$i", "=(C$i+D$i)");
  $objPHPExcel->getActiveSheet()->setCellValue("H$i", "=(F$i+G$i)");
  $objPHPExcel->getActiveSheet()->setCellValue("I$i", "=(E$i+H$i)");
}

$objPHPExcel->getActiveSheet()->setCellValue("I13", "=SUM(I7:I12)");

// BAGIAN BAWAH
$objPHPExcel->getActiveSheet()->setCellValue("F16", "=E16/E17");
$objPHPExcel->getActiveSheet()->setCellValue("F21", "=F16*0.1%");
$objPHPExcel->getActiveSheet()->setCellValue("F23", "");
$objPHPExcel->getActiveSheet()->setCellValue("F26", "=F21-F23");
$objPHPExcel->getActiveSheet()->setCellValue("F28", "=F21");
$objPHPExcel->getActiveSheet()->setCellValue("F30", "");
$objPHPExcel->getActiveSheet()->setCellValue("F33", "=F26+F28-F30");
$objPHPExcel->getActiveSheet()->setCellValue("F36", "");
$objPHPExcel->getActiveSheet()->setCellValue("F39", "=+F33-F36");



$objPHPExcel->getActiveSheet()->getStyle('C7:I12')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('I13')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E16')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F16')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F21')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F26')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F30')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F33')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F36')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setTitle('PERIODE_SEBELUMNYA');

$objPHPExcel->getSheetByName('PERIODE_SEBELUMNYA')->setSheetState(PHPExcel_Worksheet::SHEETSTATE_HIDDEN);

$objPHPExcel->setActiveSheetIndex(0);


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_PREMI_lps_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Report_PREMI_lps_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);


?>



<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> LAPORAN PREMI LPS
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
                                        LAPORAN PREMI LPS </a>
                                    </li>
                                   
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_PREMI_lps_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>
    <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> POSISI SIMPANAN BULAN :    <?php echo $array_bulan["0"]." s/d ".$array_bulan["5"]." $tahun"; ?>
</b>
                                    </div>                                  
                                        
                                        <p>
                                        
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="2"><b>No</b></td>
                                                <td width="10%" align="center" rowspan="2"><b>Bulan </b></td>
                                                <td width="36%" align="center" colspan="3"><b>Simpanan Konvensional </b></td>
                                                <td width="36%" align="center" colspan="3"><b>Simpanan Syariah / UUS </b></td>
                                                <td width="10%" align="center" rowspan="2"><b>Total </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="15%" align="center"><b>Simpanan Pihak Ketiga</b></td>
                                                <td width="15%" align="center" ><b>Simpanan dari bank Lain </b></td>
                                                <td width="15%" align="center" ><b>Sub Total 1 </b></td>
                                                <td width="15%" align="center"><b>Simpanan Pihak Ketiga</b></td>
                                                <td width="15%" align="center" ><b>Simpanan dari bank Lain </b></td>
                                                <td width="15%" align="center" ><b>Sub Total 1 </b></td>
                                                
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                

                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(0);
                                                $number_dash1=13;
                                                for ($i=7; $i <= 12 ; $i++) { 

                                                    ?>
                                                <tr>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

                                                <?php       
                                                    }
                                                ?>
                                                <tr class="danger">
                                                <td width="90%" align="center" colspan="8"><b>Total Simpanan</b></td>
                                                <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>
                                 
                                      <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <tbody>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%"></td>
                                                <td  style="font-size:12px" align="left" width="65"><b>Dasar Perhitungan Premi</b></td>
                                                <td  style="font-size:12px" align="center" width="10">Total Simpanan (A) / Jumlah Bulan </td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F16")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">1</td>
                                                <td  style="font-size:12px" align="left" width="65">PENYESUAIAN PREMI PERIODE </td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">1.a</td>
                                                <td  style="font-size:12px" align="left" width="65">REALISASI PREMI (0,1% x Dasar Perhitungan Premi)</td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F21")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">1.b</td>
                                                <td  style="font-size:12px" align="left" width="65">PREMI AWAL (0,1% x Dasar Perhitungan Premi Periode Sebelumnya) <?php echo $array_bulan_sebelumnya["0"]." s/d ".$array_bulan_sebelumnya["5"]." $tahun_sebelumnya"?> </td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F23")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">1.c</td>
                                                <td  style="font-size:12px" align="left" width="65">PENYESUAIAN PREMI = C - D</td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F26")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">2</td>
                                                <td  style="font-size:12px" align="left" width="65">PREMI AWAL PERIODE <?php echo $array_bulan["0"]." s/d ".$array_bulan["5"]." $tahun"?></td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F28")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">3</td>
                                                <td  style="font-size:12px" align="left" width="65">JUMLAH PREMI HASIL VERIFIKASI: PERIODE <?php echo $array_bulan["0"]." s/d ".$array_bulan["5"]." $tahun"?> </td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F30")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">4</td>
                                                <td  style="font-size:12px" align="left" width="65">JUMLAH PREMI PERIODE <?php echo $array_bulan["0"]." s/d ".$array_bulan["5"]." $tahun"?> = (E + F + G)</td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F33")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">5</td>
                                                <td  style="font-size:12px" align="left" width="65">SALDO PREMI PERIODE SEBELUMNYA</td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F36")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td  style="font-size:12px" align="right" width="5%">6</td>
                                                <td  style="font-size:12px" align="left" width="65">JUMLAH PREMI YANG HARUS DIBAYARKAN = H+I</td>
                                                <td  style="font-size:12px" align="center" width="10"></td>
                                                <td  style="font-size:12px" align="right" width="20"><?php echo $objPHPExcel->getActiveSheet()->getCell("F39")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                       
                                </div>
                            </div>
                            
                        </div>
                </div>

