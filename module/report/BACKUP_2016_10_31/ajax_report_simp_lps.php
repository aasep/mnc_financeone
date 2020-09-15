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
logActivity("generate simp lps",date('Y_m_d_H_i_s'));

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


#############################################################################################
$table_giro="DM_LiabilitasGiro_$var_tabel";
$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

$qcekTable=" select top 1* from  $table_giro ";
$res_cek=odbc_exec($connection2, $qcekTable);
if (!$res_cek){
        echo "<div class='lert alert-danger' align='center'> <br><b> Tabel $table_giro Tidak tersedia ... ! </b> <br><br></div>";
        die();
}

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









#--Query Lampiran 1B--
#--Simpanan Pihak Ketiga Rupiah--


$query_kurs=" select Kurs_Tengah from DM_Kurs where Jenis_Mata_Uang='USD' and DataDate='$curr_tgl' ";
        $result=odbc_exec($connection2, $query_kurs);
        $row_kurs=odbc_fetch_array($result);

        $nilai_kurs=floatval($row_kurs['Kurs_Tengah']);
/*
$range1=" and Nominal between '0' and '100000000' ";
$range2=" and Nominal >'100000000' and Nominal<'200000001' ";
$range3=" and Nominal >'200000001' and Nominal<'500000001' ";
$range4=" and Nominal >'500000001' and Nominal<'1000000001' ";
$range5=" and Nominal >'1000000000' and Nominal<'2000000001' ";
$range6=" and Nominal >'2000000001' and Nominal<'5000000001' ";
$range7=" and Nominal > '5000000000' ";
*/

/*

--Simpanan Pihak Ketiga Rupiah--
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM dbo.DM_LiabilitasGiro WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM dbo.DM_LiabilitasTabungan WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM dbo.DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 
and status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM dbo.DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 


--Simpanan Pihak Ketiga Valuta Asing --
SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM dbo.DM_LiabilitasGiro WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM dbo.DM_LiabilitasTabungan WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM dbo.DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas 
FROM dbo.DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas



--Simpanan Bank Lain Rupiah--
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR'
AND FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS JumlahBulanLaporan_Rupiah
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 
AND FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 
AND FlagDPK='Deposito' and status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Rupiah
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='IDR' 
AND FlagDPK='Deposito'

--Simpanan Bank Lain Valuta Asing --
SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
AND FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
AND FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
AND FlagDPK='Deposito' and status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS JumlahBulanLaporan_Valas 
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='2016-07-31' AND JenisMataUang='USD'
AND FlagDPK='Deposito'
)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas


*/





$range1=" and JumlahBulanLaporan >='0' and JumlahBulanLaporan <'100000001' ";
$range2=" and JumlahBulanLaporan >='100000001' and JumlahBulanLaporan <'200000001' ";
$range3=" and JumlahBulanLaporan >='200000001' and JumlahBulanLaporan <'500000001' ";
$range4=" and JumlahBulanLaporan >='500000001' and JumlahBulanLaporan <'1000000001' ";
$range5=" and JumlahBulanLaporan >='1000000001' and JumlahBulanLaporan <'2000000001' ";
$range6=" and JumlahBulanLaporan >='2000000001' and JumlahBulanLaporan <'5000000001' ";
$range7=" and JumlahBulanLaporan > '5000000000' ";


$array_range1=array("$range1","$range2","$range3","$range4","$range5","$range6","$range7");
/*
$query =" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_giro WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_tabungan WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' 
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' 
AND status_oncall ='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' and status_oncall='n'  ";
*/
$query =" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM101000000'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM103000000'
and a.status_oncall='y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR'and b.LAPOSIM_Level_2='LAPOSIM103000000' ";



//echo $query."<br>";
//die();
        $result2=odbc_exec($connection2, $query);
//echo $query;
//die();
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
/*        
$query =" SELECT FLAG,Nominal_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasGiro WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasTabungan WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
AND status_oncall ='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas 
FROM DM_LiabilitasDeposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'


)AS Tabel1
GROUP BY FLAG,Nominal_Valas ";

*/
/*
$query ="SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_giro WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_tabungan WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_deposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'and status_oncall='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_deposito WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'and status_oncall='N'
)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas ";
*/


$query=" SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_giro a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR' and b.LAPOSIM_Level_2='LAPOSIM101000000' 
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_tabungan a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR' and b.LAPOSIM_Level_2='LAPOSIM102000000'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR'and a.status_oncall='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_deposito a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR'and a.status_oncall='N' and b.LAPOSIM_Level_2='LAPOSIM103000000'
)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas ";




//echo $query."<br>";


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
/*
$query=" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR'
AND FlagDPK ='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' 
AND FlagDPK ='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' 
AND FlagDPK ='Deposito' AND status_oncall='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='IDR' 
AND FlagDPK ='Deposito' and status_oncall='n' ";
*/

$query=" SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan) AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' 
AND a.FlagDPK='Deposito' and a.status_oncall='y' and b.LAPOSIM_Level_2='LAPOSIM203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS Nominal_Rupiah 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and b.LAPOSIM_Level_2='LAPOSIM203000000'
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
        /*
$query=" SELECT FLAG,Nominal_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
AND FlagDPK ='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
AND FlagDPK ='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
AND FlagDPK ='Deposito' AND status_oncall='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporanorgamount)AS Nominal_Valas 
FROM DM_LiabilitasKepadaBankLain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang='USD'
AND FlagDPK ='Deposito' 
)AS Tabel1
GROUP BY FLAG,Nominal_Valas ";

*/
/*
$query=" SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
AND FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
AND FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
AND FlagDPK='Deposito' and status_oncall='Y'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_banklain WITH (NOLOCK)
WHERE DataDate ='$curr_tgl' AND JenisMataUang<>'IDR'
AND FlagDPK='Deposito' and status_oncall='N'
)AS Tabel1
GROUP BY FLAG,JumlahBulanLaporan_Valas ";
*/


$query=" SELECT FLAG,JumlahBulanLaporan_Valas from (
SELECT DISTINCT 'GIRO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR' and b.LAPOSIM_Level_2='LAPOSIM201000000'
AND a.FlagDPK='Giro'
UNION
SELECT DISTINCT 'TABUNGAN' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR' and b.LAPOSIM_Level_2='LAPOSIM202000000'
AND a.FlagDPK='Tabungan'
UNION
SELECT DISTINCT 'DEPOSIT ON CALL' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='Y' and b.LAPOSIM_Level_2='LAPOSIM203000000'
UNION
SELECT DISTINCT 'DEPOSITO' AS FLAG,SUM(jumlahbulanlaporan)AS JumlahBulanLaporan_Valas 
FROM $table_banklain a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang<>'IDR'
AND a.FlagDPK='Deposito' and a.status_oncall='N' and b.LAPOSIM_Level_2='LAPOSIM203000000'
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



// SHEET KE 2 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->getStyle('A1:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A20:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('AS1:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A6:AR17')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A1:AR9')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(20);

$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(15);


$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth(20);



$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth(20);

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:H2');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A6:A8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B6:B8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C6:H6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C7:E7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('F7:H7');

//$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I2:P2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I1:V2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('W1:AJ2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AK1:AR2');


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I6:I8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('J6:J8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K6:P6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K7:M7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('N7:P7');


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('Q6:V6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('Q7:S7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('T7:V7');



$objPHPExcel->setActiveSheetIndex(1)->mergeCells('W6:W8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('X6:X8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('Y6:AD6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('Y7:AA7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AB7:AD7');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AE6:AJ6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AE7:AG7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AH7:AJ7');



$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AM6:AR6');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AK6:AK8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AL6:AL8');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AM6:AR6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AM7:AO7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AP7:AR7');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A17:B17');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I17:J17');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('W17:X17');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('AK17:AL17');


$objPHPExcel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A6:A8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A6:A8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('B6:AR8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B6:AR8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('I1:AR2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('I1:AR2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI PIHAK KETIGA");
#====================== GIRO =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('A3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('A4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('A5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('C3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('C4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('C5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('A6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('C6', "Giro");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('F7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('E8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('H8', "Sub Total");

#====================== Tabungan =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('I1', "RINCIAN POSISI PIHAK KETIGA");
$objPHPExcel->getActiveSheet()->setCellValue('I3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('I4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('I5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('K3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('K4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('K5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('I6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('J6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('K6', "Tabungan");
$objPHPExcel->getActiveSheet()->setCellValue('K7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('L8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('N8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('O8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('P8', "Sub Total");

$objPHPExcel->getActiveSheet()->setCellValue('Q6', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('T7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('Q8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('R8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('S8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('T8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('U8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('V8', "Sub Total");

#====================== Deposito =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('W1', "RINCIAN POSISI PIHAK KETIGA");
$objPHPExcel->getActiveSheet()->setCellValue('W3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('W4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('W5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('Y3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('Y4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('Y5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('W6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('X6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('Y6', "Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('Y7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AB7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('Y8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('Z8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AA8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AB8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AC8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AD8', "Sub Total");

$objPHPExcel->getActiveSheet()->setCellValue('AE6', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('AE7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AH7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('AE8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AF8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AG8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AH8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AI8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AJ8', "Sub Total");

#====================== Jumlah =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('AK1', "RINCIAN POSISI PIHAK KETIGA");
$objPHPExcel->getActiveSheet()->setCellValue('AK3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('AK4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('AK5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('AM3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('AM4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('AM5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('AK6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('AL6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('AM6', " Jumlah *) ");
$objPHPExcel->getActiveSheet()->setCellValue('AM7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AP7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('AM8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AN8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AO8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AP8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AQ8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AR8', "Sub Total");


$objPHPExcel->getActiveSheet()->setCellValue('A17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('I17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('W17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('AK17', "Total Simpanan");





//$counter=0;
$i=1;
$rowexcel=10;
foreach ($label_nominal as $nilai ) {
 
$objPHPExcel->getActiveSheet()->setCellValue("A$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("B$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("I$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("J$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("W$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("X$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("AK$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("AL$rowexcel", "$nilai");


//$counter++;
$i++;
$rowexcel++;

}
/*

select count(norekening) from dm_LiabilitasDeposito where datadate='2016-07-31' 
--and JumlahBulanLaporan >='0' and JumlahBulanLaporan <'100000001'
--and JumlahBulanLaporan >='100000001' and JumlahBulanLaporan <'200000001'
--and JumlahBulanLaporan >='200000001' and JumlahBulanLaporan <'500000001'
--and JumlahBulanLaporan >='500000001' and JumlahBulanLaporan <'1000000001'
--and JumlahBulanLaporan >='1000000001' and JumlahBulanLaporan <'2000000001'
--and JumlahBulanLaporan >='2000000001' and JumlahBulanLaporan <'5000000001'
and JumlahBulanLaporan > '5000000000'
and JenisMataUang='IDR'
and status_oncall='N'




*/

#--Query Lampiran 1D--
#-- GIRO --------------------------------------------------------------------------------------
#--Rekening Giro Rupiah--
$giro_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah 
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM101000000'";
#--Rekening Giro Valas --
$giro_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah 
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM101000000'";
#----Nominal Giro Rupiah--
$giro_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah 
FROM $table_giro a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR'  and b.LAPOSIM_Level_2='LAPOSIM101000000'";
#--Nominal Giro Valas--
$giro_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah 
FROM $table_giro a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR'  and b.LAPOSIM_Level_2='LAPOSIM101000000'";



$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Giro Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($giro_rek_rp,$row1['Jumlah']);
        #--Rekening Giro Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_rek_val,$row2['Jumlah']);
        #----Nominal Giro Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($giro_nom_rp,$row3['Jumlah']);
        #--Nominal Giro Valas--
        //echo $query_nom_val.$array_range1["$counter2"];
        //die();
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($giro_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    
#-- TABUNGAN --------------------------------------------------------------------------------------
#--Rekening Tabungan Rupiah--
$tabungan_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_tabungan a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM102000000' ";
#--Rekening Tabungan Valas --
$tabungan_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_tabungan a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM102000000' ";
#----Nominal Tabungan Rupiah--
$tabungan_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_tabungan a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang='IDR' and   b.LAPOSIM_Level_2='LAPOSIM102000000' ";
#--Nominal Tabungan Valas--
$tabungan_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_tabungan a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR'  and b.LAPOSIM_Level_2='LAPOSIM102000000' ";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Tabungan Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);

        //echo $query_rek_rp.$array_range1["$counter2"]."<br>";
        

        $row1=odbc_fetch_array($result1);
        array_push($tabungan_rek_rp,$row1['Jumlah']);
        //echo $row1['Jumlah'];
        //die();
        #--Rekening Tabungan Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_rek_val,$row2['Jumlah']);
        #----Nominal Tabungan Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($tabungan_nom_rp,$row3['Jumlah']);
        #--Nominal Tabungan Valas--
        //echo $query_nom_val.$array_range1["$counter2"];
        //die();
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($tabungan_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    


#-- DEPOSITO OM CALL  --------------------------------------------------------------------------------------
#--Rekening Deposito Rupiah--
$deposito_call_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.status_oncall='Y' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000' ";
#--Rekening Deposito Valas --
$deposito_call_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.status_oncall='Y' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000' ";
#----Nominal Deposito Rupiah--
$deposito_call_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.status_oncall='Y'  and b.LAPOSIM_Level_2='LAPOSIM103000000' ";
#--Nominal Deposito Valas--
$deposito_call_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.status_oncall='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000' ";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Deposito Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($deposito_call_rek_rp,$row1['Jumlah']);
        #--Rekening Deposito Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rek_val,$row2['Jumlah']);
        #----Nominal Deposito Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($deposito_call_nom_rp,$row3['Jumlah']);
        //var_dump($deposito_call_nom_rp);
        //die();
        #--Nominal Deposito Valas--
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($deposito_call_nom_val,$row4['Jumlah']);

    $counter2++;   
}
 /*   
var_dump($deposito_call_rek_rp);
echo "<br>";
var_dump($deposito_call_rek_val);
echo "<br>";
var_dump($deposito_call_nom_rp);
echo "<br>";
var_dump($deposito_call_nom_val);
echo "<br>";
die();
*/






#-- DEPOSITO --------------------------------------------------------------------------------------
#--Rekening Deposito Rupiah--
$deposito_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.status_oncall='N' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000'";
#--Rekening Deposito Valas --
$deposito_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.status_oncall='N' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM103000000'";
#----Nominal Deposito Rupiah--
$deposito_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.status_oncall='N'  and b.LAPOSIM_Level_2='LAPOSIM103000000'";
#--Nominal Deposito Valas--
$deposito_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_deposito a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.status_oncall='N'  and b.LAPOSIM_Level_2='LAPOSIM103000000'";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Deposito Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($deposito_rek_rp,$row1['Jumlah']);
        #--Rekening Deposito Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rek_val,$row2['Jumlah']);
        #----Nominal Deposito Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($deposito_nom_rp,$row3['Jumlah']);
        #--Nominal Deposito Valas--
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($deposito_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    
/*
var_dump($deposito_rek_rp);
echo "<br>";
var_dump($deposito_rek_val);
echo "<br>";
var_dump($deposito_nom_rp);
echo "<br>";
var_dump($deposito_nom_val);
echo "<br>";
die();
*/


# PRINT GIRO--------------------------------

$index=10;
foreach ($giro_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("C$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$index", floatval($nilai ));
$index++;
}

# PRINT TABUNGAN--------------------------------

$index=10;
foreach ($tabungan_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("N$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("O$index", floatval($nilai ));
$index++;
}


# PRINT DEPOSITO ONCALL--------------------------------

$index=10;
foreach ($deposito_call_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Q$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("R$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("T$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("U$index", floatval($nilai ));
$index++;
}






# PRINT DEPOSITO--------------------------------

$index=10;
foreach ($deposito_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Y$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Z$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("AB$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("AC$index", floatval($nilai ));
$index++;
}






for ($i=10; $i <= 16; $i++) { 

$objPHPExcel->getActiveSheet()->setCellValue("AM$i", "=C$i+K$i+Q$i+Y$i+AE$i");
$objPHPExcel->getActiveSheet()->setCellValue("AN$i", "=D$i+L$i+R$i+Z$i+AF$i");
$objPHPExcel->getActiveSheet()->setCellValue("AP$i", "=F$i+N$i+T$i+AB$i+AH$i");
$objPHPExcel->getActiveSheet()->setCellValue("AQ$i", "=G$i+O$i+U$i+AC$i+AI$i");


}

// SUM SUBTOTAL=================================

for ($i=10; $i <= 16; $i++) { 

$objPHPExcel->getActiveSheet()->setCellValue("E$i", "=C$i+D$i");
$objPHPExcel->getActiveSheet()->setCellValue("H$i", "=F$i+G$i");
$objPHPExcel->getActiveSheet()->setCellValue("M$i", "=K$i+L$i");
$objPHPExcel->getActiveSheet()->setCellValue("P$i", "=N$i+O$i");
$objPHPExcel->getActiveSheet()->setCellValue("S$i", "=Q$i+R$i");
$objPHPExcel->getActiveSheet()->setCellValue("V$i", "=T$i+U$i");
$objPHPExcel->getActiveSheet()->setCellValue("AA$i", "=Y$i+Z$i");
$objPHPExcel->getActiveSheet()->setCellValue("AD$i", "=AB$i+AC$i");
$objPHPExcel->getActiveSheet()->setCellValue("AG$i", "=AE$i+AF$i");
$objPHPExcel->getActiveSheet()->setCellValue("AJ$i", "=AH$i+AI$i");
$objPHPExcel->getActiveSheet()->setCellValue("AO$i", "=AM$i+AN$i");
$objPHPExcel->getActiveSheet()->setCellValue("AR$i", "=AP$i+AQ$i");


}



$objPHPExcel->getActiveSheet()->setCellValue("C17", "=SUM(C10:C16)");
$objPHPExcel->getActiveSheet()->setCellValue("D17", "=SUM(D10:D16)");
$objPHPExcel->getActiveSheet()->setCellValue("E17", "=SUM(E10:E16)");
$objPHPExcel->getActiveSheet()->setCellValue("F17", "=SUM(F10:F16)");
$objPHPExcel->getActiveSheet()->setCellValue("G17", "=SUM(G10:G16)");
$objPHPExcel->getActiveSheet()->setCellValue("H17", "=SUM(H10:H16)");
$objPHPExcel->getActiveSheet()->setCellValue("K17", "=SUM(K10:K16)");
$objPHPExcel->getActiveSheet()->setCellValue("L17", "=SUM(L10:L16)");
$objPHPExcel->getActiveSheet()->setCellValue("M17", "=SUM(M10:M16)");
$objPHPExcel->getActiveSheet()->setCellValue("N17", "=SUM(N10:N16)");
$objPHPExcel->getActiveSheet()->setCellValue("O17", "=SUM(O10:O16)");
$objPHPExcel->getActiveSheet()->setCellValue("P17", "=SUM(P10:P16)");
$objPHPExcel->getActiveSheet()->setCellValue("Q17", "=SUM(Q10:Q16)");
$objPHPExcel->getActiveSheet()->setCellValue("R17", "=SUM(R10:R16)");
$objPHPExcel->getActiveSheet()->setCellValue("S17", "=SUM(S10:S16)");
$objPHPExcel->getActiveSheet()->setCellValue("T17", "=SUM(T10:T16)");
$objPHPExcel->getActiveSheet()->setCellValue("U17", "=SUM(U10:U16)");
$objPHPExcel->getActiveSheet()->setCellValue("V17", "=SUM(V10:V16)");

$objPHPExcel->getActiveSheet()->setCellValue("Y17", "=SUM(Y10:Y16)");
$objPHPExcel->getActiveSheet()->setCellValue("Z17", "=SUM(Z10:Z16)");
$objPHPExcel->getActiveSheet()->setCellValue("AA17", "=SUM(AA10:AA16)");
$objPHPExcel->getActiveSheet()->setCellValue("AB17", "=SUM(AB10:AB16)");
$objPHPExcel->getActiveSheet()->setCellValue("AC17", "=SUM(AC10:AC16)");
$objPHPExcel->getActiveSheet()->setCellValue("AD17", "=SUM(AD10:AD16)");
$objPHPExcel->getActiveSheet()->setCellValue("AE17", "=SUM(AE10:AE16)");
$objPHPExcel->getActiveSheet()->setCellValue("AF17", "=SUM(AF10:AF16)");
$objPHPExcel->getActiveSheet()->setCellValue("AG17", "=SUM(AG10:AG16)");
$objPHPExcel->getActiveSheet()->setCellValue("AH17", "=SUM(AH10:AH16)");
$objPHPExcel->getActiveSheet()->setCellValue("AI17", "=SUM(AI10:AI16)");
$objPHPExcel->getActiveSheet()->setCellValue("AJ17", "=SUM(AJ10:AJ16)");

$objPHPExcel->getActiveSheet()->setCellValue("AM17", "=SUM(AM10:AM16)");
$objPHPExcel->getActiveSheet()->setCellValue("AN17", "=SUM(AN10:AN16)");
$objPHPExcel->getActiveSheet()->setCellValue("AO17", "=SUM(AO10:AO16)");
$objPHPExcel->getActiveSheet()->setCellValue("AP17", "=SUM(AP10:AP16)");
$objPHPExcel->getActiveSheet()->setCellValue("AQ17", "=SUM(AQ10:AQ16)");
$objPHPExcel->getActiveSheet()->setCellValue("AR17", "=SUM(AR10:AR16)");

$objPHPExcel->getActiveSheet()->getStyle('F10:H17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:P17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('T10:V17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AB10:AD17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AH10:AJ17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AP10:AR17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');



$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1D');





// SHEET KE 3 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2); 

$objPHPExcel->getActiveSheet()->getStyle('A1:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A20:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('AS1:AZ25')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A6:AR17')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A1:AR9')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(20);

$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(15);


$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth(20);

$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth(20);

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A2:H2');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A6:A8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B6:B8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C6:H6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C7:E7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('F7:H7');

//$objPHPExcel->setActiveSheetIndex(2)->mergeCells('I2:P2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('I1:V2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('W1:AJ2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AK1:AR2');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('I6:I8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('J6:J8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('K6:P6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('K7:M7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('N7:P7');


$objPHPExcel->setActiveSheetIndex(2)->mergeCells('Q6:V6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('Q7:S7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('T7:V7');



$objPHPExcel->setActiveSheetIndex(2)->mergeCells('W6:W8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('X6:X8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('Y6:AD6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('Y7:AA7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AB7:AD7');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AE6:AJ6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AE7:AG7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AH7:AJ7');



$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AM6:AR6');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AK6:AK8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AL6:AL8');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AM6:AR6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AM7:AO7');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AP7:AR7');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A17:B17');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('I17:J17');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('W17:X17');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('AK17:AL17');


$objPHPExcel->getActiveSheet()->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A6:A8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A6:A8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('B6:AR8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B6:AR8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('I1:AR2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('I1:AR2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->setCellValue('A2', "RINCIAN POSISI SIMPANAN DARI BANK LAIN");
#====================== GIRO =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('A3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('A4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('A5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('C3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('C4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('C5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('A6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('C6', "Giro");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('F7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('E8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('H8', "Sub Total");

#====================== Tabungan =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('I1', "RINCIAN POSISI SIMPANAN DARI BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('I3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('I4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('I5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('K3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('K4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('K5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('I6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('J6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('K6', "Tabungan");
$objPHPExcel->getActiveSheet()->setCellValue('K7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('N7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('L8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('N8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('O8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('P8', "Sub Total");

$objPHPExcel->getActiveSheet()->setCellValue('Q6', "Deposito On Call");
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('T7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('Q8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('R8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('S8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('T8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('U8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('V8', "Sub Total");

#====================== Deposito =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('W1', "RINCIAN POSISI SIMPANAN DARI BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('W3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('W4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('W5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('Y3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('Y4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('Y5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('W6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('X6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('Y6', "Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('Y7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AB7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('Y8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('Z8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AA8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AB8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AC8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AD8', "Sub Total");

$objPHPExcel->getActiveSheet()->setCellValue('AE6', "Sertifikat Deposito");
$objPHPExcel->getActiveSheet()->setCellValue('AE7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AH7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('AE8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AF8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AG8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AH8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AI8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AJ8', "Sub Total");

#====================== Jumlah =======================================================================
$objPHPExcel->getActiveSheet()->setCellValue('AK1', "RINCIAN POSISI SIMPANAN DARI BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('AK3', "PERAKHIR BULAN : ");
$objPHPExcel->getActiveSheet()->setCellValue('AK4', "TAHUN: ");
$objPHPExcel->getActiveSheet()->setCellValue('AK5', "BANK: ");

$objPHPExcel->getActiveSheet()->setCellValue('AM3', "$label_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('AM4', "'$year_modal");
$objPHPExcel->getActiveSheet()->setCellValue('AM5', "BANK MNC Internasional, Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('AK6', "No");
$objPHPExcel->getActiveSheet()->setCellValue('AL6', "Jumlah Nominal (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('AM6', " Jumlah *) ");
$objPHPExcel->getActiveSheet()->setCellValue('AM7', "Jumlah Rekening");
$objPHPExcel->getActiveSheet()->setCellValue('AP7', "Jumlah Nominal");
$objPHPExcel->getActiveSheet()->setCellValue('AM8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AN8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AO8', "Sub Total");
$objPHPExcel->getActiveSheet()->setCellValue('AP8', "Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('AQ8', "Valas");
$objPHPExcel->getActiveSheet()->setCellValue('AR8', "Sub Total");

$objPHPExcel->getActiveSheet()->setCellValue('A17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('I17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('W17', "Total Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('AK17', "Total Simpanan");

//$counter=0;
$i=1;
$rowexcel=10;
foreach ($label_nominal as $nilai ) {
 
$objPHPExcel->getActiveSheet()->setCellValue("A$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("B$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("I$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("J$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("W$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("X$rowexcel", "$nilai");

$objPHPExcel->getActiveSheet()->setCellValue("AK$rowexcel", "$i");
$objPHPExcel->getActiveSheet()->setCellValue("AL$rowexcel", "$nilai");


//$counter++;
$i++;
$rowexcel++;

}





#--Query Lampiran 1E--
#-- GIRO --------------------------------------------------------------------------------------
#--Rekening Giro Rupiah--
$giro_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Giro' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM201000000' ";

#--Rekening Giro Valas --
$giro_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Giro' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM201000000' ";
#----Nominal Giro Rupiah--
$giro_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Giro'  and b.LAPOSIM_Level_2='LAPOSIM201000000' ";
#--Nominal Giro Valas--
$giro_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Giro'  and b.LAPOSIM_Level_2='LAPOSIM201000000' ";



$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Giro Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        //echo $query_rek_rp.$array_range1["$counter2"];
        //die();
        $row1=odbc_fetch_array($result1);
        array_push($giro_rek_rp,$row1['Jumlah']);
        #--Rekening Giro Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($giro_rek_val,$row2['Jumlah']);
        #----Nominal Giro Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($giro_nom_rp,$row3['Jumlah']);
        #--Nominal Giro Valas--
        //echo $query_nom_val.$array_range1["$counter2"];
        //die();
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($giro_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    
#-- TABUNGAN --------------------------------------------------------------------------------------
#--Rekening Tabungan Rupiah--
$tabungan_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Tabungan' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM202000000' ";
#--Rekening Tabungan Valas --
$tabungan_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Tabungan' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM202000000' ";
#----Nominal Tabungan Rupiah--
$tabungan_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang = 'IDR' AND a.flagDPK='Tabungan'  and b.LAPOSIM_Level_2='LAPOSIM202000000' ";
#--Nominal Tabungan Valas--
$tabungan_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2 
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Tabungan'  and b.LAPOSIM_Level_2='LAPOSIM202000000' ";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Tabungan Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($tabungan_rek_rp,$row1['Jumlah']);

        //echo $query_rek_rp.$array_range1["$counter2"]."<br>";
        //echo $row1['Jumlah'];
        //die();
        #--Rekening Tabungan Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($tabungan_rek_val,$row2['Jumlah']);
        #----Nominal Tabungan Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($tabungan_nom_rp,$row3['Jumlah']);
        #--Nominal Tabungan Valas--
        //echo $query_nom_val.$array_range1["$counter2"];
        //die();
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($tabungan_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    
#-- DEPOSITO ON CALL--------------------------------------------------------------------------------------
#--Rekening Deposito Rupiah--
$deposito_call_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='Y' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#--Rekening Deposito Valas --
$deposito_call_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='Y' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#----Nominal Deposito Rupiah--
$deposito_call_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='Y'  and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#--Nominal Deposito Valas--
$deposito_call_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='Y'  and b.LAPOSIM_Level_2='LAPOSIM203000000' ";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Deposito Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($deposito_call_rek_rp,$row1['Jumlah']);
        #--Rekening Deposito Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_call_rek_val,$row2['Jumlah']);
        #----Nominal Deposito Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($deposito_call_nom_rp,$row3['Jumlah']);
        #--Nominal Deposito Valas--
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($deposito_call_nom_val,$row4['Jumlah']);

    $counter2++;   
}

#-- DEPOSITO --------------------------------------------------------------------------------------
#--Rekening Deposito Rupiah--
$deposito_rek_rp=array();
$query_rek_rp =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='N' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#--Rekening Deposito Valas --
$deposito_rek_val=array();
$query_rek_val =" SELECT COUNT(NoRekening) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='N' and a.StatusLaporan_LBU='Y' and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#----Nominal Deposito Rupiah--
$deposito_nom_rp=array();
$query_nom_rp =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang ='IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='N'  and b.LAPOSIM_Level_2='LAPOSIM203000000' ";
#--Nominal Deposito Valas--
$deposito_nom_val=array();
$query_nom_val =" SELECT SUM(jumlahbulanlaporan) as Jumlah
FROM $table_banklain a 
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_LAPOSIM c ON c.LAPOSIM_Level_2 = b.LAPOSIM_Level_2
WHERE a.DataDate ='$curr_tgl' AND a.JenisMataUang <>'IDR' AND a.flagDPK='Deposito'  AND a.status_oncall='N'  and b.LAPOSIM_Level_2='LAPOSIM203000000' ";

$counter2=0;

foreach ($array_range1 as $value2) {
        #--Rekening Deposito Rupiah--
        $result1=odbc_exec($connection2, $query_rek_rp.$array_range1["$counter2"]);
        $row1=odbc_fetch_array($result1);
        array_push($deposito_rek_rp,$row1['Jumlah']);
        #--Rekening Deposito Valas--
        $result2=odbc_exec($connection2, $query_rek_val.$array_range1["$counter2"]);
        $row2=odbc_fetch_array($result2);
        array_push($deposito_rek_val,$row2['Jumlah']);
        #----Nominal Deposito Rupiah--
        $result3=odbc_exec($connection2, $query_nom_rp.$array_range1["$counter2"]);
        $row3=odbc_fetch_array($result3);
        array_push($deposito_nom_rp,$row3['Jumlah']);
        #--Nominal Deposito Valas--
        $result4=odbc_exec($connection2, $query_nom_val.$array_range1["$counter2"]);
        $row4=odbc_fetch_array($result4);
        array_push($deposito_nom_val,$row4['Jumlah']);

    $counter2++;   
}
    

# PRINT GIRO--------------------------------




$index=10;
foreach ($giro_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("C$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("D$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("F$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($giro_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("G$index", floatval($nilai ));
$index++;
}

# PRINT TABUNGAN--------------------------------

$index=10;
foreach ($tabungan_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("K$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("L$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("N$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($tabungan_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("O$index", floatval($nilai ));
$index++;
}


# PRINT DEPOSITO ONCALL--------------------------------*




$index=10;
foreach ($deposito_call_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Q$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("R$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("T$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_call_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("U$index", floatval($nilai ));
$index++;
}





# PRINT DEPOSITO--------------------------------



$index=10;
foreach ($deposito_rek_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Y$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_rek_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("Z$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_nom_rp as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("AB$index", floatval($nilai ));
$index++;
}
$index=10;
foreach ($deposito_nom_val as $nilai ) {
  $objPHPExcel->getActiveSheet()->setCellValue("AC$index", floatval($nilai ));
$index++;
}






for ($i=10; $i <= 16; $i++) { 

$objPHPExcel->getActiveSheet()->setCellValue("AM$i", "=C$i+K$i+Q$i+Y$i+AE$i");
$objPHPExcel->getActiveSheet()->setCellValue("AN$i", "=D$i+L$i+R$i+Z$i+AF$i");
$objPHPExcel->getActiveSheet()->setCellValue("AP$i", "=F$i+N$i+T$i+AB$i+AH$i");
$objPHPExcel->getActiveSheet()->setCellValue("AQ$i", "=G$i+O$i+U$i+AC$i+AI$i");


}


// SUM SUBTOTAL=================================

for ($i=10; $i <= 16; $i++) { 

$objPHPExcel->getActiveSheet()->setCellValue("E$i", "=C$i+D$i");
$objPHPExcel->getActiveSheet()->setCellValue("H$i", "=F$i+G$i");
$objPHPExcel->getActiveSheet()->setCellValue("M$i", "=K$i+L$i");
$objPHPExcel->getActiveSheet()->setCellValue("P$i", "=N$i+O$i");
$objPHPExcel->getActiveSheet()->setCellValue("S$i", "=Q$i+R$i");
$objPHPExcel->getActiveSheet()->setCellValue("V$i", "=T$i+U$i");
$objPHPExcel->getActiveSheet()->setCellValue("AA$i", "=Y$i+Z$i");
$objPHPExcel->getActiveSheet()->setCellValue("AD$i", "=AB$i+AC$i");
$objPHPExcel->getActiveSheet()->setCellValue("AG$i", "=AE$i+AF$i");
$objPHPExcel->getActiveSheet()->setCellValue("AJ$i", "=AH$i+AI$i");
$objPHPExcel->getActiveSheet()->setCellValue("AO$i", "=AM$i+AN$i");
$objPHPExcel->getActiveSheet()->setCellValue("AR$i", "=AP$i+AQ$i");


}


$objPHPExcel->getActiveSheet()->setCellValue("C17", "=SUM(C10:C16)");
$objPHPExcel->getActiveSheet()->setCellValue("D17", "=SUM(D10:D16)");
$objPHPExcel->getActiveSheet()->setCellValue("E17", "=SUM(E10:E16)");
$objPHPExcel->getActiveSheet()->setCellValue("F17", "=SUM(F10:F16)");
$objPHPExcel->getActiveSheet()->setCellValue("G17", "=SUM(G10:G16)");
$objPHPExcel->getActiveSheet()->setCellValue("H17", "=SUM(H10:H16)");
$objPHPExcel->getActiveSheet()->setCellValue("K17", "=SUM(K10:K16)");
$objPHPExcel->getActiveSheet()->setCellValue("L17", "=SUM(L10:L16)");
$objPHPExcel->getActiveSheet()->setCellValue("M17", "=SUM(M10:M16)");
$objPHPExcel->getActiveSheet()->setCellValue("N17", "=SUM(N10:N16)");
$objPHPExcel->getActiveSheet()->setCellValue("O17", "=SUM(O10:O16)");
$objPHPExcel->getActiveSheet()->setCellValue("P17", "=SUM(P10:P16)");
$objPHPExcel->getActiveSheet()->setCellValue("Q17", "=SUM(Q10:Q16)");
$objPHPExcel->getActiveSheet()->setCellValue("R17", "=SUM(R10:R16)");
$objPHPExcel->getActiveSheet()->setCellValue("S17", "=SUM(S10:S16)");
$objPHPExcel->getActiveSheet()->setCellValue("T17", "=SUM(T10:T16)");
$objPHPExcel->getActiveSheet()->setCellValue("U17", "=SUM(U10:U16)");
$objPHPExcel->getActiveSheet()->setCellValue("V17", "=SUM(V10:V16)");

$objPHPExcel->getActiveSheet()->setCellValue("Y17", "=SUM(Y10:Y16)");
$objPHPExcel->getActiveSheet()->setCellValue("Z17", "=SUM(Z10:Z16)");
$objPHPExcel->getActiveSheet()->setCellValue("AA17", "=SUM(AA10:AA16)");
$objPHPExcel->getActiveSheet()->setCellValue("AB17", "=SUM(AB10:AB16)");
$objPHPExcel->getActiveSheet()->setCellValue("AC17", "=SUM(AC10:AC16)");
$objPHPExcel->getActiveSheet()->setCellValue("AD17", "=SUM(AD10:AD16)");
$objPHPExcel->getActiveSheet()->setCellValue("AE17", "=SUM(AE10:AE16)");
$objPHPExcel->getActiveSheet()->setCellValue("AF17", "=SUM(AF10:AF16)");
$objPHPExcel->getActiveSheet()->setCellValue("AG17", "=SUM(AG10:AG16)");
$objPHPExcel->getActiveSheet()->setCellValue("AH17", "=SUM(AH10:AH16)");
$objPHPExcel->getActiveSheet()->setCellValue("AI17", "=SUM(AI10:AI16)");
$objPHPExcel->getActiveSheet()->setCellValue("AJ17", "=SUM(AJ10:AJ16)");

$objPHPExcel->getActiveSheet()->setCellValue("AM17", "=SUM(AM10:AM16)");
$objPHPExcel->getActiveSheet()->setCellValue("AN17", "=SUM(AN10:AN16)");
$objPHPExcel->getActiveSheet()->setCellValue("AO17", "=SUM(AO10:AO16)");
$objPHPExcel->getActiveSheet()->setCellValue("AP17", "=SUM(AP10:AP16)");
$objPHPExcel->getActiveSheet()->setCellValue("AQ17", "=SUM(AQ10:AQ16)");
$objPHPExcel->getActiveSheet()->setCellValue("AR17", "=SUM(AR10:AR16)");




$objPHPExcel->getActiveSheet()->getStyle('F10:H17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N10:P17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('T10:V17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AB10:AD17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AH10:AJ17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('AP10:AR17')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');




$objPHPExcel->getActiveSheet()->setTitle('LAMPIRAN 1E');



$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/SIMP_LPS_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/SIMP_LPS_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);



//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save("download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls");





?>

<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> SIMPANAN LPS (LAPORAN PIHAK KETIGA)
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
                                        LAMPIRAN 1 B </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        LAMPIRAN 1 D </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_3" data-toggle="tab">
                                        LAMPIRAN 1 E </a>
                                    </li> 
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/SIMP_LPS_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div></b> 

                                            
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
                                        <div class="tab-pane" id="tab_15_2">
                                         <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 2</b><br>
                                       <b> RINCIAN POSISI SIMPANAN ANTAR BANK PASSIVA (SIMPANAN DARI BANK LAIN) </b><br>
                                       <b> RUPIAH </b>
                                       <?php
                                        $objPHPExcel->setActiveSheetIndex(1);
                                       ?>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Giro </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Tabungan </b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("O$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("P$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Deposito </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                    <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("W$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("X$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("Y$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("Z$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AA$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AB$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AC$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AD$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Jumlah </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                    <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AK$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AL$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AM$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AN$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AO$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AP$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AQ$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AR$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                
                                                </tbody>
                                            </table>
                                        </div>



                                        </div>
                                        <div class="tab-pane" id="tab_15_3">
                                          <div class="alert alert-info">
                                        <button class="close" data-close="alert"></button>
                                       <b> FORM 3 </b><br>
                                       <b> RINCIAN POSISI DANA PIHAK KETIGA (VALUTA ASING) </b><br>
                                       <?php
                                        $objPHPExcel->setActiveSheetIndex(2);
                                       ?>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Giro </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Tabungan </b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("O$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("P$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Deposito </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                    <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("W$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("X$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("Y$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("Z$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AA$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AB$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AC$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AD$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
                                            }
                                                ?>
                                                
                                                </tbody>
                                            </table>
                                        </div>

                                        <br>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" rowspan="3"><b>No</b></td>
                                                <td width="20%" align="center" rowspan="3"><b>Jumlah<br>Nominal (Rupiah)</b></td>
                                                <td width="75%" align="center" colspan="6"><b>Jumlah </b></td>
                                                
                                               
                                                </tr>
                                                <tr class="active">
                                                 <td width="30%" align="center" colspan="3"><b>Jml. Rekening</b></td>
                                                 <td width="45%" align="center" colspan="3"><b>Jml. Nominal</b></td>
                                                </tr>
                                                <tr class="active">
                                                 <td width="10%" align="center" ><b>Rupiah</b></td>
                                                 <td width="10%" align="center" ><b>Valas</b></td>
                                                 <td width="10%" align="center" ><b>Sub Total</b></td>
                                                 <td width="15%" align="center" ><b>Rupiah</b></td>
                                                 <td width="15%" align="center" ><b>Valas</b></td>
                                                 <td width="15%" align="center" ><b>Sub Total</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                    <tr >
                                                <?php 
                                                for ($i=10; $i <=16 ; $i++) { 
                                                
                                                ?>
                                                 <td width="5%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AK$i")->getValue(); ?></b></td>
                                                 <td width="20%" align="left" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AL$i")->getValue(); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AM$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AN$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="10%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AO$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AP$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AQ$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>
                                                 <td width="15%" align="right" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("AR$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></b></td>

                                                </tr>
                                                <?php
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
                </div>

