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
logActivity("generate counter rate",date('Y_m_d_H_i_s'));

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


#--Jumlah Deposito Berjangka-- 
#--Deposito Rupiah 1 Bulan--
/* ## OLD 12-08-2016
$query = "SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(JumlahBulanLaporan) AS Deposito
FROM DM_LiabilitasDeposito 
WHERE datadate='$curr_tgl' 
and JumlahBulanLaporan between '0' and '100000001'
and DateDiff(Day,JangkaWaktuMulai,JangkaWaktuJatuhTempo) <='33'
and JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(JumlahBulanLaporan) AS Deposito
FROM DM_LiabilitasKepadaBankLain
where datadate='$curr_tgl' AND FlagDPK ='Deposito' 
and JumlahBulanLaporan between '0' and '100000000'
and DateDiff(Day,JangkaWaktuMulai,JangkaWaktuJatuhTempo) <='33'
and JenisMataUang = 'IDR'
) as tabel1 ";
*/

$var_tabel=date('Ymd',strtotime($tanggal));


#############################################################################################
//$table_giro="DM_LiabilitasGiro_$var_tabel";
//$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

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



/*
$query = "SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(JumlahBulanLaporan) AS Deposito
FROM $table_deposito 
WHERE datadate='$curr_tgl' 
and JumlahBulanLaporan >='0' and JumlahBulanLaporan <'100000001'
and DateDiff(Day,JangkaWaktuMulai,JangkaWaktuJatuhTempo) <='33'
and JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(JumlahBulanLaporan) AS Deposito
FROM $table_banklain
where datadate='$curr_tgl' AND FlagDPK ='Deposito' 
and JumlahBulanLaporan >='0' and JumlahBulanLaporan <'100000001'
and DateDiff(Day,JangkaWaktuMulai,JangkaWaktuJatuhTempo) <='33'
and JenisMataUang = 'IDR'
) as tabel1 ";
*/
$query=" SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1  ";



        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c4=$row_res['1_Bulan_Rupiah'];

//=============== and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'  
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c5=$row_res['1_Bulan_Rupiah'];

//====================== and a.JumlahBulanLaporan >'200000001' and a.JumlahBulanLaporan<'500000001'
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >'200000001' and a.JumlahBulanLaporan<'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >'200000001' and a.JumlahBulanLaporan<'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c6=$row_res['1_Bulan_Rupiah'];

//echo $query;
//die();

//===========and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >'200000001' and a.JumlahBulanLaporan<'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >'200000001' and a.JumlahBulanLaporan<'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c7=$row_res['1_Bulan_Rupiah'];
//===============  and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001' 
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c8=$row_res['1_Bulan_Rupiah'];
//==============  and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1   ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c9=$row_res['1_Bulan_Rupiah'];
//=============================== and a.JumlahBulanLaporan > '5000000000'
$query = " SELECT SUM(Deposito) AS '1_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'  
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang = 'IDR'
) as tabel1   ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $c10=$row_res['1_Bulan_Rupiah'];

#######################################################################################################################

# --Deposito Rupiah 3 Bulan--
// and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
$query=" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001' 
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001' 
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d4=$row_res['3_Bulan_Rupiah'];

############################
// and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
$query=" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1 ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d5=$row_res['3_Bulan_Rupiah'];

############################
// and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
$query="  SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1 ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d6=$row_res['3_Bulan_Rupiah'];

############################
// and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
$query=" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1 ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d7=$row_res['3_Bulan_Rupiah'];
############################
// and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
$query=" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d8=$row_res['3_Bulan_Rupiah'];
############################
//  and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001' 
$query =" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1 ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d9=$row_res['3_Bulan_Rupiah'];
############################
//  and a.JumlahBulanLaporan > '5000000000'
$query=" SELECT SUM(Deposito) AS '3_Bulan_Rupiah' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR' 
UNION ALL 
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang = 'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $d10=$row_res['3_Bulan_Rupiah'];

#########################################################################################################################
# --Deposito Valas 1 Bulan--
// and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1  ";
        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e4=$row_res['1_Bulan_Valas'];

############################
// and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'

$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1  ";
         $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e5=$row_res['1_Bulan_Valas'];

############################
//  and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
         $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e6=$row_res['1_Bulan_Valas'];

############################

//  and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1  ";
         $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e7=$row_res['1_Bulan_Valas'];
############################
//  and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
$query ="  SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
         $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e8=$row_res['1_Bulan_Valas'];
############################
// and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1 ";

         $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e9=$row_res['1_Bulan_Valas'];
############################
//  and a.JumlahBulanLaporan > '5000000000'
$query =" SELECT SUM(Deposito) AS '1_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='33'
and a.JenisMataUang <>'IDR'
) as tabel1 ";

        $res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $e10=$row_res['1_Bulan_Valas'];
############################
//  and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
#--Deposito Valas 3 Bulan--
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='0' and a.JumlahBulanLaporan <'100000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";

//echo $query;
//die();
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f4=$row_res['3_Bulan_Valas'];
 
#########################
//  and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='100000001' and a.JumlahBulanLaporan <'200000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f5=$row_res['3_Bulan_Valas'];
#########################
// and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='200000001' and a.JumlahBulanLaporan <'500000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f6=$row_res['3_Bulan_Valas'];
########################
//  and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='500000001' and a.JumlahBulanLaporan <'1000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f7=$row_res['3_Bulan_Valas'];
########################
// and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='1000000001' and a.JumlahBulanLaporan <'2000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f8=$row_res['3_Bulan_Valas'];
########################
//  and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'        
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan >='2000000001' and a.JumlahBulanLaporan <'5000000001'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f9=$row_res['3_Bulan_Valas'];
########################
//  and a.JumlahBulanLaporan > '5000000000'
$query =" SELECT SUM(Deposito) AS '3_Bulan_Valas' FROM (
SELECT SUM(a.JumlahBulanLaporan) AS Deposito FROM $table_deposito a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
WHERE b.CRLPS_Level_2='CRLPS101000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
UNION ALL
SELECT SUM(a.JumlahBulanLaporan) AS Deposito
FROM $table_banklain a
JOIN Referensi_GL_02 b ON b.GLNO = a.Managed_GL_Code AND b.PRODNO = a.Managed_GL_Prod_Code
JOIN Referensi_Counter_Rate_LPS c ON c.CRLPS_Level_2 = b.CRLPS_Level_2
where a.FlagDPK ='Deposito' AND b.CRLPS_Level_2='CRLPS201000000'
and a.JumlahBulanLaporan > '5000000000'
and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) >'33' and DateDiff(Day,a.JangkaWaktuMulai,a.JangkaWaktuJatuhTempo) <='95'
and a.JenisMataUang <>'IDR'
) as tabel1 ";
$res=odbc_exec($connection2, $query);
        $row_res=odbc_fetch_array($res);
        $f10=$row_res['3_Bulan_Valas'];



######################## FROM MASTER_COUNTER_RATE ##################################################
//========================= IDR ============================
//$tanggal=$_POST['tanggal']; 

$var_bulan=date('m',strtotime($tanggal));
$var_tahun=date('Y',strtotime($tanggal));

$query_c_idr =" select * from Master_Counter_Rate WHERE Month(DataDate)='$var_bulan' and Year(DataDate)='$var_tahun' and JenisMatauang ='IDR' ";
        $res_idr=odbc_exec($connection2, $query_c_idr);
        $row_res=odbc_fetch_array($res_idr);
        $min_rate1_idr=$row_res['Min_Rate1'];
        $max_rate1_idr=$row_res['Max_Rate1'];
        $min_rate3_idr=$row_res['Min_Rate3'];
        $max_rate3_idr=$row_res['Max_Rate3'];

//echo $query_c_idr;
//die();

$query_c_usd =" select * from Master_Counter_Rate WHERE Month(DataDate)='$var_bulan' and Year(DataDate)='$var_tahun' and JenisMatauang ='USD' ";

        $res_usd=odbc_exec($connection2, $query_c_usd);
        $row_res2=odbc_fetch_array($res_usd);
        $min_rate1_usd=$row_res2['Min_Rate1'];
        $max_rate1_usd=$row_res2['Max_Rate1'];
        $min_rate3_usd=$row_res2['Min_Rate3'];
        $max_rate3_usd=$row_res2['Max_Rate3'];


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


$objPHPExcel->getActiveSheet()->getStyle('B2:F3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B14:H15')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B26:H27')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('B3:H3')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B27:H28')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B15:H16')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C11')->applyFromArray($styleArrayAlignment1);
//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B3:F11')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B15:H23')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B27:H35')->applyFromArray($styleArrayBorder1);

//$objPHPExcel->getActiveSheet()->getStyle('D15:D57')->applyFromArray($styleArray);


//$objPHPExcel->getActiveSheet()->getStyle('I34:I57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

//FILL COLOR
/*
$objPHPExcel->getActiveSheet()->getStyle('A1:Z12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A58:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('L13:Z57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F9:Z10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A13:A57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A9:A10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
*/
//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(25);
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C2:D2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E2:F2'); 
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C11:F11');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C14:D14');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B15:B16');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C15:D15');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E15:E16');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F15:G15');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H15:H16');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B17:B23');


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C26:D26');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B27:B28');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C27:D27');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E27:E28');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F27:G27');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H27:H28');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B29:B35');


//PT Sinergi Multi Andalan

$objPHPExcel->getActiveSheet()->setCellValue('C2', 'RUPIAH');
$objPHPExcel->getActiveSheet()->setCellValue('E2', 'VALUTA ASING');

$objPHPExcel->getActiveSheet()->setCellValue('B3', 'Deposito Berjangka');
$objPHPExcel->getActiveSheet()->setCellValue('C3', '1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('D3', '3 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('E3', '1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('F3', '3 Bulan');

$objPHPExcel->getActiveSheet()->setCellValue('B4', '0 < N <= 100 Jt');
$objPHPExcel->getActiveSheet()->setCellValue('B5', '100 Jt < N <= 200 Jt');
$objPHPExcel->getActiveSheet()->setCellValue('B6', '200 Jt < N <= 500 Jt');
$objPHPExcel->getActiveSheet()->setCellValue('B7', '500 Jt < N <= 1 M');
$objPHPExcel->getActiveSheet()->setCellValue('B8', '1 M < N <= 2 M');
$objPHPExcel->getActiveSheet()->setCellValue('B9', '2 M < N <= 5 M');
$objPHPExcel->getActiveSheet()->setCellValue('B10', '> 5 Milyar');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Posisi Tanggal');
$objPHPExcel->getActiveSheet()->setCellValue('C11', $label_tgl);
############################## 1 Bulan RP ###################################
$objPHPExcel->getActiveSheet()->setCellValue('C4', floatval($c4));
$objPHPExcel->getActiveSheet()->setCellValue('C5', floatval($c5));
$objPHPExcel->getActiveSheet()->setCellValue('C6', floatval($c6));
$objPHPExcel->getActiveSheet()->setCellValue('C7', floatval($c7));
$objPHPExcel->getActiveSheet()->setCellValue('C8', floatval($c8));
$objPHPExcel->getActiveSheet()->setCellValue('C9', floatval($c9));
$objPHPExcel->getActiveSheet()->setCellValue('C10', floatval($c10));
############################## 3 Bulan RP ###################################
$objPHPExcel->getActiveSheet()->setCellValue('D4', floatval($d4));
$objPHPExcel->getActiveSheet()->setCellValue('D5', floatval($d5));
$objPHPExcel->getActiveSheet()->setCellValue('D6', floatval($d6));
$objPHPExcel->getActiveSheet()->setCellValue('D7', floatval($d7));
$objPHPExcel->getActiveSheet()->setCellValue('D8', floatval($d8));
$objPHPExcel->getActiveSheet()->setCellValue('D9', floatval($d9));
$objPHPExcel->getActiveSheet()->setCellValue('D10', floatval($d10));
############################## 1 Bulan VALAS ###################################
$objPHPExcel->getActiveSheet()->setCellValue('E4', floatval($e4));
$objPHPExcel->getActiveSheet()->setCellValue('E5', floatval($e5));
$objPHPExcel->getActiveSheet()->setCellValue('E6', floatval($e6));
$objPHPExcel->getActiveSheet()->setCellValue('E7', floatval($e7));
$objPHPExcel->getActiveSheet()->setCellValue('E8', floatval($e8));
$objPHPExcel->getActiveSheet()->setCellValue('E9', floatval($e9));
$objPHPExcel->getActiveSheet()->setCellValue('E10', floatval($e10));
############################## 3 Bulan VALAS ###################################
$objPHPExcel->getActiveSheet()->setCellValue('F4', floatval($f4));
$objPHPExcel->getActiveSheet()->setCellValue('F5', floatval($f5));
$objPHPExcel->getActiveSheet()->setCellValue('F6', floatval($f6));
$objPHPExcel->getActiveSheet()->setCellValue('F7', floatval($f7));
$objPHPExcel->getActiveSheet()->setCellValue('F8', floatval($f8));
$objPHPExcel->getActiveSheet()->setCellValue('F9', floatval($f9));
$objPHPExcel->getActiveSheet()->setCellValue('F10', floatval($f10));
##########################################################################

#### TABEL KE 2  IDR ####################
## MIN_RATE1
$objPHPExcel->getActiveSheet()->setCellValue('C17', floatval($min_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('C18', floatval($min_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('C19', floatval($min_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('C20', floatval($min_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('C21', floatval($min_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('C22', floatval($min_rate1_idr));
## MAX_RATE1
$objPHPExcel->getActiveSheet()->setCellValue('D17', floatval($max_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('D18', floatval($max_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('D19', floatval($max_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('D20', floatval($max_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('D21', floatval($max_rate1_idr));
$objPHPExcel->getActiveSheet()->setCellValue('D22', floatval($max_rate1_idr));
## MIN_RATE3
$objPHPExcel->getActiveSheet()->setCellValue('F17', floatval($min_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('F18', floatval($min_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('F19', floatval($min_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('F20', floatval($min_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('F21', floatval($min_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('F22', floatval($min_rate3_idr));
## MAX_RATE3
$objPHPExcel->getActiveSheet()->setCellValue('G17', floatval($max_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('G18', floatval($max_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('G19', floatval($max_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('G20', floatval($max_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('G21', floatval($max_rate3_idr));
$objPHPExcel->getActiveSheet()->setCellValue('G22', floatval($max_rate3_idr));
## RATA-RATA 1
$objPHPExcel->getActiveSheet()->setCellValue('E17', "=(C17+D17)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E18', "=(C18+D18)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E19', "=(C19+D19)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E20', "=(C20+D20)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E21', "=(C21+D21)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E22', "=(C22+D22)/2");
## RATA-RATA 2
$objPHPExcel->getActiveSheet()->setCellValue('H17', "=(F17+G17)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H18', "=(F18+G18)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H19', "=(F19+G19)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H20', "=(F20+G20)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H21', "=(F21+G21)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H22', "=(F22+G22)/2");
#### TABEL KE 3  USD ####################
## MIN_RATE1
$objPHPExcel->getActiveSheet()->setCellValue('C29', floatval($min_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('C30', floatval($min_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('C31', floatval($min_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('C32', floatval($min_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('C33', floatval($min_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('C34', floatval($min_rate1_usd));
## MAX_RATE1
$objPHPExcel->getActiveSheet()->setCellValue('D29', floatval($max_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('D30', floatval($max_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('D31', floatval($max_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('D32', floatval($max_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('D33', floatval($max_rate1_usd));
$objPHPExcel->getActiveSheet()->setCellValue('D34', floatval($max_rate1_usd));
## MIN_RATE3
$objPHPExcel->getActiveSheet()->setCellValue('F29', floatval($min_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('F30', floatval($min_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('F31', floatval($min_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('F32', floatval($min_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('F33', floatval($min_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('F34', floatval($min_rate3_usd));
## MAX_RATE3
$objPHPExcel->getActiveSheet()->setCellValue('G29', floatval($max_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('G30', floatval($max_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('G31', floatval($max_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('G32', floatval($max_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('G33', floatval($max_rate3_usd));
$objPHPExcel->getActiveSheet()->setCellValue('G34', floatval($max_rate3_usd));
## RATA-RATA 1
$objPHPExcel->getActiveSheet()->setCellValue('E29', "=(C29+D29)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E30', "=(C30+D30)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E31', "=(C31+D31)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E32', "=(C32+D32)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E33', "=(C33+D33)/2");
$objPHPExcel->getActiveSheet()->setCellValue('E34', "=(C34+D34)/2");
## RATA-RATA 2
$objPHPExcel->getActiveSheet()->setCellValue('H29', "=(F29+G29)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H30', "=(F30+G30)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H31', "=(F31+G31)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H32', "=(F32+G32)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H33', "=(F33+G33)/2");
$objPHPExcel->getActiveSheet()->setCellValue('H34', "=(F34+G34)/2");








$objPHPExcel->getActiveSheet()->setCellValue('B14', 'BANK MNC INTERNASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('C14', 'RUPIAH');
$objPHPExcel->getActiveSheet()->setCellValue('F14', 'RUPIAH');
$objPHPExcel->getActiveSheet()->setCellValue('C15', 'Counter Rate 1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('C16', 'Min');
$objPHPExcel->getActiveSheet()->setCellValue('D16', 'Max');
$objPHPExcel->getActiveSheet()->setCellValue('E15', 'Rata-Rata 1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('F15', 'Counter Rate 3 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('F16', 'Min');
$objPHPExcel->getActiveSheet()->setCellValue('G16', 'Max');
$objPHPExcel->getActiveSheet()->setCellValue('H15', 'Rata-Rata 3 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B17', 'SEMUA NOMINAL');

$objPHPExcel->getActiveSheet()->setCellValue('B26', 'BANK MNC INTERNASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('C26', 'USD');
$objPHPExcel->getActiveSheet()->setCellValue('F26', 'USD');
$objPHPExcel->getActiveSheet()->setCellValue('C27', 'Counter Rate 1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('C28', 'Min');
$objPHPExcel->getActiveSheet()->setCellValue('D28', 'Max');
$objPHPExcel->getActiveSheet()->setCellValue('E27', 'Rata-Rata 1 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('F27', 'Counter Rate 3 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('F28', 'Min');
$objPHPExcel->getActiveSheet()->setCellValue('G28', 'Max');
$objPHPExcel->getActiveSheet()->setCellValue('H27', 'Rata-Rata 3 Bulan');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'SEMUA NOMINAL');



//$objPHPExcel->getActiveSheet()->getStyle('B9:E10')->applyFromArray($styleArrayFont);
//$objPHPExcel->getActiveSheet()->getStyle('A13:K14')->applyFromArray($styleArrayAlignment);
//$objPHPExcel->getActiveSheet()->getStyle("C17:H22")->getNumberFormat()->setFormatCode('#,##0');
//$objPHPExcel->getActiveSheet()->getStyle("C29:H34")->getNumberFormat()->setFormatCode('#,##0');
$objPHPExcel->getActiveSheet()->getStyle('C17:H22')->getNumberFormat()->setFormatCode('0.00');
$objPHPExcel->getActiveSheet()->getStyle('C29:H34')->getNumberFormat()->setFormatCode('0.00');
    
$objPHPExcel->getActiveSheet()->getStyle('C4:F10')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('L8:L23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('P8:P23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
// TITLE 

$objPHPExcel->getActiveSheet()->setTitle('COUNTER RATE');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>
<!--
<b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br></div> </b>
<br>
<br>
<br>
-->
<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> Counter Rate 
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
                                        Counter Rate </a>
                                    </li>
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Counter_Rate_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div></b> 

                                            
</br>
</br>

                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               <tr class="active">
                                                <td width="15%" align="center"><b>Deposito Berjangka </b></td>
                                                <td width="15%" align="center"><b> 1 Bulan </b></td>
                                                <td width="15%" align="center"><b> 3 Bulan </b></td>
                                                <td width="10%" align="center"><b> 1 Bulan </b></td>
                                                <td width="10%" align="center"><b> 3 Bulan </b></td>
                                               
                                                </tr>

                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td width="15%" align="center"><b> 0 < N <= 100 Jt</b></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C4')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D4')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E4')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F4')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>100 Jt < N <= 200 Jt</b></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C5')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D5')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E5')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F5')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>200 Jt < N <= 500 Jt</b></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C6')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D6')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E6')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F6')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>500 Jt < N <= 1 M</b></td>
                                               <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F7')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>1 M < N <= 2 M</b></td>
                                               <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>2 M < N <= 5 M</b></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>> 5 Milyar </b></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('C10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('D10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('E10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" align="center"><?php echo $objPHPExcel->getActiveSheet()->getCell('F10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr>
                                                <td width="15%" align="center"><b>Posisi Tanggal</b></td>
                                                <td width="70%" align="center" colspan="4"><?php echo $objPHPExcel->getActiveSheet()->getCell('C11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                         <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> RUPIAH ( INDONESIA )</b>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" rowspan="2"><b></b></td>
                                                <td width="30%" align="center" colspan="2"><b> Counter Rate 1 Bulan </b></td>
                                                <td width="15%" align="center" rowspan="2"><b> Rata-Rata 1 Bulan </b></td>
                                                <td width="30%" align="center" colspan="2"><b> Counter Rate 3 Bulan </b></td>
                                                <td width="115%" align="center" rowspan="2"><b> Rata-Rata 3 Bulan </b></td>
                                               
                                                </tr>

                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td ><b></b></td>
                                                <td  align="center" ><b>Min</b></td>
                                                <td  align="center" ><b>Max</b></td>
                                                <td ><b></b></td>
                                                <td  align="center" ><b>Min</b></td>
                                                <td  align="center" ><b>Max</b></td>
                                                <td ><b></b></td>
                                                
                                                </tr>
                                                 <tr>
                                                <td align="center" rowspan="7"><b> SEMUA NOMINAL </b></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                          <div class="alert alert-info">
                                        <button class="close" data-close="alert"></button>
                                       <b> USD ( USA ) </b>
                                    </div>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="center" rowspan="2"><b></b></td>
                                                <td width="30%" align="center" colspan="2"><b> Counter Rate 1 Bulan </b></td>
                                                <td width="15%" align="center" rowspan="2"><b> Rata-Rata 1 Bulan </b></td>
                                                <td width="30%" align="center" colspan="2"><b> Counter Rate 3 Bulan </b></td>
                                                <td width="115%" align="center" rowspan="2"><b> Rata-Rata 3 Bulan </b></td>
                                               
                                                </tr>

                                                </thead>
                                                <tbody>

                                                <tr>
                                                <td ><b></b></td>
                                                <td  align="center" ><b>Min</b></td>
                                                <td  align="center" ><b>Max</b></td>
                                                <td ><b></b></td>
                                                <td  align="center" ><b>Min</b></td>
                                                <td  align="center" ><b>Max</b></td>
                                                <td ><b></b></td>
                                                
                                                </tr>
                                                 <tr>
                                                <td align="center" rowspan="7"><b> SEMUA NOMINAL </b></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('C34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('D34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('E34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('F34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('G34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell('H34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>

                                    </div>
                                  
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

