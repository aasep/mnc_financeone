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
$tanggal=$_POST['tanggal']; 
$curr_tgl=date('Y-m-d',strtotime($tanggal));

$label_txtfile=date('Ymd',strtotime($tanggal));
$tanggal_header=date('dmY',strtotime($tanggal));
$tanggal_header2=date('dmY',strtotime(date('Y-m-d',strtotime($tanggal))." 2 day"));;
//$label_tgl_min1=date('d-M-y', strtotime(date('Y-m-d',strtotime($tanggal))." -1 day")); // tanggal terpilih minus (-) 1
$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih

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
*/

/* old 2016-06-22
$query=" SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,a.nominal AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3
WHERE ";
$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY a.kodegl ,b.NOP_Level_3,a.nominal) AS tabel1 ";
*/

# excel

$query=" SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,a.nominal AS Nilai,a.kodecabang FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3
WHERE ";
$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY b.NOP_Level_3,a.kodegl,a.nominal,a.kodecabang)AS tabel1";

/*
$query=" SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,a.kodeproduct,a.kodecabang,a.JenisMataUang,SUM(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3
WHERE "; 

$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY a.kodegl,b.NOP_Level_3,a.kodeproduct,a.JenisMataUang,a.kodecabang) AS tabel1";
*/
//a.DataDate='2016-05-30' AND b.NOP_Level_4 ='NOP2020000053' AND a.JenisMataUang='jpy'




//a.DataDate='2016-05-26' AND b.NOP_Level_3 ='NOP101000001' AND a.JenisMataUang='USD'





/*
$query="SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3
WHERE ";

$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY a.kodegl ,b.NOP_Level_3) AS tabel1 ";
*/

//a.DataDate='2016-05-31' AND b.NOP_Level_3 ='NOP202000005' AND a.JenisMataUang='USD'


#--Mata Uang AUD--
/* old
$query=" SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3 WHERE ";

#-----GLOBAL
$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY a.kodegl ,b.NOP_Level_3 )AS tabel1 ";
*/

#------------Query to Master Modal


$q_modal=" select Nominal_Modal as modal_master from Master_Modal where Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal' ";
$res_modal=odbc_exec($connection2, $q_modal);
$row_modal=odbc_fetch_array($res_modal);
$found_modal=odbc_num_rows($res_modal);
$modal_nilai_fix=$row_modal['modal_master'];


if ($found_modal ==0 || !isset($found_modal)){
    echo "<div class='alert alert-danger alert-dismissable'><b>Data on month $mon_modal - $year_modal is not available. Please select the other date.</b> </div>";
    die();
    }
//echo "$q_modal<br>";
//echo $modal_nilai_fix;
//die();

/*
NOP101000001    Aktiva Tidak Termasuk Giro Pada Bank Lain
NOP101000002    Giro Pada Bank Lain
NOP102000001    Pasiva
NOP201000001    a.Kontrak Pembelian Forward
NOP201000002    b.kontrak Pembelian Future
NOP201000003    c.Kontrak Penjualan Put Option (Bank sebagai Writter)
NOP201000004    d. Kontrak pembelian call options (bank sebagai holder khusus back to back options)
NOP201000005    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward futures dan option
NOP202000001    a.Kontrak Penjualan Forward
NOP202000002    b. Kontrak penjualan futures
NOP202000003    c.  Kontrak penjualan call options (bank sebagai writter)
NOP202000004    d. Kontrak pembelian put options (bank sebagai holder khusus back to back option)
NOP202000005    e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward futures dan option

NOP101010000    Aktiva Tidak Termasuk Giro Pada Bank Lain
NOP101020000    Giro Pada Bank Lain
NOP102010000    Pasiva
NOP201010000    a.Kontrak Pembelian Forward
NOP201020000    b.kontrak Pembelian Future
NOP201030000    c.Kontrak Penjualan Put Option (Bank sebagai Writter)
NOP201040000    d. Kontrak pembelian call options (bank sebagai holder khusus back to back options)
NOP201050000    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward futures dan option
NOP202010000    a.Kontrak Penjualan Forward
NOP202020000    b. Kontrak penjualan futures
NOP202030000    c.  Kontrak penjualan call options (bank sebagai writter)
NOP202040000    d. Kontrak pembelian put options (bank sebagai holder khusus back to back option)
NOP202050000    e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward futures dan option
NOP202050000    e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward futures dan option
NOP202050000    e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward futures dan option
NOP202050000    e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward, futures, dan option
NOP201050000    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option
NOP201050000    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option



                   Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option

*/


#---------------  NOP101010000  Aktiva Valas tidak termasuk giro pada bank lain NOP101000001
$var_nop=" AND b.NOP_Level_3 ='NOP101010000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        //echo $query.$var_tgl.$var_nop.$var_curr1.$var_add."<br><br>";

        $row1=odbc_fetch_array($result1);
        $aktiva_valas_aud=$row1['Jml_Nominal'];
//echo  $query.$var_tgl.$var_nop.$var_curr1.$var_add;
//die();  
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $aktiva_valas_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $aktiva_valas_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $aktiva_valas_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $aktiva_valas_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $aktiva_valas_usd=$row6['Jml_Nominal'];

#---------------  NOP101000002  Giro pada bank lain
$var_nop=" AND b.NOP_Level_3 ='NOP101020000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $giro_aud=$row1['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr1.$var_add;
//die();

$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $giro_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $giro_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $giro_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $giro_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $giro_usd=$row6['Jml_Nominal'];
#---------------  NOP102010000  Pasiva
$var_nop=" AND b.NOP_Level_3 ='NOP102010000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $pasiva_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $pasiva_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $pasiva_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $pasiva_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $pasiva_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $pasiva_usd=$row6['Jml_Nominal'];
######  Rekening Administratif Tagihan Valas #############################

#--------------- NOP201010000 a. Kontrak pembelian forward 
$var_nop=" AND b.NOP_Level_3 ='NOP201010000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $tagihan_a_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $tagihan_a_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $tagihan_a_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $tagihan_a_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $tagihan_a_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $tagihan_a_usd=$row6['Jml_Nominal'];
#--------------- NOP201020000 b. Kontrak pembelian futures
$var_nop=" AND b.NOP_Level_3 ='NOP201020000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $tagihan_b_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $tagihan_b_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $tagihan_b_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $tagihan_b_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $tagihan_b_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $tagihan_b_usd=$row6['Jml_Nominal'];
#--------------- NOP201030000 c. Kontrak penjualan put options (bank sebagai writter)
$var_nop=" AND b.NOP_Level_3 ='NOP201030000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $tagihan_c_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $tagihan_c_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $tagihan_c_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $tagihan_c_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $tagihan_c_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $tagihan_c_usd=$row6['Jml_Nominal'];
#--------------- NOP201040000 d. Kontrak pembelian call options (bank sebagai holder, khusus back to back options)
$var_nop=" AND b.NOP_Level_3 ='NOP201040000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $tagihan_d_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $tagihan_d_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $tagihan_d_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $tagihan_d_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $tagihan_d_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $tagihan_d_usd=$row6['Jml_Nominal'];
#--------------- NOP201050000 e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option
$var_nop=" AND b.NOP_Level_3 ='NOP201050000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $tagihan_e_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $tagihan_e_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $tagihan_e_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $tagihan_e_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $tagihan_e_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $tagihan_e_usd=$row6['Jml_Nominal'];
###### Rekening Administratif Kewajiban Valas #############################

#--------------- NOP202010000 a. Kontrak penjualan forward
$var_nop=" AND b.NOP_Level_3 ='NOP202010000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_a_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_a_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_a_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_a_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_a_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_a_usd=$row6['Jml_Nominal'];

#--------------- NOP202020000 b. Kontrak penjualan futures
$var_nop=" AND b.NOP_Level_3 ='NOP202020000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_b_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_b_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_b_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_b_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_b_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_b_usd=$row6['Jml_Nominal'];
#--------------- NOP202030000 c.  Kontrak penjualan call options (bank sebagai writter)
$var_nop=" AND b.NOP_Level_3 ='NOP202030000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_c_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_c_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_c_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_c_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_c_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_c_usd=$row6['Jml_Nominal'];
#--------------- NOP202040000 d. Kontrak pembelian put options (bank sebagai holder, khusus back to back option)
$var_nop=" AND b.NOP_Level_3 ='NOP202040000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_d_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_d_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_d_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_d_jpy=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_d_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_d_usd=$row6['Jml_Nominal'];



#--------------- NOP202000005 e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward, futures, dan option
 /*       
$query ="SELECT SUM(Nilai) AS Jml_Nominal FROM( SELECT a.kodegl,a.kodeproduct,a.kodecabang,a.JenisMataUang,SUM(a.nominal) AS Nilai 
FROM DM_Journal a WITH (NOLOCK) 
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct 
JOIN Referensi_NOP c ON c.NOP_Level_4 = b.NOP_Level_4 WHERE ";


$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add=" GROUP BY a.kodegl,b.NOP_Level_4,a.kodeproduct,a.JenisMataUang,a.kodecabang) AS tabel1";
*/

$var_nop=" AND b.NOP_Level_3 ='NOP202050000' ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy=$row4['Jml_Nominal'];



//        echo $query.$var_tgl.$var_nop.$var_curr4.$var_add;
//        die();
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd=$row6['Jml_Nominal'];

#NOP2020000051   Rekening Administratif
#NOP2020000053   Kontrak Penjualan Forward
#NOP2020000052   Transaksi Derivatif diluar kontrak Penjualan Forward,Futures dan Option
///======================= A==============================

/* comment 2016-09-01
$query ="SELECT SUM(Nilai) AS Jml_Nominal FROM( SELECT a.kodegl,a.kodeproduct,a.kodecabang,a.JenisMataUang,SUM(a.nominal) AS Nilai 
FROM DM_Journal a WITH (NOLOCK) 
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct 
JOIN Referensi_NOP c ON c.NOP_Level_4 = b.NOP_Level_4 WHERE ";
*/

/*
$query=" SELECT CONCAT(kode_PDN,MataUang,SUM(Nilai))AS Hasil FROM(
SELECT a.kodegl,a.JenisMataUang AS MataUang,SUM(a.nominal)AS Nilai,d.PDN_CODE AS kode_PDN FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON c.NOP_Level_3 = b.NOP_Level_3 AND c.NOP_Level_4 = b.NOP_Level_4
JOIN PARAMETER_PDN_CODE d ON d.NOP_Level_4 = b.NOP_Level_4 
WHERE ";
*/
$query=" SELECT SUM(Nilai) AS Jml_Nominal FROM(
SELECT a.kodegl,a.JenisMataUang AS MataUang,SUM(a.nominal)AS Nilai,d.PDN_CODE AS kode_PDN FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_NOP c ON  c.NOP_Level_4 = b.NOP_Level_4
JOIN PARAMETER_PDN_CODE d ON d.NOP_Level_4 = b.NOP_Level_4 
WHERE ";


//WHERE a.DataDate='2016-05-31' AND b.NOP_Level_3 ='NOP101000001' AND d.NOP_Level_4='NOP101000001' AND a.JenisMataUang='IDR' ;



$var_tgl=" a.DataDate='$curr_tgl' ";
$var_add="GROUP BY a.kodegl ,b.NOP_Level_3,a.JenisMataUang,d.NOP_Level_4,d.PDN_CODE)AS tabel1 group by MataUang,kode_PDN ";









//a.DataDate='2016-05-30' AND b.NOP_Level_4 ='NOP2020000051' AND a.JenisMataUang='JPY' 


/*
NOP2020000051
NOP2020000052
NOP2020000053
NOP2020000054
NOP2010000055
NOP2010000056

*/



/*
NOP101010010    10
NOP101020010    15
NOP102010010    29
NOP201050010    31
NOP201010010    32
NOP201020010    33
NOP201050030    34
NOP202050010    35
NOP202010010    39
NOP202020010    51
NOP202050020    52
NULL    55
NULL    59
NOP201030010    61
NOP202040010    65
NOP202030010    67
NOP201040010    69
NULL    71


===================================
NOP101010000    NOP101010010    10
NOP101020000    NOP101020010    15
NOP102010000    NOP102010010    29
NOP201010000    NOP201010010    32
NOP201020000    NOP201020010    33
NOP201030000    NOP201030010    61
NOP201040000    NOP201040010    69
NOP201050000    NOP201050010    31
NOP202010000    NOP202010010    39
NOP202020000    NOP202020010    51
NOP202030000    NOP202030010    67
NOP202040000    NOP202040010    65
NOP202050000    NOP202050010    35
NOP202050000    NOP202050020    52
NOP202050000    NOP202050030
NOP202050000    NOP202050040
NOP201050000    NOP201050020
NOP201050000    NOP201050030    34




====== benerin============
NOP201050010   NOP202050040 31
NOP201010010   NOP201050020 32
NOP202010010   NOP202050030 39



NOP101010000    NOP101010010
NOP101020000    NOP101020010
NOP102010000    NOP102010010
NOP201010000    NOP201010010
NOP201020000    NOP201020010
NOP201030000    NOP201030010
NOP201040000    NOP201040010
NOP201050000    NOP201050010
NOP202010000    NOP202010010
NOP202020000    NOP202020010
NOP202030000    NOP202030010
NOP202040000    NOP202040010
NOP202050000    NOP202050010
NOP202050000    NOP202050020
NOP202050000    NOP202050030
NOP202050000    NOP202050040
NOP201050000    NOP201050020
NOP201050000    NOP201050030




NOP101010010    Aktiva Tidak Termasuk Giro Pada Bank Lain    10  Aktiva valas tidak termasuk giro pada bank lain
*/
        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_idr=$row1['Jml_Nominal'];



//echo $query.$var_tgl.$var_nop.$var_add;
//die();
        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_aud=$row1['Jml_Nominal'];


        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_eur=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_hkd=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_jpy=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_sgd=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101010010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_10_usd=$row1['Jml_Nominal'];

//   NOP101020010    Giro Pada Bank Lain    15  Aktiva valas giro pada bank lain

        $var_nop="  AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_idr=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_aud=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_eur=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_hkd=$row1['Jml_Nominal'];

        $var_nop="  AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP101020010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_15_usd=$row1['Jml_Nominal'];

//   NOP102010010    Pasiva      29  Pasiva valas

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_aud=$row1['Jml_Nominal'];


        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP102010010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_29_usd=$row1['Jml_Nominal'];

//  NOP201010000    NOP201010010    32   new   NOP201050020
        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050020' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_32_usd=$row1['Jml_Nominal'];

//  NOP201020000    NOP201020010    33

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_aud=$row1['Jml_Nominal'];


        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201020010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_33_usd=$row1['Jml_Nominal'];


//  NOP201030000    NOP201030010    61

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_aud=$row1['Jml_Nominal'];


        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201030010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_61_usd=$row1['Jml_Nominal'];

//  NOP201040000    NOP201040010    69

         $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_aud=$row1['Jml_Nominal'];


        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201040010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_69_usd=$row1['Jml_Nominal'];




//  NOP201050000    NOP201050010    31 new  NOP202050040 
        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050040' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_31_usd=$row1['Jml_Nominal'];



//  NOP202010000    NOP202010010    39   new  NOP202050030

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050030' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_39_usd=$row1['Jml_Nominal'];

//  NOP202020000    NOP202020010    51
        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202020010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_51_usd=$row1['Jml_Nominal'];


//  NOP202030000    NOP202030010    67
        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202030010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_67_usd=$row1['Jml_Nominal'];

//  NOP202040000    NOP202040010    65

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202040010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_65_usd=$row1['Jml_Nominal'];

//  NOP202050000    NOP202050010    35

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050010' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_35_usd=$row1['Jml_Nominal'];
//  NOP202050000    NOP202050020    52
        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP202050020' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_52_usd=$row1['Jml_Nominal'];

//  NOP201050000    NOP201050030    34
        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='IDR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_idr=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='AUD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_aud=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='EUR'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_eur=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='HKD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_hkd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='JPY'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_jpy=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='SGD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_sgd=$row1['Jml_Nominal'];

        $var_nop="   AND d.NOP_Level_4='NOP201050030' AND a.JenisMataUang='USD'  ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_add);
        $row1=odbc_fetch_array($result1);
        $code_34_usd=$row1['Jml_Nominal'];

/*

NOP201010010    a.Kontrak Pembelian Forward                                                 32  Kontrak Pembelian Forward
NOP201020010    b.kontrak Pembelian Future                                                  33  kontrak pembelian Futures
NOP201030010    c.Kontrak Penjualan Put Option (Bank sebagai Writter)/Tagihan valas         61  Kontrak Penjualan Put Option (bank sebagai writer)
NOP201040010    d. Kontrak pembelian call options (bank sebagai holder, khusus back to back options)/Tagihan valas  69  Kontrak Pembelian Call Option (bank sebagai holder, khusus option yang identik)
NOP201050010    Rekening Tagihan Administratif                                              31  Rekening administratif-Tagihan Valas
NOP201050030    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option   34  Transaksi Derivatif diluar kontrak Pembelian Forward, Futures dan Option
NOP202010010    Kontrak Penjualan Forward                                                   39  Kontrak Penjualan Forward
NOP202020010    b. Kontrak penjualan futures                                                51  Kontrak Penjualan Futures
NOP202030010    c.  Kontrak penjualan call options (bank sebagai writter)/Kewajiban valas   67  Kontrak Penjualan Call Option (bank sebagai writer)
NOP202040010    d. Kontrak pembelian put options (bank sebagai holder, khusus back to back option)/Kewajiban valas  65  Kontrak Pembelian Put Option (bank sebagai holder, khusus option yang identik)
NOP202050010    Rekening Administratif                                                      35  Rekening Administratif - Kewajiban Valas
NOP202050020    Transaksi Derivatif diluar kontrak Penjualan Forward,Futures dan Option     52  Transaksi Derivatif diluar kontrak Penjualan Forward, Futures dan Option
*/
//============================0051
        $var_nop=" AND b.NOP_Level_4 ='NOP2020000051'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_a=$row1['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr1.$var_add;
//die();


$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_a=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_a=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_a=$row4['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr4.$var_add;
//die(); 
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_a=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_a=$row6['Jml_Nominal'];



//======================== B =============================0052
        $var_nop=" AND b.NOP_Level_4 ='NOP2020000052'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_b=$row1['Jml_Nominal'];

    
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_b=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_b=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_b=$row4['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr4.$var_add;
//die();

$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_b=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_b=$row6['Jml_Nominal'];
///======================= C =============================0053
        $var_nop=" AND b.NOP_Level_4 ='NOP2020000053'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_c=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_c=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_c=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_c=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_c=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_c=$row6['Jml_Nominal'];

###################################

        // D ============================0054
        $var_nop=" AND b.NOP_Level_4 ='NOP2020000054'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_d=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_d=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_d=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_d=$row4['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr4.$var_add;
//die(); 
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_d=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_d=$row6['Jml_Nominal'];



//======================== E =============================0055
        $var_nop=" AND b.NOP_Level_4 ='NOP2010000055'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_e=$row1['Jml_Nominal'];

    
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_e=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_e=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_e=$row4['Jml_Nominal'];
//echo $query.$var_tgl.$var_nop.$var_curr4.$var_add;
//die();

$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_e=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_e=$row6['Jml_Nominal'];
///======================= F =============================0053
        $var_nop=" AND b.NOP_Level_4 ='NOP2010000056'  ";
$var_curr1=" AND a.JenisMataUang='AUD' ";
        $result1=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr1.$var_add);
        $row1=odbc_fetch_array($result1);
        $kewajiban_e_aud_f=$row1['Jml_Nominal'];
$var_curr2=" AND a.JenisMataUang='EUR' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr2.$var_add);
        $row2=odbc_fetch_array($result2);
        $kewajiban_e_eur_f=$row2['Jml_Nominal'];
$var_curr3=" AND a.JenisMataUang='HKD' ";
        $result3=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr3.$var_add);
        $row3=odbc_fetch_array($result3);
        $kewajiban_e_hkd_f=$row3['Jml_Nominal'];
$var_curr4=" AND a.JenisMataUang='JPY' ";
        $result4=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr4.$var_add);
        $row4=odbc_fetch_array($result4);
        $kewajiban_e_jpy_f=$row4['Jml_Nominal'];
$var_curr5=" AND a.JenisMataUang='SGD' ";
        $result5=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr5.$var_add);
        $row5=odbc_fetch_array($result5);
        $kewajiban_e_sgd_f=$row5['Jml_Nominal'];
$var_curr6=" AND a.JenisMataUang='USD' ";
        $result6=odbc_exec($connection2, $query.$var_tgl.$var_nop.$var_curr6.$var_add);
        $row6=odbc_fetch_array($result6);
        $kewajiban_e_usd_f=$row6['Jml_Nominal'];


#############  KONDISI GIRO PADA BANK LAIN  NEGATIVE ( - ) ##############################
/*
$objPHPExcel->getActiveSheet()->setCellValue('E18', $giro_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F18', $giro_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G18', $giro_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H18', $giro_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $giro_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J18', $giro_usd);

$objPHPExcel->getActiveSheet()->setCellValue('E20', -1*$pasiva_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F20', -1*$pasiva_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G20', -1*$pasiva_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H20', -1*$pasiva_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I20', -1*$pasiva_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J20', -1*$pasiva_usd);




*/

/*$query_rak=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' 
group by a.KodeGL
)as tabel1 ";
$res_rak=odbc_exec($connection2, $query_rak);
$row_rak=odbc_fetch_array($res_rak);
$found_rak=$row_rak['total'];
*/
$q_rak_usd=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='USD'
group by a.KodeGL
)as tabel1 ";
$res_rak_usd=odbc_exec($connection2, $q_rak_usd);
$row_rak_usd=odbc_fetch_array($res_rak_usd);
$num_rak_usd=$row_rak_usd['total'];
if ($num_rak_usd >= 0)
{  
$aktiva_valas_usd=$aktiva_valas_usd+abs($num_rak_usd);
} else {
$pasiva_usd=$pasiva_usd+abs($num_rak_usd);
}
######## AUD
$q_rak_aud=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='AUD'
group by a.KodeGL
)as tabel1 ";
$res_rak_aud=odbc_exec($connection2, $q_rak_aud);
$row_rak_aud=odbc_fetch_array($res_rak_aud);
$num_rak_aud=$row_rak_aud['total'];
if ($num_rak_aud >= 0)
{  
$aktiva_valas_aud=$aktiva_valas_aud+abs($num_rak_aud);
} else {
$pasiva_aud=$pasiva_aud+abs($num_rak_aud);
}
############ EUR
$q_rak_eur=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='EUR'
group by a.KodeGL
)as tabel1 ";
$res_rak_eur=odbc_exec($connection2, $q_rak_eur);
$row_rak_eur=odbc_fetch_array($res_rak_eur);
$num_rak_eur=$row_rak_eur['total'];
if ($num_rak_eur >= 0)
{  
$aktiva_valas_eur=$aktiva_valas_eur+abs($num_rak_eur);
} else {
$pasiva_eur=$pasiva_eur+abs($num_rak_eur);
}

############ JPY
$q_rak_jpy=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='JPY'
group by a.KodeGL
)as tabel1 ";
$res_rak_jpy=odbc_exec($connection2, $q_rak_jpy);
$row_rak_jpy=odbc_fetch_array($res_rak_jpy);
$num_rak_jpy=$row_rak_jpy['total'];
if ($num_rak_jpy >= 0)
{  
$aktiva_valas_jpy=$aktiva_valas_jpy+abs($num_rak_jpy);
} else {
$pasiva_jpy=$pasiva_jpy+abs($num_rak_jpy);
}

############ HKD
$q_rak_hkd=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='HKD'
group by a.KodeGL
)as tabel1 ";
$res_rak_hkd=odbc_exec($connection2, $q_rak_hkd);
$row_rak_hkd=odbc_fetch_array($res_rak_hkd);
$num_rak_hkd=$row_rak_hkd['total'];
if ($num_rak_hkd >= 0)
{  
$aktiva_valas_hkd=$aktiva_valas_hkd+abs($num_rak_hkd);
} else {
$pasiva_hkd=$pasiva_hkd+abs($num_rak_hkd);
}


############ SGD
$q_rak_sgd=" select sum(total) as total from (
SELECT sum (nominal) as total,a.kodegl
FROM DM_JOURNAL a
JOIN Referensi_GL_02 b ON a.KodeGL=b.GLNO AND a. KodeProduct =b.PRODNO
WHERE DataDate='$curr_tgl' and b.RAK_Flag='y' and a.JenisMataUang='SGD'
group by a.KodeGL
)as tabel1 ";
$res_rak_sgd=odbc_exec($connection2, $q_rak_sgd);
$row_rak_sgd=odbc_fetch_array($res_rak_sgd);
$num_rak_sgd=$row_rak_sgd['total'];
if ($num_rak_sgd >= 0)
{  
$aktiva_valas_sgd=$aktiva_valas_sgd+abs($num_rak_sgd);
} else {
$pasiva_sgd=$pasiva_sgd+abs($num_rak_sgd);
}



//echo $found_rak."<br><br>";
//echo $num_rak_usd;







//--and a.JenisMataUang='USD'



##---------- AUD---------------------
if ( $giro_aud < 0 ){

$pasiva_aud=(-1*$pasiva_aud) + (-1*$giro_aud);
$giro_aud=0;

} else {

$pasiva_aud=(-1*$pasiva_aud);

}
##---------- EUR ---------------------
if ( $giro_eur < 0 ){

$pasiva_eur=(-1*$pasiva_eur) + (-1*$giro_eur);
$giro_eur=0;

} else {

$pasiva_eur=(-1*$pasiva_eur);

}

##---------- HKD ---------------------
if ( $giro_hkd < 0 ){

$pasiva_hkd=(-1*$pasiva_hkd) + (-1*$giro_hkd);
$giro_hkd=0;


} else {

$pasiva_hkd=(-1*$pasiva_hkd);

}

##---------- JPY ---------------------
if ( $giro_jpy < 0 ){

$pasiva_jpy=(-1*$pasiva_jpy) + (-1*$giro_jpy);
$giro_jpy=0;


} else {

$pasiva_jpy=(-1*$pasiva_jpy);

}

##---------- SGD ---------------------
if ( $giro_sgd < 0 ){

$pasiva_sgd=(-1*$pasiva_sgd) + (-1*$giro_sgd);

$giro_sgd=0;
} else {

$pasiva_sgd=(-1*$pasiva_sgd);

}##---------- USD ---------------------
if ( $giro_usd < 0 ){

$pasiva_usd=(-1*$pasiva_usd) + (-1*$giro_usd);
$giro_usd=0;

} else {

$pasiva_usd=(-1*$pasiva_usd);

}


#########################################################################################





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
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(12);
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'MASA LAPORAN');
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'BANK : PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('K12', '(Jutaan Rupiah)');
$objPHPExcel->getActiveSheet()->setCellValue('E9', $label_tgl);
$objPHPExcel->getActiveSheet()->getStyle('B9:E10')->applyFromArray($styleArrayFont);
$objPHPExcel->getActiveSheet()->getStyle('A13:K14')->applyFromArray($styleArrayAlignment);

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
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
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'A.');
$objPHPExcel->getActiveSheet()->setCellValue('D15', 'Neraca');
$objPHPExcel->getActiveSheet()->getStyle('B15:Z15')->applyFromArray($styleArrayFont);
//1. Aktiva Valas
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C16', '1');
$objPHPExcel->getActiveSheet()->setCellValue('D16', 'Aktiva Valas');
$objPHPExcel->getActiveSheet()->getStyle('B16:Z16')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D17', '- Aktiva Valas tidak termasuk giro pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('E17', $angka);
$objPHPExcel->getActiveSheet()->setCellValue('D18', '- Giro pada bank lain');
$objPHPExcel->getActiveSheet()->getStyle('A17:Z18')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C20', '2');
$objPHPExcel->getActiveSheet()->setCellValue('D20', 'Pasiva Valas');

$objPHPExcel->getActiveSheet()->setCellValue('C21', '3');
$objPHPExcel->getActiveSheet()->setCellValue('D21', 'Selisih Aktiva dan Pasiva Valas (A.1 - A.2)');
$objPHPExcel->getActiveSheet()->getStyle('A20:Z21')->applyFromArray($styleArrayFont);

$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D22', 'Selisih Aktiva dan Pasiva Valas (Nilai Absolut)');
$objPHPExcel->getActiveSheet()->getStyle('A22:Z22')->applyFromArray($styleArrayFont);

//B. REKENING ADMINISTRATIF
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'B.');
$objPHPExcel->getActiveSheet()->setCellValue('D24', 'Rekening Administratif');
$objPHPExcel->getActiveSheet()->setCellValue('C25', '1');
$objPHPExcel->getActiveSheet()->setCellValue('D25', 'Rekening Administratif Tagihan Valas');
$objPHPExcel->getActiveSheet()->getStyle('A24:Z25')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D26', 'a. Kontrak pembelian forward');
$objPHPExcel->getActiveSheet()->setCellValue('D27', 'b. Kontrak pembelian futures');
$objPHPExcel->getActiveSheet()->setCellValue('D28', 'c. Kontrak penjualan put options (bank sebagai writter)');
$objPHPExcel->getActiveSheet()->setCellValue('D29', 'd. Kontrak pembelian call options (bank sebagai');
$objPHPExcel->getActiveSheet()->setCellValue('D30', '   holder, khusus back to back options)');
$objPHPExcel->getActiveSheet()->setCellValue('D31', 'e. Rekening Administratif Tagihan Valas diluar ');
$objPHPExcel->getActiveSheet()->setCellValue('D32', '    kontrak pemberian forward, futures, dan option');
$objPHPExcel->getActiveSheet()->getStyle('A26:D32')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C34', '2');
$objPHPExcel->getActiveSheet()->setCellValue('D34', 'Rekening Administratif Kewajiban Valas');
$objPHPExcel->getActiveSheet()->getStyle('A34:Z34')->applyFromArray($styleArrayFont);
$styleArrayFont = array('font' => array(''  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('D35', 'a. Kontrak penjualan forward');
$objPHPExcel->getActiveSheet()->setCellValue('D36', 'b. Kontrak penjualan futures');
$objPHPExcel->getActiveSheet()->setCellValue('D37', 'c.  Kontrak penjualan call options (bank sebagai writter)');
$objPHPExcel->getActiveSheet()->setCellValue('D38', 'd. Kontrak pembelian put options (bank sebagai');
$objPHPExcel->getActiveSheet()->setCellValue('D39', '   holder, khusus back to back option)');
$objPHPExcel->getActiveSheet()->setCellValue('D40', 'e. Rekening Administratif Kewajiban Valas diluar');
$objPHPExcel->getActiveSheet()->setCellValue('D41', '	kontrak penjualan forward, futures, dan option
');
$objPHPExcel->getActiveSheet()->getStyle('A35:D41')->applyFromArray($styleArrayFont);

$objPHPExcel->getActiveSheet()->setCellValue('E16', "=SUM(E17:E18)");
$objPHPExcel->getActiveSheet()->setCellValue('F16', "=SUM(F17:F18)");
$objPHPExcel->getActiveSheet()->setCellValue('G16', "=SUM(G17:G18)");
$objPHPExcel->getActiveSheet()->setCellValue('H16', "=SUM(H17:H18)");
$objPHPExcel->getActiveSheet()->setCellValue('I16', "=SUM(I17:I18)");
$objPHPExcel->getActiveSheet()->setCellValue('J16', "=SUM(J17:J18)");
$objPHPExcel->getActiveSheet()->setCellValue('K16', "=SUM(E16:J16)");

$objPHPExcel->getActiveSheet()->setCellValue('E21', "=(E16-E20)");
$objPHPExcel->getActiveSheet()->setCellValue('F21', "=(F16-F20)");
$objPHPExcel->getActiveSheet()->setCellValue('G21', "=(G16-G20)");
$objPHPExcel->getActiveSheet()->setCellValue('H21', "=(H16-H20)");
$objPHPExcel->getActiveSheet()->setCellValue('I21', "=(I16-I20)");
$objPHPExcel->getActiveSheet()->setCellValue('J21', "=(J16-J20)");
$objPHPExcel->getActiveSheet()->setCellValue('K21', "=SUM(E21:J21)");

$objPHPExcel->getActiveSheet()->setCellValue('E22', "=ABS(E21)");
$objPHPExcel->getActiveSheet()->setCellValue('F22', "=ABS(F21)");
$objPHPExcel->getActiveSheet()->setCellValue('G22', "=ABS(G21)");
$objPHPExcel->getActiveSheet()->setCellValue('H22', "=ABS(H21)");
$objPHPExcel->getActiveSheet()->setCellValue('I22', "=ABS(I21)");
$objPHPExcel->getActiveSheet()->setCellValue('J22', "=ABS(J21)");
$objPHPExcel->getActiveSheet()->setCellValue('K22', "=SUM(E22:J22)");

$objPHPExcel->getActiveSheet()->setCellValue('E25', "=SUM(E26:E32)");
$objPHPExcel->getActiveSheet()->setCellValue('F25', "=SUM(F26:F32)");
$objPHPExcel->getActiveSheet()->setCellValue('G25', "=SUM(G26:G32)");
$objPHPExcel->getActiveSheet()->setCellValue('H25', "=SUM(H26:H32)");
$objPHPExcel->getActiveSheet()->setCellValue('I25', "=SUM(I26:I32)");
$objPHPExcel->getActiveSheet()->setCellValue('J25', "=SUM(J26:J32)");
$objPHPExcel->getActiveSheet()->setCellValue('K25', "=SUM(E25:J25)");


$objPHPExcel->getActiveSheet()->setCellValue('E34', "=SUM(E35:E41)");
$objPHPExcel->getActiveSheet()->setCellValue('F34', "=SUM(F35:F41)");
$objPHPExcel->getActiveSheet()->setCellValue('G34', "=SUM(G35:G41)");
$objPHPExcel->getActiveSheet()->setCellValue('H34', "=SUM(H35:H41)");
$objPHPExcel->getActiveSheet()->setCellValue('I34', "=SUM(I35:I41)");
$objPHPExcel->getActiveSheet()->setCellValue('J34', "=SUM(J35:J41)");
$objPHPExcel->getActiveSheet()->setCellValue('K34', "=SUM(E34:J34)");

$objPHPExcel->getActiveSheet()->setCellValue('E17', $aktiva_valas_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F17', $aktiva_valas_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G17', $aktiva_valas_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H17', $aktiva_valas_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I17', $aktiva_valas_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J17', $aktiva_valas_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K17', "=SUM(E17:J17)");

/*
#### bayangan 
$objPHPExcel->getActiveSheet()->setCellValue('M17', $aktiva_valas_aud);
$objPHPExcel->getActiveSheet()->setCellValue('N17', $aktiva_valas_eur);
$objPHPExcel->getActiveSheet()->setCellValue('O17', $aktiva_valas_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('P17', $aktiva_valas_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('Q17', $aktiva_valas_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('R17', $aktiva_valas_usd);

$objPHPExcel->getActiveSheet()->setCellValue('M18', floatval($num_rak_aud));
$objPHPExcel->getActiveSheet()->setCellValue('N18', floatval($num_rak_eur));
$objPHPExcel->getActiveSheet()->setCellValue('O18', floatval($num_rak_hkd));
$objPHPExcel->getActiveSheet()->setCellValue('P18', floatval($num_rak_jpy));
$objPHPExcel->getActiveSheet()->setCellValue('Q18', floatval($num_rak_sgd));
$objPHPExcel->getActiveSheet()->setCellValue('R18', floatval($num_rak_usd));
*/

//$objPHPExcel->getActiveSheet()->setCellValue('M17', floatval($aktiva_valas_usd));


$objPHPExcel->getActiveSheet()->setCellValue('E18', $giro_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F18', $giro_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G18', $giro_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H18', $giro_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $giro_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J18', $giro_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K18', "=SUM(E18:J18)");



$objPHPExcel->getActiveSheet()->setCellValue('E20', $pasiva_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F20', $pasiva_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G20', $pasiva_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H20', $pasiva_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I20', $pasiva_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J20', $pasiva_usd);


/*
#### bayangan
$objPHPExcel->getActiveSheet()->setCellValue('M20', $pasiva_aud);
$objPHPExcel->getActiveSheet()->setCellValue('N20', $pasiva_eur);
$objPHPExcel->getActiveSheet()->setCellValue('O20', $pasiva_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('P20', $pasiva_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('Q20', $pasiva_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('R20', $pasiva_usd);
*/

$objPHPExcel->getActiveSheet()->setCellValue('K20', "=SUM(E20:J20)");


#------------------Output Rekening Administratif Tagihan Valas
#a. Kontrak pembelian forward

$objPHPExcel->getActiveSheet()->setCellValue('E26', $tagihan_a_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F26', $tagihan_a_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G26', $tagihan_a_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H26', $tagihan_a_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I26', $tagihan_a_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J26', $tagihan_a_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K26', "=SUM(E26:J26)");

$objPHPExcel->getActiveSheet()->setCellValue('E27', $tagihan_b_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F27', $tagihan_b_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G27', $tagihan_b_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H27', $tagihan_b_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I27', $tagihan_b_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J27', $tagihan_b_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K27', "=SUM(E27:J27)");

$objPHPExcel->getActiveSheet()->setCellValue('E28', $tagihan_c_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F28', $tagihan_c_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G28', $tagihan_c_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H28', $tagihan_c_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I28', $tagihan_c_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J28', $tagihan_c_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K28', "=SUM(E28:J28)");

$objPHPExcel->getActiveSheet()->setCellValue('E29', $tagihan_d_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F29', $tagihan_d_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G29', $tagihan_d_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H29', $tagihan_d_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I29', $tagihan_d_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J29', $tagihan_d_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K29', "=SUM(E29:J29)");

$objPHPExcel->getActiveSheet()->setCellValue('E32', $tagihan_e_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F32', $tagihan_e_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G32', $tagihan_e_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H32', $tagihan_e_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I32', $tagihan_e_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J32', $tagihan_e_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K32', "=SUM(E32:J32)");



#------------------Rekening Administratif Kewajiban Valas

$objPHPExcel->getActiveSheet()->setCellValue('E35', (-1)*$kewajiban_a_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F35', (-1)*$kewajiban_a_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G35', (-1)*$kewajiban_a_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H35', (-1)*$kewajiban_a_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I35', (-1)*$kewajiban_a_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J35', (-1)*$kewajiban_a_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K35', "=SUM(E35:J35)");

$objPHPExcel->getActiveSheet()->setCellValue('E36', (-1)*$kewajiban_b_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F36', (-1)*$kewajiban_b_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G36', (-1)*$kewajiban_b_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H36', (-1)*$kewajiban_b_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I36', (-1)*$kewajiban_b_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J36', (-1)*$kewajiban_b_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K36', "=SUM(E36:J36)");

$objPHPExcel->getActiveSheet()->setCellValue('E37', (-1)*$kewajiban_c_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F37', (-1)*$kewajiban_c_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G37', (-1)*$kewajiban_c_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H37', (-1)*$kewajiban_c_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I37', (-1)*$kewajiban_c_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J37', (-1)*$kewajiban_c_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K37', "=SUM(E37:J37)");

$objPHPExcel->getActiveSheet()->setCellValue('E38', (-1)*$kewajiban_d_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F38', (-1)*$kewajiban_d_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G38', (-1)*$kewajiban_d_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H38', (-1)*$kewajiban_d_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I38', (-1)*$kewajiban_d_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J38', (-1)*$kewajiban_d_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K38', "=SUM(E38:J38)");

$objPHPExcel->getActiveSheet()->setCellValue('E41', (-1)*$kewajiban_e_aud);
$objPHPExcel->getActiveSheet()->setCellValue('F41', (-1)*$kewajiban_e_eur);
$objPHPExcel->getActiveSheet()->setCellValue('G41', (-1)*$kewajiban_e_hkd);
$objPHPExcel->getActiveSheet()->setCellValue('H41', (-1)*$kewajiban_e_jpy);
$objPHPExcel->getActiveSheet()->setCellValue('I41', (-1)*$kewajiban_e_sgd);
$objPHPExcel->getActiveSheet()->setCellValue('J41', (-1)*$kewajiban_e_usd);
$objPHPExcel->getActiveSheet()->setCellValue('K41', "=SUM(E41:J41)");

#NEW
$objPHPExcel->getActiveSheet()->setCellValue('E43', "=(E25-E34)");
$objPHPExcel->getActiveSheet()->setCellValue('F43', "=(F25-F34)");
$objPHPExcel->getActiveSheet()->setCellValue('G43', "=(G25-G34)");
$objPHPExcel->getActiveSheet()->setCellValue('H43', "=(H25-H34)");
$objPHPExcel->getActiveSheet()->setCellValue('I43', "=(I25-I34)");
$objPHPExcel->getActiveSheet()->setCellValue('J43', "=(J25-J34)");
$objPHPExcel->getActiveSheet()->setCellValue('K43', "=SUM(E43:J43)");


$objPHPExcel->getActiveSheet()->setCellValue('E46', "=(E21+E43)");
$objPHPExcel->getActiveSheet()->setCellValue('F46', "=(F21+F43)");
$objPHPExcel->getActiveSheet()->setCellValue('G46', "=(G21+G43)");
$objPHPExcel->getActiveSheet()->setCellValue('H46', "=(H21+H43)");
$objPHPExcel->getActiveSheet()->setCellValue('I46', "=(I21+I43)");
$objPHPExcel->getActiveSheet()->setCellValue('J46', "=(J21+J43)");
$objPHPExcel->getActiveSheet()->setCellValue('K46', "=SUM(E46:J46)");

$objPHPExcel->getActiveSheet()->setCellValue('E49', "=ABS(E46)");
$objPHPExcel->getActiveSheet()->setCellValue('F49', "=ABS(F46)");
$objPHPExcel->getActiveSheet()->setCellValue('G49', "=ABS(G46)");
$objPHPExcel->getActiveSheet()->setCellValue('H49', "=ABS(H46)");
$objPHPExcel->getActiveSheet()->setCellValue('I49', "=ABS(I46)");
$objPHPExcel->getActiveSheet()->setCellValue('J49', "=ABS(J46)");
$objPHPExcel->getActiveSheet()->setCellValue('K49', "=SUM(E49:J49)");

$objPHPExcel->getActiveSheet()->setCellValue('E51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('F51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('G51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('H51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('I51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('J51', $modal_nilai_fix);
$objPHPExcel->getActiveSheet()->setCellValue('K51', $modal_nilai_fix);
//E21/E51 dijadikan %  ===>53


$objPHPExcel->getActiveSheet()->setCellValue('E53', "=ABS(+IF(E21=0,0,(E21/E51)))");
$objPHPExcel->getActiveSheet()->setCellValue('F53', "=ABS(+IF(F21=0,0,(F21/F51)))");
$objPHPExcel->getActiveSheet()->setCellValue('G53', "=ABS(+IF(G21=0,0,(G21/G51)))");
$objPHPExcel->getActiveSheet()->setCellValue('H53', "=ABS(+IF(H21=0,0,(H21/H51)))");
$objPHPExcel->getActiveSheet()->setCellValue('I53', "=ABS(+IF(I21=0,0,(I21/I51)))");
$objPHPExcel->getActiveSheet()->setCellValue('J53', "=ABS(+IF(J21=0,0,(J21/J51)))");
$objPHPExcel->getActiveSheet()->setCellValue('K53', "=SUM(E53:J53)");
//E49/E51 dijadikan %  ===>55
$objPHPExcel->getActiveSheet()->setCellValue('E55', "=ABS(+IF(E49=0,0,(E49/E51)))");
$objPHPExcel->getActiveSheet()->setCellValue('F55', "=ABS(+IF(F49=0,0,(F49/F51)))");
$objPHPExcel->getActiveSheet()->setCellValue('G55', "=ABS(+IF(G49=0,0,(G49/G51)))");
$objPHPExcel->getActiveSheet()->setCellValue('H55', "=ABS(+IF(H49=0,0,(H49/H51)))");
$objPHPExcel->getActiveSheet()->setCellValue('I55', "=ABS(+IF(I49=0,0,(I49/I51)))");
$objPHPExcel->getActiveSheet()->setCellValue('J55', "=ABS(+IF(J49=0,0,(J49/J51)))");
$objPHPExcel->getActiveSheet()->setCellValue('K55', "=SUM(E55:J55)");



//Data tambahan tapi di blank (white font color)


$objPHPExcel->getActiveSheet()->setCellValue('E100', $kewajiban_e_aud_a);
$objPHPExcel->getActiveSheet()->setCellValue('F100', $kewajiban_e_eur_a);
$objPHPExcel->getActiveSheet()->setCellValue('G100', $kewajiban_e_hkd_a);
$objPHPExcel->getActiveSheet()->setCellValue('H100', $kewajiban_e_jpy_a);
$objPHPExcel->getActiveSheet()->setCellValue('I100', $kewajiban_e_sgd_a);
$objPHPExcel->getActiveSheet()->setCellValue('J100', $kewajiban_e_usd_a);

$objPHPExcel->getActiveSheet()->setCellValue('E101', $kewajiban_e_aud_b);
$objPHPExcel->getActiveSheet()->setCellValue('F101', $kewajiban_e_eur_b);
$objPHPExcel->getActiveSheet()->setCellValue('G101', $kewajiban_e_hkd_b);
$objPHPExcel->getActiveSheet()->setCellValue('H101', $kewajiban_e_jpy_b);
$objPHPExcel->getActiveSheet()->setCellValue('I101', $kewajiban_e_sgd_b);
$objPHPExcel->getActiveSheet()->setCellValue('J101', $kewajiban_e_usd_b);

$objPHPExcel->getActiveSheet()->setCellValue('E102', $kewajiban_e_aud_c);
$objPHPExcel->getActiveSheet()->setCellValue('F102', $kewajiban_e_eur_c);
$objPHPExcel->getActiveSheet()->setCellValue('G102', $kewajiban_e_hkd_c);
$objPHPExcel->getActiveSheet()->setCellValue('H102', $kewajiban_e_jpy_c);
$objPHPExcel->getActiveSheet()->setCellValue('I102', $kewajiban_e_sgd_c);
$objPHPExcel->getActiveSheet()->setCellValue('J102', $kewajiban_e_usd_c);

// tambahan  

$objPHPExcel->getActiveSheet()->setCellValue('E103', $kewajiban_e_aud_d);
$objPHPExcel->getActiveSheet()->setCellValue('F103', $kewajiban_e_eur_d);
$objPHPExcel->getActiveSheet()->setCellValue('G103', $kewajiban_e_hkd_d);
$objPHPExcel->getActiveSheet()->setCellValue('H103', $kewajiban_e_jpy_d);
$objPHPExcel->getActiveSheet()->setCellValue('I103', $kewajiban_e_sgd_d);
$objPHPExcel->getActiveSheet()->setCellValue('J103', $kewajiban_e_usd_d);

$objPHPExcel->getActiveSheet()->setCellValue('E104', $kewajiban_e_aud_e);
$objPHPExcel->getActiveSheet()->setCellValue('F104', $kewajiban_e_eur_e);
$objPHPExcel->getActiveSheet()->setCellValue('G104', $kewajiban_e_hkd_e);
$objPHPExcel->getActiveSheet()->setCellValue('H104', $kewajiban_e_jpy_e);
$objPHPExcel->getActiveSheet()->setCellValue('I104', $kewajiban_e_sgd_e);
$objPHPExcel->getActiveSheet()->setCellValue('J104', $kewajiban_e_usd_e);

$objPHPExcel->getActiveSheet()->setCellValue('E105', $kewajiban_e_aud_f);
$objPHPExcel->getActiveSheet()->setCellValue('F105', $kewajiban_e_eur_f);
$objPHPExcel->getActiveSheet()->setCellValue('G105', $kewajiban_e_hkd_f);
$objPHPExcel->getActiveSheet()->setCellValue('H105', $kewajiban_e_jpy_f);
$objPHPExcel->getActiveSheet()->setCellValue('I105', $kewajiban_e_sgd_f);
$objPHPExcel->getActiveSheet()->setCellValue('J105', $kewajiban_e_usd_f);


#########  ONLY FOR EXCEL ###########################
/*
NOP101010000    NOP101010010    10
NOP101020000    NOP101020010    15
NOP102010000    NOP102010010    29
NOP201010000    NOP201010010    32
NOP201020000    NOP201020010    33
NOP201030000    NOP201030010    61
NOP201040000    NOP201040010    69
NOP201050000    NOP201050010    31
NOP202010000    NOP202010010    39
NOP202020000    NOP202020010    51
NOP202030000    NOP202030010    67
NOP202040000    NOP202040010    65
NOP202050000    NOP202050010    35
NOP202050000    NOP202050020    52
NOP202050000    NOP202050030
NOP202050000    NOP202050040
NOP201050000    NOP201050020
NOP201050000    NOP201050030    34
*/

//$array_code=array("10","15","29","31","32","33","34","35","39","51","52","61","65","67","69");
//$num_row_excel=200;
//foreach ($array_code as $key => $value) {

$objPHPExcel->getActiveSheet()->setCellValue("E200", "$code_10_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F200", "$code_10_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G200", "$code_10_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H200", "$code_10_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I200", "$code_10_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J200", "$code_10_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K200", "$code_10_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E201", "$code_15_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F201", "$code_15_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G201", "$code_15_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H201", "$code_15_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I201", "$code_15_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J201", "$code_15_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K201", "$code_15_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E202", "$code_29_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F202", "$code_29_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G202", "$code_29_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H202", "$code_29_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I202", "$code_29_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J202", "$code_29_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K202", "$code_29_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E203", "$code_31_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F203", "$code_31_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G203", "$code_31_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H203", "$code_31_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I203", "$code_31_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J203", "$code_31_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K203", "$code_31_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E204", "$code_32_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F204", "$code_32_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G204", "$code_32_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H204", "$code_32_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I204", "$code_32_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J204", "$code_32_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K204", "$code_32_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E205", "$code_33_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F205", "$code_33_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G205", "$code_33_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H205", "$code_33_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I205", "$code_33_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J205", "$code_33_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K205", "$code_33_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E206", "$code_34_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F206", "$code_34_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G206", "$code_34_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H206", "$code_34_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I206", "$code_34_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J206", "$code_34_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K206", "$code_34_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E207", "$code_35_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F207", "$code_35_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G207", "$code_35_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H207", "$code_35_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I207", "$code_35_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J207", "$code_35_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K207", "$code_35_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E208", "$code_39_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F208", "$code_39_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G208", "$code_39_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H208", "$code_39_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I208", "$code_39_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J208", "$code_39_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K208", "$code_39_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E209", "$code_51_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F209", "$code_51_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G209", "$code_51_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H209", "$code_51_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I209", "$code_51_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J209", "$code_51_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K209", "$code_51_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E210", "$code_52_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F210", "$code_52_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G210", "$code_52_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H210", "$code_52_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I210", "$code_52_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J210", "$code_52_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K210", "$code_52_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E211", "$code_61_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F211", "$code_61_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G211", "$code_61_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H211", "$code_61_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I211", "$code_61_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J211", "$code_61_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K211", "$code_61_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E212", "$code_65_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F212", "$code_65_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G212", "$code_65_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H212", "$code_65_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I212", "$code_65_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J212", "$code_65_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K212", "$code_65_idr");

$objPHPExcel->getActiveSheet()->setCellValue("E213", "$code_67_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F213", "$code_67_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G213", "$code_67_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H213", "$code_67_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I213", "$code_67_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J213", "$code_67_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K213", "$code_67_idr");


$objPHPExcel->getActiveSheet()->setCellValue("E214", "$code_69_aud");
$objPHPExcel->getActiveSheet()->setCellValue("F214", "$code_69_eur");
$objPHPExcel->getActiveSheet()->setCellValue("G214", "$code_69_hkd");
$objPHPExcel->getActiveSheet()->setCellValue("H214", "$code_69_jpy");
$objPHPExcel->getActiveSheet()->setCellValue("I214", "$code_69_sgd");
$objPHPExcel->getActiveSheet()->setCellValue("J214", "$code_69_usd");
$objPHPExcel->getActiveSheet()->setCellValue("K214", "$code_69_idr");



//$num_row_excel++;

//}






$objPHPExcel->getActiveSheet()->getStyle('E16:K18')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E20:K22')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E25:K29')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('E31:K31')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E34:K38')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('E40:K40')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('E43:K43')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E46:K46')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E49:K49')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('E51:K51')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E53:K53')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E55:K55')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('E32:K32')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('E41:K41')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('E100:J105')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');


//$objPHPExcel->getActiveSheet()->getStyle('M21')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

#########  untuk txt disimpan pada excel


$objPHPExcel->getActiveSheet()->getStyle('E200:K222')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');



$objPHPExcel->getActiveSheet()->getStyle('E53:K53')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('E55:K55')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
//$objPHPExcel->getActiveSheet()->getStyle('E42:E47')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
//$objPHPExcel->getActiveSheet()->getStyle('E50:E61')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));



$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
//$styleArraybackgroundRed = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->getStyle('E16:K57')->applyFromArray($styleArrayFontBold);

//-------
for ($i=100;$i<104;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}
for ($i=100;$i<104;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=100;$i<104;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}

for($i=100;$i<104;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=100;$i<104;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=100;$i<104;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
//-------

for ($i=16;$i<19;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}
for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}

for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=16;$i<19;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}



for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=20;$i<23;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}



for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=25;$i<30;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}



for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=32;$i<33;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}



for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=34;$i<39;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}



for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=41;$i<42;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}
//SELISIH REKENING ADM
$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('C43', '3');
$objPHPExcel->getActiveSheet()->setCellValue('D43', 'Selisih Rekening Administratif (B.1 - B.2)');
$objPHPExcel->getActiveSheet()->getStyle('A43:D43')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'C.');
$objPHPExcel->getActiveSheet()->setCellValue('D45', 'Posisi Devisa Netto per Valuta');
$objPHPExcel->getActiveSheet()->setCellValue('D46', '(A.3 + B.3)');
$objPHPExcel->getActiveSheet()->getStyle('A45:Z45')->applyFromArray($styleArrayFont);


$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B48', 'D.');
$objPHPExcel->getActiveSheet()->setCellValue('D48', 'Posisi Devisa Netto');
$objPHPExcel->getActiveSheet()->setCellValue('D49', '(Nilai Absolut C)');
$objPHPExcel->getActiveSheet()->getStyle('A48')->applyFromArray($styleArrayFont);

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$objPHPExcel->getActiveSheet()->setCellValue('B51', 'E.');
$objPHPExcel->getActiveSheet()->setCellValue('D51', 'Modal dalam Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B53', 'F.');
$objPHPExcel->getActiveSheet()->setCellValue('D53', '% PDN terhadap modal (A/E) Neraca');
$objPHPExcel->getActiveSheet()->setCellValue('B55', 'G.');
$objPHPExcel->getActiveSheet()->setCellValue('D55', '% PDN terhadap modal (D/E) Neraca & Rek. Adm.');

//$objPHPExcel->getActiveSheet()->getStyle('A51:Z55')->applyFromArray($styleArrayFont);


$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

//$objPHPExcel->getActiveSheet()->getCell('A1')->setValue('Some text');
$objPHPExcel->getActiveSheet()->getStyle('E100:K250')->applyFromArray($styleArrayColorFont);

//$objPHPExcel->getActiveSheet()->setCellValue('A1', 'For The Month');



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


//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArraybackgroundRed);
//$objPHPExcel->getActiveSheet()->getStyle('A1:Z8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:Z12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A58:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('L13:Z57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F9:Z10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A13:A57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A9:A10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet()->setCellValue('H14', 'JPY');
$objPHPExcel->getActiveSheet()->setCellValue('I14', 'SGD');
$objPHPExcel->getActiveSheet()->setCellValue('J14', 'USD');
$objPHPExcel->getActiveSheet()->setCellValue('K13', 'Jumlah');



############### TAMBAHAN EXCEL #################################


#baris 221 E sd K   =IF(E201>0,E201,0)
$objPHPExcel->getActiveSheet()->setCellValue("E221", "=IF(E201>0,E201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("F221", "=IF(F201>0,F201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("G221", "=IF(G201>0,G201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("H221", "=IF(H201>0,H201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("I221", "=IF(I201>0,I201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("J221", "=IF(J201>0,J201,0)");
$objPHPExcel->getActiveSheet()->setCellValue("K221", "=IF(K201>0,K201,0)");
#baris 222 =IF(E201<0,(ABS(E201)+ABS(E202)),ABS(E202))
$objPHPExcel->getActiveSheet()->setCellValue("E222", "=IF(E201<0,(ABS(E201)+ABS(E202)),ABS(E202))");
$objPHPExcel->getActiveSheet()->setCellValue("F222", "=IF(F201<0,(ABS(F201)+ABS(F202)),ABS(F202))");
$objPHPExcel->getActiveSheet()->setCellValue("G222", "=IF(G201<0,(ABS(G201)+ABS(G202)),ABS(G202))");
$objPHPExcel->getActiveSheet()->setCellValue("H222", "=IF(H201<0,(ABS(H201)+ABS(H202)),ABS(H202))");
$objPHPExcel->getActiveSheet()->setCellValue("I222", "=IF(I201<0,(ABS(I201)+ABS(I202)),ABS(I202))");
$objPHPExcel->getActiveSheet()->setCellValue("J222", "=IF(J201<0,(ABS(J201)+ABS(J202)),ABS(J202))");
$objPHPExcel->getActiveSheet()->setCellValue("K222", "=IF(K201<0,(ABS(K201)+ABS(K202)),ABS(K202))");








// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('FORMAT PDN IN BANK GARANSI');

// Create a new worksheet, after the default sheet

#---hitung

// A. total aktiva valas  
$akval_aud=$aktiva_valas_aud+$giro_aud;
$akval_eur=$aktiva_valas_eur+$giro_eur;
$akval_hkd=$aktiva_valas_hkd+$giro_hkd;
$akval_jpy=$aktiva_valas_jpy+$giro_jpy;
$akval_sgd=$aktiva_valas_sgd+$giro_sgd;
$akval_usd=$aktiva_valas_usd+$giro_usd;
$akval_tot=$akval_aud+$akval_eur+$akval_hkd+$akval_jpy+$akval_sgd+$akval_usd;
//selisih aktiva dan pasiva valas---------
$a1_min_a2_aud=$akval_aud-(-1)*$pasiva_aud;
$a1_min_a2_eur=$akval_eur-(-1)*$pasiva_eur;
$a1_min_a2_hkd=$akval_hkd-(-1)*$pasiva_hkd;
$a1_min_a2_jpy=$akval_jpy-(-1)*$pasiva_jpy;
$a1_min_a2_sgd=$akval_sgd-(-1)*$pasiva_sgd;
$a1_min_a2_usd=$akval_usd-(-1)*$pasiva_usd;
$a1_min_a2_tot=$a1_min_a2_aud+$a1_min_a2_eur+$a1_min_a2_hkd+$a1_min_a2_jpy+$a1_min_a2_sgd+$a1_min_a2_usd;
$a1_min_a2_tot_abs=abs($a1_min_a2_aud)+abs($a1_min_a2_eur)+abs($a1_min_a2_hkd)+abs($a1_min_a2_jpy)+abs($a1_min_a2_sgd)+abs($a1_min_a2_usd);

//
$b1_min_b2_aud=$tagihan_a_aud+$tagihan_b_aud+$tagihan_c_aud+$tagihan_d_aud+$tagihan_e_aud;
$b1_min_b2_eur=$tagihan_a_eur+$tagihan_b_eur+$tagihan_c_eur+$tagihan_d_eur+$tagihan_e_eur;
$b1_min_b2_hkd=$tagihan_a_hkd+$tagihan_b_hkd+$tagihan_c_hkd+$tagihan_d_hkd+$tagihan_e_hkd;
$b1_min_b2_jpy=$tagihan_a_jpy+$tagihan_b_jpy+$tagihan_c_jpy+$tagihan_d_jpy+$tagihan_e_jpy;
$b1_min_b2_sgd=$tagihan_a_sgd+$tagihan_b_sgd+$tagihan_c_sgd+$tagihan_d_sgd+$tagihan_e_sgd;
$b1_min_b2_usd=$tagihan_a_usd+$tagihan_b_usd+$tagihan_c_usd+$tagihan_d_usd+$tagihan_e_usd;
$b1_min_b2_tot=$b1_min_b2_aud+$b1_min_b2_eur+$b1_min_b2_hkd+$b1_min_b2_jpy+$b1_min_b2_sgd+$b1_min_b2_usd;

$a3_plus_b3_aud=$kewajiban_a_aud+$kewajiban_b_aud+$kewajiban_c_aud+$kewajiban_d_aud+$kewajiban_e_aud;
$a3_plus_b3_eur=$kewajiban_a_eur+$kewajiban_b_eur+$kewajiban_c_eur+$kewajiban_d_eur+$kewajiban_e_eur;
$a3_plus_b3_hkd=$kewajiban_a_hkd+$kewajiban_b_hkd+$kewajiban_c_hkd+$kewajiban_d_hkd+$kewajiban_e_hkd;
$a3_plus_b3_jpy=$kewajiban_a_jpy+$kewajiban_b_jpy+$kewajiban_c_jpy+$kewajiban_d_jpy+$kewajiban_e_jpy;
$a3_plus_b3_sgd=$kewajiban_a_sgd+$kewajiban_b_sgd+$kewajiban_c_sgd+$kewajiban_d_sgd+$kewajiban_e_sgd;
$a3_plus_b3_usd=$kewajiban_a_usd+$kewajiban_b_usd+$kewajiban_c_usd+$kewajiban_d_usd+$kewajiban_e_usd;


$a3_plus_b3_tot=$a3_plus_b3_aud+$a3_plus_b3_eur+$a3_plus_b3_hkd+$a3_plus_b3_jpy+$a3_plus_b3_sgd+$a3_plus_b3_usd;



if($a3_plus_b3_aud>0) $a3_plus_b3_aud=$a3_plus_b3_aud; else $a3_plus_b3_aud=-$a3_plus_b3_aud;
if($a3_plus_b3_eur>0) $a3_plus_b3_eur=$a3_plus_b3_eur; else $a3_plus_b3_eur=-$a3_plus_b3_eur; 
if($a3_plus_b3_hkd>0) $a3_plus_b3_hkd=$a3_plus_b3_hkd; else $a3_plus_b3_hkd=-$a3_plus_b3_hkd; 
if($a3_plus_b3_jpy>0) $a3_plus_b3_jpy=$a3_plus_b3_jpy; else $a3_plus_b3_jpy=-$a3_plus_b3_jpy; 
if($a3_plus_b3_sgd>0) $a3_plus_b3_sgd=$a3_plus_b3_sgd; else $a3_plus_b3_sgd=-$a3_plus_b3_sgd;
if($a3_plus_b3_usd>0) $a3_plus_b3_usd=$a3_plus_b3_usd; else $a3_plus_b3_usd=-$a3_plus_b3_usd; 

#selisih
$b1_b2_aud=$b1_min_b2_aud-$a3_plus_b3_aud;
$b1_b2_eur=$b1_min_b2_eur-$a3_plus_b3_eur;
$b1_b2_hkd=$b1_min_b2_hkd-$a3_plus_b3_hkd;
$b1_b2_jpy=$b1_min_b2_jpy-$a3_plus_b3_jpy;
$b1_b2_sgd=$b1_min_b2_sgd-$a3_plus_b3_sgd;
$b1_b2_usd=$b1_min_b2_usd-$a3_plus_b3_usd;
$b1_b2_tot=$b1_min_b2_tot-$a3_plus_b3_usd;
//$b1_b2_tot=$a3_plus_b3_aud+$a3_plus_b3_eur+$a3_plus_b3_hkd+$a3_plus_b3_jpy+$a3_plus_b3_sgd+$a3_plus_b3_usd;

$x2=rand(1,999999);


// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.ms-excel');
//header('Content-Disposition: attachment;filename="name_of_file.xls"');
//header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('php://output');
$objWriter->save("download/NOP_".$label_tgl."_".$file_eksport.".xls");


$jml_baris_txt=0;


//10  NOP101000001    Aktiva Tidak Termasuk Giro Pada Bank Lain
$objPHPExcel = PHPExcel_IOFactory::load("download/NOP_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();


$modal_table=$objPHPExcel->getActiveSheet()->getCell('E51')->getFormattedValue('#,##0,,;(#,##0,,);"-"');
if (!isset($modal_table) || $modal_table=="" || $modal_table==NULL || $modal_table==0){
$var_modal_idr="";
}else {
$var_modal_idr="99IDR".getTextValue($modal_table).PHP_EOL;
$jml_baris_txt++; 
}

/*  comment 2016-09-02

$aktiva_valas_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$aktiva_valas_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$aktiva_valas_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$aktiva_valas_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$aktiva_valas_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$aktiva_valas_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($aktiva_valas_aud) || $aktiva_valas_aud=="" || $aktiva_valas_aud==NULL || $aktiva_valas_aud==0){
$var_aktiva_aud="";
}else {
$var_aktiva_aud="10AUD".getTextValue($aktiva_valas_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($aktiva_valas_eur) || $aktiva_valas_eur=="" || $aktiva_valas_eur==NULL || $aktiva_valas_eur==0){
$var_aktiva_eur="";
}else {
$var_aktiva_eur="10EUR".getTextValue($aktiva_valas_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($aktiva_valas_hkd) || $aktiva_valas_hkd=="" || $aktiva_valas_hkd==NULL || $aktiva_valas_hkd==0){
$var_aktiva_hkd="";
}else {
$var_aktiva_hkd="10HKD".getTextValue($aktiva_valas_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($aktiva_valas_jpy) || $aktiva_valas_jpy=="" || $aktiva_valas_jpy==NULL || $aktiva_valas_jpy==0){
$var_aktiva_jpy="";
}else {
$var_aktiva_jpy="10JPY".getTextValue($aktiva_valas_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($aktiva_valas_sgd) || $aktiva_valas_sgd=="" || $aktiva_valas_sgd==NULL || $aktiva_valas_sgd==0){
$var_aktiva_sgd="";
}else {
$var_aktiva_sgd="10SGD".getTextValue($aktiva_valas_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($aktiva_valas_usd) || $aktiva_valas_usd=="" || $aktiva_valas_usd==NULL || $aktiva_valas_usd==0){
$var_aktiva_usd="";
}else {
$var_aktiva_usd="10USD".getTextValue($aktiva_valas_usd).PHP_EOL;
$jml_baris_txt++; 
}


//15  NOP101000002    Giro Pada Bank Lain
$giro_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$giro_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$giro_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$giro_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$giro_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$giro_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($giro_aud) || $giro_aud=="" || $giro_aud==NULL || $giro_aud==0){
$var_giro_aud="";
}else {
$var_giro_aud="15AUD".getTextValue($giro_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($giro_eur) || $giro_eur=="" || $giro_eur==NULL || $giro_eur==0){
$var_giro_eur="";
}else {
$var_giro_eur="15EUR".getTextValue($giro_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($giro_hkd) || $giro_hkd=="" || $giro_hkd==NULL || $giro_hkd==0){
$var_giro_hkd="";
}else {
$var_giro_hkd="15HKD".getTextValue($giro_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($giro_jpy) || $giro_jpy=="" || $giro_jpy==NULL || $giro_jpy==0){
$var_giro_jpy="";
}else {
$var_giro_jpy="15JPY".getTextValue($giro_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($giro_sgd) || $giro_sgd=="" || $giro_sgd==NULL || $giro_sgd==0){
$var_giro_sgd="";
}else {
$var_giro_sgd="15SGD".getTextValue($giro_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($giro_usd) || $giro_usd=="" || $giro_usd==NULL || $giro_usd==0){
$var_giro_usd="";
}else {
$var_giro_usd="15USD".getTextValue($giro_usd).PHP_EOL;
$jml_baris_txt++; 
}

//29  NOP102000001    Pasiva
$pasiva_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$pasiva_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$pasiva_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$pasiva_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$pasiva_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$pasiva_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
if (!isset($pasiva_aud) || $pasiva_aud=="" || $pasiva_aud==NULL || $pasiva_aud==0){
$var_pasiva_aud="";
}else {
$var_pasiva_aud="29AUD".getTextValue($pasiva_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($pasiva_eur) || $pasiva_eur=="" || $pasiva_eur==NULL || $pasiva_eur==0){
$var_pasiva_eur="";
}else {
$var_pasiva_eur="29EUR".getTextValue($pasiva_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($pasiva_hkd) || $pasiva_hkd=="" || $pasiva_hkd==NULL || $pasiva_hkd==0){
$var_pasiva_hkd="";
}else {
$var_pasiva_hkd="29HKD".getTextValue($pasiva_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($pasiva_jpy) || $pasiva_jpy=="" || $pasiva_jpy==NULL || $pasiva_jpy==0){
$var_pasiva_jpy="";
}else {
$var_pasiva_jpy="29JPY".getTextValue($pasiva_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($pasiva_sgd) || $pasiva_sgd=="" || $pasiva_sgd==NULL || $pasiva_sgd==0){
$var_pasiva_sgd="";
}else {
$var_pasiva_sgd="29SGD".getTextValue($pasiva_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($pasiva_usd) || $pasiva_usd=="" || $pasiva_usd==NULL || $pasiva_usd==0){
$var_pasiva_usd="";
}else {
$var_pasiva_usd="29USD".getTextValue($pasiva_usd).PHP_EOL;
$jml_baris_txt++; 
}

//32  NOP201000001    a.Kontrak Pembelian Forward
$tagihan_a_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_a_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_a_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_a_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_a_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_a_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
if (!isset($tagihan_a_aud) || $tagihan_a_aud=="" || $tagihan_a_aud==NULL || $tagihan_a_aud==0){
$var_tagihan_a_aud="";
}else {
$var_tagihan_a_aud="32AUD".getTextValue($tagihan_a_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_a_eur) || $tagihan_a_eur=="" || $tagihan_a_eur==NULL || $tagihan_a_eur==0){
$var_tagihan_a_eur="";
}else {
$var_tagihan_a_eur="32EUR".getTextValue($tagihan_a_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_a_hkd) || $tagihan_a_hkd=="" || $tagihan_a_hkd==NULL || $tagihan_a_hkd==0){
$var_tagihan_a_hkd="";
}else {
$var_tagihan_a_hkd="32HKD".getTextValue($tagihan_a_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_a_jpy) || $tagihan_a_jpy=="" || $tagihan_a_jpy==NULL || $tagihan_a_jpy==0){
$var_tagihan_a_jpy="";
}else {
$var_tagihan_a_jpy="32JPY".getTextValue($tagihan_a_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_a_sgd) || $tagihan_a_sgd=="" || $tagihan_a_sgd==NULL || $tagihan_a_sgd==0){
$var_atagihan_a_sgd="";
}else {
$var_tagihan_a_sgd="32SGD".getTextValue($tagihan_a_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_a_usd) || $tagihan_a_usd=="" || $tagihan_a_usd==NULL || $tagihan_a_usd==0){
$var_tagihan_a_usd="";
}else {
$var_tagihan_a_usd="32USD".getTextValue($tagihan_a_usd).PHP_EOL;
$jml_baris_txt++; 
}


//33  NOP201000002    b.kontrak Pembelian Future
$tagihan_b_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_b_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_b_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_b_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_b_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_b_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
if (!isset($tagihan_b_aud) || $tagihan_b_aud=="" || $tagihan_b_aud==NULL || $tagihan_b_aud==0){
$var_tagihan_b_aud="";
}else {
$var_tagihan_b_aud="33AUD".getTextValue($tagihan_b_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_b_eur) || $tagihan_b_eur=="" || $tagihan_b_eur==NULL || $tagihan_b_eur==0){
$var_tagihan_b_eur="";
}else {
$var_tagihan_b_eur="33EUR".getTextValue($tagihan_b_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_b_hkd) || $tagihan_b_hkd=="" || $tagihan_b_hkd==NULL || $tagihan_b_hkd==0){
$var_tagihan_b_hkd="";
}else {
$var_tagihan_b_hkd="33HKD".getTextValue($tagihan_b_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_b_jpy) || $tagihan_b_jpy=="" || $tagihan_b_jpy==NULL || $tagihan_b_jpy==0){
$var_tagihan_b_jpy="";
}else {
$var_tagihan_b_jpy="33JPY".getTextValue($tagihan_b_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_b_sgd) || $tagihan_b_sgd=="" || $tagihan_b_sgd==NULL || $tagihan_b_sgd==0){
$var_tagihan_b_sgd="";
}else {
$var_tagihan_b_sgd="33SGD".getTextValue($tagihan_b_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_b_usd) || $tagihan_b_usd=="" || $tagihan_b_usd==NULL || $tagihan_b_usd==0){
$var_tagihan_b_usd="";
}else {
$var_tagihan_b_usd="33USD".getTextValue($tagihan_b_usd).PHP_EOL;
$jml_baris_txt++; 
}

###########  ini tidak ada #######################################################
//34  NOP201000005    e. Rekening Administratif Tagihan Valas diluar kontrak pemberian forward, futures, dan option
/*
$tagihan_e_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_e_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_e_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_e_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_e_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_e_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($tagihan_e_aud) || $tagihan_e_aud=="" || $tagihan_e_aud==NULL || $tagihan_e_aud==0){
$var_tagihan_e_aud="";
}else {
$var_tagihan_e_aud="34AUD".getTextValue($tagihan_e_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_e_eur) || $tagihan_e_eur=="" || $tagihan_e_eur==NULL || $tagihan_e_eur==0){
$var_tagihan_e_eur="";
}else {
$var_tagihan_e_eur="34EUR".getTextValue($tagihan_e_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_e_hkd) || $tagihan_e_hkd=="" || $tagihan_e_hkd==NULL || $tagihan_e_hkd==0){
$var_tagihan_e_hkd="";
}else {
$var_tagihan_e_hkd="34HKD".getTextValue($tagihan_e_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_e_jpy) || $tagihan_e_jpy=="" || $tagihan_e_jpy==NULL || $tagihan_e_jpy==0){
$var_tagihan_e_jpy="";
}else {
$var_tagihan_e_jpy="34JPY".getTextValue($tagihan_e_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_e_sgd) || $tagihan_e_sgd=="" || $tagihan_e_sgd==NULL || $tagihan_e_sgd==0){
$var_tagihan_e_sgd="";
}else {
$var_tagihan_e_sgd="34SGD".getTextValue($tagihan_e_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_e_usd) || $tagihan_e_usd=="" || $tagihan_e_usd==NULL || $tagihan_e_usd==0){
$var_tagihan_e_usd="";
}else {
$var_tagihan_e_usd="34USD".getTextValue($tagihan_e_usd).PHP_EOL;
$jml_baris_txt++; 
}
*/  // end comment 2016-09-02


###########################################  end premi ##################

#####################################################

/*

#####---Beda dan belum dicari nominal ----------
//35  NOP2020000051   Rekening Administratif
$kewajiban_e_aud_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_a=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J100')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($kewajiban_e_aud_a) || $kewajiban_e_aud_a=="" || $kewajiban_e_aud_a==NULL || $kewajiban_e_aud_a==0){
$var_akewajiban_e_aud="";
}else {
$var_akewajiban_e_aud="35AUD".getTextValue($kewajiban_e_aud_a).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_a) || $kewajiban_e_eur_a=="" || $kewajiban_e_eur_a==NULL || $kewajiban_e_eur_a==0){
$var_akewajiban_e_eur="";
}else {
$var_akewajiban_e_eur="35EUR".getTextValue($kewajiban_e_eur_a).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_a) || $kewajiban_e_hkd_a=="" || $kewajiban_e_hkd_a==NULL || $kewajiban_e_hkd_a==0){
$var_akewajiban_e_hkd="";
}else {
$var_akewajiban_e_hkd="35HKD".getTextValue($kewajiban_e_hkd_a).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_a) || $kewajiban_e_jpy_a=="" || $kewajiban_e_jpy_a==NULL || $kewajiban_e_jpy_a==0){
$var_akewajiban_e_jpy="";
}else {
$var_akewajiban_e_jpy="35JPY".getTextValue($kewajiban_e_jpy_a).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_a) || $kewajiban_e_sgd_a=="" || $kewajiban_e_sgd_a==NULL || $kewajiban_e_sgd_a==0){
$var_akewajiban_e_sgd="";
}else {
$var_akewajiban_e_sgd="35SGD".getTextValue($kewajiban_e_sgd_a).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_a) || $kewajiban_e_usd_a=="" || $kewajiban_e_usd_a==NULL || $kewajiban_e_usd_a==0){
$var_akewajiban_e_usd="";
}else {
$var_akewajiban_e_usd="35USD".getTextValue($kewajiban_e_usd_a).PHP_EOL;
$jml_baris_txt++; 
}
//NOP2020000053   Kontrak Penjualan Forward  39
//NOP2020000052   Transaksi Derivatif diluar kontrak Penjualan Forward,Futures dan Option 52


//52  NOP2020000052   Transaksi Derivatif diluar kontrak Penjualan Forward,Futures dan Option
$kewajiban_e_aud_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_c=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J101')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
        if (!isset($kewajiban_e_aud_c) || $kewajiban_e_aud_c=="" || $kewajiban_e_aud_c==NULL || $kewajiban_e_aud_c==0){
$var_ckewajiban_e_aud="";
}else {
$var_ckewajiban_e_aud="52AUD".getTextValue($kewajiban_e_aud_c).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_c) || $kewajiban_e_eur_c=="" || $kewajiban_e_eur_c==NULL || $kewajiban_e_eur_c==0){
$var_ckewajiban_e_eur="";
}else {
$var_ckewajiban_e_eur="52EUR".getTextValue($kewajiban_e_eur_c).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_c) || $kewajiban_e_hkd_c=="" || $kewajiban_e_hkd_c==NULL || $kewajiban_e_hkd_c==0){
$var_ckewajiban_e_hkd="";
}else {
$var_ckewajiban_e_hkd="52HKD".getTextValue($kewajiban_e_hkd_c).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_c) || $kewajiban_e_jpy_c=="" || $kewajiban_e_jpy_c==NULL || $kewajiban_e_jpy_c==0){
$var_ckewajiban_e_jpy="";
}else {
$var_ckewajiban_e_jpy="52JPY".getTextValue($kewajiban_e_jpy_c).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_c) || $kewajiban_e_sgd_c=="" || $kewajiban_e_sgd_c==NULL || $kewajiban_e_sgd_c==0){
$var_ckewajiban_e_sgd="";
}else {
$var_ckewajiban_e_sgd="52SGD".getTextValue($kewajiban_e_sgd_c).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_c) || $kewajiban_e_usd_c=="" || $kewajiban_e_usd_c==NULL || $kewajiban_e_usd_c==0){
$var_ckewajiban_e_usd="";
}else {
$var_ckewajiban_e_usd="52USD".getTextValue($kewajiban_e_usd_c).PHP_EOL;
$jml_baris_txt++; 
}



//39  NOP2020000053   Kontrak Penjualan Forward
$kewajiban_e_aud_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_b=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J102')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($kewajiban_e_aud_b) || $kewajiban_e_aud_b=="" || $kewajiban_e_aud_b==NULL || $kewajiban_e_aud_b==0){
$var_bkewajiban_e_aud="";
}else {
$var_bkewajiban_e_aud="39AUD".getTextValue($kewajiban_e_aud_b).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_b) || $kewajiban_e_eur_b=="" || $kewajiban_e_eur_b==NULL || $kewajiban_e_eur_b==0){
$var_bkewajiban_e_eur="";
}else {
$var_bkewajiban_e_eur="39EUR".getTextValue($kewajiban_e_eur_b).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_b) || $kewajiban_e_hkd_b=="" || $kewajiban_e_hkd_b==NULL || $kewajiban_e_hkd_b==0){
$var_bkewajiban_e_hkd="";
}else {
$var_bkewajiban_e_hkd="39HKD".getTextValue($kewajiban_e_hkd_b).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_b) || $kewajiban_e_jpy_b=="" || $kewajiban_e_jpy_b==NULL || $kewajiban_e_jpy_b==0){
$var_bkewajiban_e_jpy="";
}else {
$var_bkewajiban_e_jpy="39JPY".getTextValue($kewajiban_e_jpy_b).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_b) || $kewajiban_e_sgd_b=="" || $kewajiban_e_sgd_b==NULL || $kewajiban_e_sgd_b==0){
$var_bkewajiban_e_sgd="";
}else {
$var_bkewajiban_e_sgd="39SGD".getTextValue($kewajiban_e_sgd_b).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_b) || $kewajiban_e_usd_b=="" || $kewajiban_e_usd_b==NULL || $kewajiban_e_usd_b==0){
$var_bkewajiban_e_usd="";
}else {
$var_bkewajiban_e_usd="39USD".getTextValue($kewajiban_e_usd_b).PHP_EOL;
$jml_baris_txt++; 
}


//NOP2020000054   Rekening Tagihan Administratif  31
$kewajiban_e_aud_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_d=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J103')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
if (!isset($kewajiban_e_aud_d) || $kewajiban_e_aud_d=="" || $kewajiban_e_aud_d==NULL || $kewajiban_e_aud_d==0){
$var_dkewajiban_e_aud="";
}else {
$var_dkewajiban_e_aud="31AUD".getTextValue($kewajiban_e_aud_d).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_d) || $kewajiban_e_eur_d=="" || $kewajiban_e_eur_d==NULL || $kewajiban_e_eur_d==0){
$var_dkewajiban_e_eur="";
}else {
$var_dkewajiban_e_eur="31EUR".getTextValue($kewajiban_e_eur_d).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_d) || $kewajiban_e_hkd_d=="" || $kewajiban_e_hkd_d==NULL || $kewajiban_e_hkd_d==0){
$var_dkewajiban_e_hkd="";
}else {
$var_dkewajiban_e_hkd="31HKD".getTextValue($kewajiban_e_hkd_d).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_d) || $kewajiban_e_jpy_d=="" || $kewajiban_e_jpy_d==NULL || $kewajiban_e_jpy_d==0){
$var_dkewajiban_e_jpy="";
}else {
$var_dkewajiban_e_jpy="31JPY".getTextValue($kewajiban_e_jpy_d).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_d) || $kewajiban_e_sgd_d=="" || $kewajiban_e_sgd_d==NULL || $kewajiban_e_sgd_d==0){
$var_dkewajiban_e_sgd="";
}else {
$var_dkewajiban_e_sgd="31SGD".getTextValue($kewajiban_e_sgd_d).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_d) || $kewajiban_e_usd_d=="" || $kewajiban_e_usd_d==NULL || $kewajiban_e_usd_d==0){
$var_dkewajiban_e_usd="";
}else {
$var_dkewajiban_e_usd="31USD".getTextValue($kewajiban_e_usd_d).PHP_EOL;
$jml_baris_txt++; 
}

//NOP2010000055   Kontrak Pembelian Forward    32

$kewajiban_e_aud_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_e=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J104')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
if (!isset($kewajiban_e_aud_e) || $kewajiban_e_aud_e=="" || $kewajiban_e_aud_e==NULL || $kewajiban_e_aud_e==0){
$var_ekewajiban_e_aud="";
}else {
$var_ekewajiban_e_aud="32AUD".getTextValue($kewajiban_e_aud_e).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_e) || $kewajiban_e_eur_e=="" || $kewajiban_e_eur_e==NULL || $kewajiban_e_eur_e==0){
$var_ekewajiban_e_eur="";
}else {
$var_ekewajiban_e_eur="32EUR".getTextValue($kewajiban_e_eur_e).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_e) || $kewajiban_e_hkd_e=="" || $kewajiban_e_hkd_e==NULL || $kewajiban_e_hkd_e==0){
$var_ekewajiban_e_hkd="";
}else {
$var_ekewajiban_e_hkd="32HKD".getTextValue($kewajiban_e_hkd_e).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_e) || $kewajiban_e_jpy_e=="" || $kewajiban_e_jpy_e==NULL || $kewajiban_e_jpy_e==0){
$var_ekewajiban_e_jpy="";
}else {
$var_ekewajiban_e_jpy="32JPY".getTextValue($kewajiban_e_jpy_e).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_e) || $kewajiban_e_sgd_e=="" || $kewajiban_e_sgd_e==NULL || $kewajiban_e_sgd_e==0){
$var_ekewajiban_e_sgd="";
}else {
$var_ekewajiban_e_sgd="32SGD".getTextValue($kewajiban_e_sgd_e).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_e) || $kewajiban_e_usd_e=="" || $kewajiban_e_usd_e==NULL || $kewajiban_e_usd_e==0){
$var_ekewajiban_e_usd="";
}else {
$var_ekewajiban_e_usd="32USD".getTextValue($kewajiban_e_usd_e).PHP_EOL;
$jml_baris_txt++; 
}



//NOP2010000056   Transaksi Derivatif diluar kontrak Pembelian Forward,Futures dan Option 34
$kewajiban_e_aud_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_eur_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_hkd_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_jpy_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_sgd_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_e_usd_f=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J105')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($kewajiban_e_aud_f) || $kewajiban_e_aud_f=="" || $kewajiban_e_aud_f==NULL || $kewajiban_e_aud_f==0){
$var_fkewajiban_e_aud="";
}else {
$var_fkewajiban_e_aud="34AUD".getTextValue($kewajiban_e_aud_f).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_eur_f) || $kewajiban_e_eur_f=="" || $kewajiban_e_eur_f==NULL || $kewajiban_e_eur_f==0){
$var_fkewajiban_e_eur="";
}else {
$var_fkewajiban_e_eur="34EUR".getTextValue($kewajiban_e_eur_f).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_hkd_f) || $kewajiban_e_hkd_f=="" || $kewajiban_e_hkd_f==NULL || $kewajiban_e_hkd_f==0){
$var_fkewajiban_e_hkd="";
}else {
$var_fkewajiban_e_hkd="34HKD".getTextValue($kewajiban_e_hkd_f).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_jpy_f) || $kewajiban_e_jpy_f=="" || $kewajiban_e_jpy_f==NULL || $kewajiban_e_jpy_f==0){
$var_fkewajiban_e_jpy="";
}else {
$var_fkewajiban_e_jpy="34JPY".getTextValue($kewajiban_e_jpy_f).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_sgd_f) || $kewajiban_e_sgd_f=="" || $kewajiban_e_sgd_f==NULL || $kewajiban_e_sgd_f==0){
$var_fkewajiban_e_sgd="";
}else {
$var_fkewajiban_e_sgd="34SGD".getTextValue($kewajiban_e_sgd_f).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_e_usd_f) || $kewajiban_e_usd_f=="" || $kewajiban_e_usd_f==NULL || $kewajiban_e_usd_f==0){
$var_fkewajiban_e_usd="";
}else {
$var_fkewajiban_e_usd="34USD".getTextValue($kewajiban_e_usd_f).PHP_EOL;
$jml_baris_txt++; 
}








//echo $var_ckewajiban_e_usd;
//die();

#####---------------------------------------------------------------------------------------------------
//61  NOP201000003    c.Kontrak Penjualan Put Option (Bank sebagai Writter)/Tagihan valas
$tagihan_c_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_c_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_c_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_c_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_c_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_c_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($tagihan_c_aud) || $tagihan_c_aud=="" || $tagihan_c_aud==NULL || $tagihan_c_aud==0){
$var_tagihan_c_aud="";
}else {
$var_tagihan_c_aud="61AUD".getTextValue($tagihan_c_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_c_eur) || $tagihan_c_eur=="" || $tagihan_c_eur==NULL || $tagihan_c_eur==0){
$var_tagihan_c_eur="";
}else {
$var_tagihan_c_eur="61EUR".getTextValue($tagihan_c_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_c_hkd) || $tagihan_c_hkd=="" || $tagihan_c_hkd==NULL || $tagihan_c_hkd==0){
$var_tagihan_c_hkd="";
}else {
$var_tagihan_c_hkd="61HKD".getTextValue($tagihan_c_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_c_jpy) || $tagihan_c_jpy=="" || $tagihan_c_jpy==NULL || $tagihan_c_jpy==0){
$var_tagihan_c_jpy="";
}else {
$var_tagihan_c_jpy="61JPY".getTextValue($tagihan_c_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_c_sgd) || $tagihan_c_sgd=="" || $tagihan_c_sgd==NULL || $tagihan_c_sgd==0){
$var_tagihan_c_sgd="";
}else {
$var_tagihan_c_sgd="61SGD".getTextValue($tagihan_c_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_c_usd) || $tagihan_c_usd=="" || $tagihan_c_usd==NULL || $tagihan_c_usd==0){
$var_tagihan_c_usd="";
}else {
$var_tagihan_c_usd="61USD".getTextValue($tagihan_c_usd).PHP_EOL;
$jml_baris_txt++; 
}
 
//65  NOP202000004    d. Kontrak pembelian put options (bank sebagai holder, khusus back to back option)/Kewajiban valas
$kewajiban_d_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_d_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_d_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_d_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_d_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_d_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($kewajiban_d_aud) || $kewajiban_d_aud=="" || $kewajiban_d_aud==NULL || $kewajiban_d_aud==0){
$var_kewajiban_d_aud="";
}else {
$var_kewajiban_d_aud="65AUD".getTextValue($kewajiban_d_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_d_eur) || $kewajiban_d_eur=="" || $kewajiban_d_eur==NULL || $kewajiban_d_eur==0){
$var_kewajiban_d_eur="";
}else {
$var_kewajiban_d_eur="65EUR".getTextValue($kewajiban_d_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_d_hkd) || $kewajiban_d_hkd=="" || $kewajiban_d_hkd==NULL || $kewajiban_d_hkd==0){
$var_kewajiban_d_hkd="";
}else {
$var_kewajiban_d_hkd="65HKD".getTextValue($kewajiban_d_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_d_jpy) || $kewajiban_d_jpy=="" || $kewajiban_d_jpy==NULL || $kewajiban_d_jpy==0){
$var_kewajiban_d_jpy="";
}else {
$var_kewajiban_d_jpy="65JPY".getTextValue($kewajiban_d_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_d_sgd) || $kewajiban_d_sgd=="" || $kewajiban_d_sgd==NULL || $kewajiban_d_sgd==0){
$var_kewajiban_d_sgd="";
}else {
$var_kewajiban_d_sgd="65SGD".getTextValue($kewajiban_d_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_d_usd) || $kewajiban_d_usd=="" || $kewajiban_d_usd==NULL || $kewajiban_d_usd==0){
$var_kewajiban_d_usd="";
}else {
$var_kewajiban_d_usd="65USD".getTextValue($kewajiban_d_usd).PHP_EOL;
$jml_baris_txt++; 
}

//67  NOP202000003    c.  Kontrak penjualan call options (bank sebagai writter)/Kewajiban valas
$kewajiban_c_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_c_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_c_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_c_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_c_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$kewajiban_c_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));

if (!isset($kewajiban_c_aud) || $kewajiban_c_aud=="" || $kewajiban_c_aud==NULL || $kewajiban_c_aud==0){
$var_kewajiban_c_aud="";
}else {
$var_kewajiban_c_aud="67AUD".getTextValue($kewajiban_c_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_c_eur) || $kewajiban_c_eur=="" || $kewajiban_c_eur==NULL || $kewajiban_c_eur==0){
$var_kewajiban_c_eur="";
}else {
$var_kewajiban_c_eur="67EUR".getTextValue($kewajiban_c_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_c_hkd) || $kewajiban_c_hkd=="" || $kewajiban_c_hkd==NULL || $kewajiban_c_hkd==0){
$var_kewajiban_c_hkd="";
}else {
$var_kewajiban_c_hkd="67HKD".getTextValue($kewajiban_c_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_c_jpy) || $kewajiban_c_jpy=="" || $kewajiban_c_jpy==NULL || $kewajiban_c_jpy==0){
$var_kewajiban_c_jpy="";
}else {
$var_kewajiban_c_jpy="67JPY".getTextValue($kewajiban_c_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_c_sgd) || $kewajiban_c_sgd=="" || $kewajiban_c_sgd==NULL || $kewajiban_c_sgd==0){
$var_kewajiban_c_sgd="";
}else {
$var_kewajiban_c_sgd="67SGD".getTextValue($kewajiban_c_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($kewajiban_c_usd) || $kewajiban_c_usd=="" || $kewajiban_c_usd==NULL || $kewajiban_c_usd==0){
$var_kewajiban_c_usd="";
}else {
$var_kewajiban_c_usd="67USD".getTextValue($kewajiban_c_usd).PHP_EOL;
$jml_baris_txt++; 
}

//69  NOP201000004    d. Kontrak pembelian call options (bank sebagai holder, khusus back to back options)/Tagihan valas
$tagihan_d_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_d_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_d_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_d_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_d_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$tagihan_d_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($tagihan_d_aud) || $tagihan_d_aud=="" || $tagihan_d_aud==NULL || $tagihan_d_aud==0){
$var_tagihan_d_aud="";
}else {
$var_tagihan_d_aud="69AUD".getTextValue($tagihan_d_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_d_eur) || $tagihan_d_eur=="" || $tagihan_d_eur==NULL || $tagihan_d_eur==0){
$var_tagihan_d_eur="";
}else {
$var_tagihan_d_eur="69EUR".getTextValue($tagihan_d_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_d_hkd) || $tagihan_d_hkd=="" || $tagihan_d_hkd==NULL || $tagihan_d_hkd==0){
$var_tagihan_d_hkd="";
}else {
$var_tagihan_d_hkd="69HKD".getTextValue($tagihan_d_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_d_jpy) || $tagihan_d_jpy=="" || $tagihan_d_jpy==NULL || $tagihan_d_jpy==0){
$var_tagihan_d_jpy="";
}else {
$var_tagihan_d_jpy="69JPY".getTextValue($tagihan_d_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_d_sgd) || $tagihan_d_sgd=="" || $tagihan_d_sgd==NULL || $tagihan_d_sgd==0){
$var_tagihan_d_sgd="";
}else {
$var_tagihan_d_sgd="69SGD".getTextValue($tagihan_d_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($tagihan_d_usd) || $tagihan_d_usd=="" || $tagihan_d_usd==NULL || $tagihan_d_usd==0){
$var_tagihan_d_usd="";
}else {
$var_tagihan_d_usd="69USD".getTextValue($tagihan_d_usd).PHP_EOL;
$jml_baris_txt++; 
}
*/

########################  METODE BARU TXT ##################################

$xcode_10_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_10_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_10_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_10_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_10_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_10_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J200')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_10_aud) || $xcode_10_aud=="" || $xcode_10_aud==NULL || $xcode_10_aud==0){
$var_code_aud_10="";
}else {
$var_code_aud_10="10AUD".getTextValue($xcode_10_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_10_eur) || $xcode_10_eur=="" || $xcode_10_eur==NULL || $xcode_10_eur==0){
$var_code_eur_10="";
}else {
$var_code_eur_10="10EUR".getTextValue($xcode_10_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_10_hkd) || $xcode_10_hkd=="" || $xcode_10_hkd==NULL || $xcode_10_hkd==0){
$var_code_hkd_10="";
}else {
$var_code_hkd_10="10HKD".getTextValue($xcode_10_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_10_jpy) || $xcode_10_jpy=="" || $xcode_10_jpy==NULL || $xcode_10_jpy==0){
$var_code_jpy_10="";
}else {
$var_code_jpy_10="10JPY".getTextValue($xcode_10_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_10_sgd) || $xcode_10_sgd=="" || $xcode_10_sgd==NULL || $xcode_10_sgd==0){
$var_code_sgd_10="";
}else {
$var_code_sgd_10="10SGD".getTextValue($xcode_10_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_10_usd) || $xcode_10_usd=="" || $xcode_10_usd==NULL || $xcode_10_usd==0){
$var_code_usd_10="";
}else {
$var_code_usd_10="10USD".getTextValue($xcode_10_usd).PHP_EOL;
$jml_baris_txt++; 
}

# code 15

$xcode_15_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_15_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_15_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_15_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_15_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_15_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J221')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_15_aud) || $xcode_15_aud=="" || $xcode_15_aud==NULL || $xcode_15_aud==0){
$var_code_aud_15="";
}else {
$var_code_aud_15="15AUD".getTextValue($xcode_15_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_15_eur) || $xcode_15_eur=="" || $xcode_15_eur==NULL || $xcode_15_eur==0){
$var_code_eur_15="";
}else {
$var_code_eur_15="15EUR".getTextValue($xcode_15_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_15_hkd) || $xcode_15_hkd=="" || $xcode_15_hkd==NULL || $xcode_15_hkd==0){
$var_code_hkd_15="";
}else {
$var_code_hkd_15="15HKD".getTextValue($xcode_15_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_15_jpy) || $xcode_15_jpy=="" || $xcode_15_jpy==NULL || $xcode_15_jpy==0){
$var_code_jpy_15="";
}else {
$var_code_jpy_15="15JPY".getTextValue($xcode_15_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_15_sgd) || $xcode_15_sgd=="" || $xcode_15_sgd==NULL || $xcode_15_sgd==0){
$var_code_sgd_15="";
}else {
$var_code_sgd_15="15SGD".getTextValue($xcode_15_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_15_usd) || $xcode_15_usd=="" || $xcode_15_usd==NULL || $xcode_15_usd==0){
$var_code_usd_15="";
}else {
$var_code_usd_15="15USD".getTextValue($xcode_15_usd).PHP_EOL;
$jml_baris_txt++; 
}




# code 29

$xcode_29_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_29_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_29_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_29_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_29_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_29_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J222')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_29_aud) || $xcode_29_aud=="" || $xcode_29_aud==NULL || $xcode_29_aud==0){
$var_code_aud_29="";
}else {
$var_code_aud_29="29AUD".getTextValue($xcode_29_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_29_eur) || $xcode_29_eur=="" || $xcode_29_eur==NULL || $xcode_29_eur==0){
$var_code_eur_29="";
}else {
$var_code_eur_29="29EUR".getTextValue($xcode_29_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_29_hkd) || $xcode_29_hkd=="" || $xcode_29_hkd==NULL || $xcode_29_hkd==0){
$var_code_hkd_29="";
}else {
$var_code_hkd_29="29HKD".getTextValue($xcode_29_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_29_jpy) || $xcode_29_jpy=="" || $xcode_29_jpy==NULL || $xcode_29_jpy==0){
$var_code_jpy_29="";
}else {
$var_code_jpy_29="29JPY".getTextValue($xcode_29_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_29_sgd) || $xcode_29_sgd=="" || $xcode_29_sgd==NULL || $xcode_29_sgd==0){
$var_code_sgd_29="";
}else {
$var_code_sgd_29="29SGD".getTextValue($xcode_29_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_29_usd) || $xcode_29_usd=="" || $xcode_29_usd==NULL || $xcode_29_usd==0){
$var_code_usd_29="";
}else {
$var_code_usd_29="29USD".getTextValue($xcode_29_usd).PHP_EOL;
$jml_baris_txt++; 
}

# code 31

$xcode_31_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_31_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_31_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_31_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_31_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_31_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J203')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_31_aud) || $xcode_31_aud=="" || $xcode_31_aud==NULL || $xcode_31_aud==0){
$var_code_aud_31="";
}else {
$var_code_aud_31="31AUD".getTextValue($xcode_31_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_31_eur) || $xcode_31_eur=="" || $xcode_31_eur==NULL || $xcode_31_eur==0){
$var_code_eur_31="";
}else {
$var_code_eur_31="31EUR".getTextValue($xcode_31_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_31_hkd) || $xcode_31_hkd=="" || $xcode_31_hkd==NULL || $xcode_31_hkd==0){
$var_code_hkd_31="";
}else {
$var_code_hkd_31="31HKD".getTextValue($xcode_31_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_31_jpy) || $xcode_31_jpy=="" || $xcode_31_jpy==NULL || $xcode_31_jpy==0){
$var_code_jpy_31="";
}else {
$var_code_jpy_31="31JPY".getTextValue($xcode_31_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_31_sgd) || $xcode_31_sgd=="" || $xcode_31_sgd==NULL || $xcode_31_sgd==0){
$var_code_sgd_31="";
}else {
$var_code_sgd_31="31SGD".getTextValue($xcode_31_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_31_usd) || $xcode_31_usd=="" || $xcode_31_usd==NULL || $xcode_31_usd==0){
$var_code_usd_31="";
}else {
$var_code_usd_31="31USD".getTextValue($xcode_31_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 32

$xcode_32_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_32_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_32_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_32_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_32_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_32_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J204')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_32_aud) || $xcode_32_aud=="" || $xcode_32_aud==NULL || $xcode_32_aud==0){
$var_code_aud_32="";
}else {
$var_code_aud_32="32AUD".getTextValue($xcode_32_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_32_eur) || $xcode_32_eur=="" || $xcode_32_eur==NULL || $xcode_32_eur==0){
$var_code_eur_32="";
}else {
$var_code_eur_32="32EUR".getTextValue($xcode_32_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_32_hkd) || $xcode_32_hkd=="" || $xcode_32_hkd==NULL || $xcode_32_hkd==0){
$var_code_hkd_32="";
}else {
$var_code_hkd_32="32HKD".getTextValue($xcode_32_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_32_jpy) || $xcode_32_jpy=="" || $xcode_32_jpy==NULL || $xcode_32_jpy==0){
$var_code_jpy_32="";
}else {
$var_code_jpy_32="32JPY".getTextValue($xcode_32_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_32_sgd) || $xcode_32_sgd=="" || $xcode_32_sgd==NULL || $xcode_32_sgd==0){
$var_code_sgd_32="";
}else {
$var_code_sgd_32="32SGD".getTextValue($xcode_32_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_32_usd) || $xcode_32_usd=="" || $xcode_32_usd==NULL || $xcode_32_usd==0){
$var_code_usd_32="";
}else {
$var_code_usd_32="32USD".getTextValue($xcode_32_usd).PHP_EOL;
$jml_baris_txt++; 
}

# code 33

$xcode_33_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_33_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_33_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_33_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_33_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_33_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J205')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_33_aud) || $xcode_33_aud=="" || $xcode_33_aud==NULL || $xcode_33_aud==0){
$var_code_aud_33="";
}else {
$var_code_aud_33="33AUD".getTextValue($xcode_33_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_33_eur) || $xcode_33_eur=="" || $xcode_33_eur==NULL || $xcode_33_eur==0){
$var_code_eur_33="";
}else {
$var_code_eur_33="33EUR".getTextValue($xcode_33_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_33_hkd) || $xcode_33_hkd=="" || $xcode_33_hkd==NULL || $xcode_33_hkd==0){
$var_code_hkd_33="";
}else {
$var_code_hkd_33="33HKD".getTextValue($xcode_33_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_33_jpy) || $xcode_33_jpy=="" || $xcode_33_jpy==NULL || $xcode_33_jpy==0){
$var_code_jpy_33="";
}else {
$var_code_jpy_33="33JPY".getTextValue($xcode_33_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_33_sgd) || $xcode_33_sgd=="" || $xcode_33_sgd==NULL || $xcode_33_sgd==0){
$var_code_sgd_33="";
}else {
$var_code_sgd_33="33SGD".getTextValue($xcode_33_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_33_usd) || $xcode_33_usd=="" || $xcode_33_usd==NULL || $xcode_33_usd==0){
$var_code_usd_33="";
}else {
$var_code_usd_33="33USD".getTextValue($xcode_33_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 34

$xcode_34_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_34_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_34_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_34_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_34_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_34_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J206')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_34_aud) || $xcode_34_aud=="" || $xcode_34_aud==NULL || $xcode_34_aud==0){
$var_code_aud_34="";
}else {
$var_code_aud_34="34AUD".getTextValue($xcode_34_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_34_eur) || $xcode_34_eur=="" || $xcode_34_eur==NULL || $xcode_34_eur==0){
$var_code_eur_34="";
}else {
$var_code_eur_34="34EUR".getTextValue($xcode_34_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_34_hkd) || $xcode_34_hkd=="" || $xcode_34_hkd==NULL || $xcode_34_hkd==0){
$var_code_hkd_34="";
}else {
$var_code_hkd_34="34HKD".getTextValue($xcode_34_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_34_jpy) || $xcode_34_jpy=="" || $xcode_34_jpy==NULL || $xcode_34_jpy==0){
$var_code_jpy_34="";
}else {
$var_code_jpy_34="34JPY".getTextValue($xcode_34_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_34_sgd) || $xcode_34_sgd=="" || $xcode_34_sgd==NULL || $xcode_34_sgd==0){
$var_code_sgd_34="";
}else {
$var_code_sgd_34="34SGD".getTextValue($xcode_34_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_34_usd) || $xcode_34_usd=="" || $xcode_34_usd==NULL || $xcode_34_usd==0){
$var_code_usd_34="";
}else {
$var_code_usd_34="34USD".getTextValue($xcode_34_usd).PHP_EOL;
$jml_baris_txt++; 
}

# code 35

$xcode_35_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_35_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_35_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_35_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_35_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_35_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J207')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_35_aud) || $xcode_35_aud=="" || $xcode_35_aud==NULL || $xcode_35_aud==0){
$var_code_aud_35="";
}else {
$var_code_aud_35="35AUD".getTextValue($xcode_35_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_35_eur) || $xcode_35_eur=="" || $xcode_35_eur==NULL || $xcode_35_eur==0){
$var_code_eur_35="";
}else {
$var_code_eur_35="35EUR".getTextValue($xcode_35_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_35_hkd) || $xcode_35_hkd=="" || $xcode_35_hkd==NULL || $xcode_35_hkd==0){
$var_code_hkd_35="";
}else {
$var_code_hkd_35="35HKD".getTextValue($xcode_35_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_35_jpy) || $xcode_35_jpy=="" || $xcode_35_jpy==NULL || $xcode_35_jpy==0){
$var_code_jpy_35="";
}else {
$var_code_jpy_35="35JPY".getTextValue($xcode_35_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_35_sgd) || $xcode_35_sgd=="" || $xcode_35_sgd==NULL || $xcode_35_sgd==0){
$var_code_sgd_35="";
}else {
$var_code_sgd_35="35SGD".getTextValue($xcode_35_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_35_usd) || $xcode_35_usd=="" || $xcode_35_usd==NULL || $xcode_35_usd==0){
$var_code_usd_35="";
}else {
$var_code_usd_35="35USD".getTextValue($xcode_35_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 39

$xcode_39_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_39_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_39_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_39_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_39_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_39_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J208')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_39_aud) || $xcode_39_aud=="" || $xcode_39_aud==NULL || $xcode_39_aud==0){
$var_code_aud_39="";
}else {
$var_code_aud_39="39AUD".getTextValue($xcode_39_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_39_eur) || $xcode_39_eur=="" || $xcode_39_eur==NULL || $xcode_39_eur==0){
$var_code_eur_39="";
}else {
$var_code_eur_39="39EUR".getTextValue($xcode_39_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_39_hkd) || $xcode_39_hkd=="" || $xcode_39_hkd==NULL || $xcode_39_hkd==0){
$var_code_hkd_39="";
}else {
$var_code_hkd_39="39HKD".getTextValue($xcode_39_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_39_jpy) || $xcode_39_jpy=="" || $xcode_39_jpy==NULL || $xcode_39_jpy==0){
$var_code_jpy_39="";
}else {
$var_code_jpy_39="39JPY".getTextValue($xcode_39_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_39_sgd) || $xcode_39_sgd=="" || $xcode_39_sgd==NULL || $xcode_39_sgd==0){
$var_code_sgd_39="";
}else {
$var_code_sgd_39="39SGD".getTextValue($xcode_39_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_39_usd) || $xcode_39_usd=="" || $xcode_39_usd==NULL || $xcode_39_usd==0){
$var_code_usd_39="";
}else {
$var_code_usd_39="39USD".getTextValue($xcode_39_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 51

$xcode_51_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_51_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_51_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_51_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_51_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_51_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J209')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_51_aud) || $xcode_51_aud=="" || $xcode_51_aud==NULL || $xcode_51_aud==0){
$var_code_aud_51="";
}else {
$var_code_aud_51="51AUD".getTextValue($xcode_51_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_51_eur) || $xcode_51_eur=="" || $xcode_51_eur==NULL || $xcode_51_eur==0){
$var_code_eur_51="";
}else {
$var_code_eur_51="51EUR".getTextValue($xcode_51_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_51_hkd) || $xcode_51_hkd=="" || $xcode_51_hkd==NULL || $xcode_51_hkd==0){
$var_code_hkd_51="";
}else {
$var_code_hkd_51="51HKD".getTextValue($xcode_51_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_51_jpy) || $xcode_51_jpy=="" || $xcode_51_jpy==NULL || $xcode_51_jpy==0){
$var_code_jpy_51="";
}else {
$var_code_jpy_51="51JPY".getTextValue($xcode_51_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_51_sgd) || $xcode_51_sgd=="" || $xcode_51_sgd==NULL || $xcode_51_sgd==0){
$var_code_sgd_51="";
}else {
$var_code_sgd_51="51SGD".getTextValue($xcode_51_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_51_usd) || $xcode_51_usd=="" || $xcode_51_usd==NULL || $xcode_51_usd==0){
$var_code_usd_51="";
}else {
$var_code_usd_51="51USD".getTextValue($xcode_51_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 52

$xcode_52_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_52_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_52_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_52_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_52_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_52_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J210')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_52_aud) || $xcode_52_aud=="" || $xcode_52_aud==NULL || $xcode_52_aud==0){
$var_code_aud_52="";
}else {
$var_code_aud_52="52AUD".getTextValue($xcode_52_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_52_eur) || $xcode_52_eur=="" || $xcode_52_eur==NULL || $xcode_52_eur==0){
$var_code_eur_52="";
}else {
$var_code_eur_52="52EUR".getTextValue($xcode_52_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_52_hkd) || $xcode_52_hkd=="" || $xcode_52_hkd==NULL || $xcode_52_hkd==0){
$var_code_hkd_52="";
}else {
$var_code_hkd_52="52HKD".getTextValue($xcode_52_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_52_jpy) || $xcode_52_jpy=="" || $xcode_52_jpy==NULL || $xcode_52_jpy==0){
$var_code_jpy_52="";
}else {
$var_code_jpy_52="52JPY".getTextValue($xcode_52_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_52_sgd) || $xcode_52_sgd=="" || $xcode_52_sgd==NULL || $xcode_52_sgd==0){
$var_code_sgd_52="";
}else {
$var_code_sgd_52="52SGD".getTextValue($xcode_52_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_52_usd) || $xcode_52_usd=="" || $xcode_52_usd==NULL || $xcode_52_usd==0){
$var_code_usd_52="";
}else {
$var_code_usd_52="52USD".getTextValue($xcode_52_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 61

$xcode_61_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_61_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_61_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_61_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_61_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_61_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J211')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_61_aud) || $xcode_61_aud=="" || $xcode_61_aud==NULL || $xcode_61_aud==0){
$var_code_aud_61="";
}else {
$var_code_aud_61="61AUD".getTextValue($xcode_61_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_61_eur) || $xcode_61_eur=="" || $xcode_61_eur==NULL || $xcode_61_eur==0){
$var_code_eur_61="";
}else {
$var_code_eur_61="61EUR".getTextValue($xcode_61_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_61_hkd) || $xcode_61_hkd=="" || $xcode_61_hkd==NULL || $xcode_61_hkd==0){
$var_code_hkd_61="";
}else {
$var_code_hkd_61="61HKD".getTextValue($xcode_61_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_61_jpy) || $xcode_61_jpy=="" || $xcode_61_jpy==NULL || $xcode_61_jpy==0){
$var_code_jpy_61="";
}else {
$var_code_jpy_61="61JPY".getTextValue($xcode_61_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_61_sgd) || $xcode_61_sgd=="" || $xcode_61_sgd==NULL || $xcode_61_sgd==0){
$var_code_sgd_61="";
}else {
$var_code_sgd_61="61SGD".getTextValue($xcode_61_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_61_usd) || $xcode_61_usd=="" || $xcode_61_usd==NULL || $xcode_61_usd==0){
$var_code_usd_61="";
}else {
$var_code_usd_61="61USD".getTextValue($xcode_61_usd).PHP_EOL;
$jml_baris_txt++; 
}


# code 65

$xcode_65_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_65_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_65_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_65_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_65_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_65_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J212')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_65_aud) || $xcode_65_aud=="" || $xcode_65_aud==NULL || $xcode_65_aud==0){
$var_code_aud_65="";
}else {
$var_code_aud_65="65AUD".getTextValue($xcode_65_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_65_eur) || $xcode_65_eur=="" || $xcode_65_eur==NULL || $xcode_65_eur==0){
$var_code_eur_65="";
}else {
$var_code_eur_65="65EUR".getTextValue($xcode_65_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_65_hkd) || $xcode_65_hkd=="" || $xcode_65_hkd==NULL || $xcode_65_hkd==0){
$var_code_hkd_65="";
}else {
$var_code_hkd_65="65HKD".getTextValue($xcode_65_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_65_jpy) || $xcode_65_jpy=="" || $xcode_65_jpy==NULL || $xcode_65_jpy==0){
$var_code_jpy_65="";
}else {
$var_code_jpy_65="65JPY".getTextValue($xcode_65_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_65_sgd) || $xcode_65_sgd=="" || $xcode_65_sgd==NULL || $xcode_65_sgd==0){
$var_code_sgd_65="";
}else {
$var_code_sgd_65="65SGD".getTextValue($xcode_65_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_65_usd) || $xcode_65_usd=="" || $xcode_65_usd==NULL || $xcode_65_usd==0){
$var_code_usd_65="";
}else {
$var_code_usd_65="65USD".getTextValue($xcode_65_usd).PHP_EOL;
$jml_baris_txt++; 
}



# code 67

$xcode_67_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_67_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_67_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_67_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_67_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_67_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J213')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_67_aud) || $xcode_67_aud=="" || $xcode_67_aud==NULL || $xcode_67_aud==0){
$var_code_aud_67="";
}else {
$var_code_aud_67="67AUD".getTextValue($xcode_67_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_67_eur) || $xcode_67_eur=="" || $xcode_67_eur==NULL || $xcode_67_eur==0){
$var_code_eur_67="";
}else {
$var_code_eur_67="67EUR".getTextValue($xcode_67_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_67_hkd) || $xcode_67_hkd=="" || $xcode_67_hkd==NULL || $xcode_67_hkd==0){
$var_code_hkd_67="";
}else {
$var_code_hkd_67="67HKD".getTextValue($xcode_67_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_67_jpy) || $xcode_67_jpy=="" || $xcode_67_jpy==NULL || $xcode_67_jpy==0){
$var_code_jpy_67="";
}else {
$var_code_jpy_67="67JPY".getTextValue($xcode_67_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_67_sgd) || $xcode_67_sgd=="" || $xcode_67_sgd==NULL || $xcode_67_sgd==0){
$var_code_sgd_67="";
}else {
$var_code_sgd_67="67SGD".getTextValue($xcode_67_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_67_usd) || $xcode_67_usd=="" || $xcode_67_usd==NULL || $xcode_67_usd==0){
$var_code_usd_67="";
}else {
$var_code_usd_67="67USD".getTextValue($xcode_67_usd).PHP_EOL;
$jml_baris_txt++; 
}

# code 69

$xcode_69_aud=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('E214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_69_eur=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('F214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_69_hkd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('G214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_69_jpy=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('H214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_69_sgd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('I214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));
$xcode_69_usd=str_replace(array('(', ')',','), "", $objPHPExcel->getActiveSheet()->getCell('J214')->getFormattedValue('#,##0,,;(#,##0,,);"-"'));


if (!isset($xcode_69_aud) || $xcode_69_aud=="" || $xcode_69_aud==NULL || $xcode_69_aud==0){
$var_code_aud_69="";
}else {
$var_code_aud_69="69AUD".getTextValue($xcode_69_aud).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_69_eur) || $xcode_69_eur=="" || $xcode_69_eur==NULL || $xcode_69_eur==0){
$var_code_eur_69="";
}else {
$var_code_eur_69="69EUR".getTextValue($xcode_69_eur).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_69_hkd) || $xcode_69_hkd=="" || $xcode_69_hkd==NULL || $xcode_69_hkd==0){
$var_code_hkd_69="";
}else {
$var_code_hkd_69="69HKD".getTextValue($xcode_69_hkd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_69_jpy) || $xcode_69_jpy=="" || $xcode_69_jpy==NULL || $xcode_69_jpy==0){
$var_code_jpy_69="";
}else {
$var_code_jpy_69="69JPY".getTextValue($xcode_69_jpy).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_69_sgd) || $xcode_69_sgd=="" || $xcode_69_sgd==NULL || $xcode_69_sgd==0){
$var_code_sgd_69="";
}else {
$var_code_sgd_69="69SGD".getTextValue($xcode_69_sgd).PHP_EOL;
$jml_baris_txt++; 
}
if (!isset($xcode_69_usd) || $xcode_69_usd=="" || $xcode_69_usd==NULL || $xcode_69_usd==0){
$var_code_usd_69="";
}else {
$var_code_usd_69="69USD".getTextValue($xcode_69_usd).PHP_EOL;
$jml_baris_txt++; 
}


$tanggal=date('dmY');
$tgl_filetxt=date('Ymd');

$header='48501'.$tanggal_header2.'401'.generateNominal2($jml_baris_txt).PHP_EOL;
$content = $var_modal_idr;
/*
// 10
$content .= $var_aktiva_aud;
$content .= $var_aktiva_eur;
$content .= $var_aktiva_hkd;
$content .= $var_aktiva_jpy;
$content .= $var_aktiva_sgd;
$content .= $var_aktiva_usd;
//15
$content .=$var_giro_aud;
$content .=$var_giro_eur;
$content .=$var_giro_hkd;
$content .=$var_giro_jpy;
$content .=$var_giro_sgd;
$content .=$var_giro_usd;
//29
$content .=$var_pasiva_aud;
$content .=$var_pasiva_eur;
$content .=$var_pasiva_hkd;
$content .=$var_pasiva_jpy;
$content .=$var_pasiva_sgd;
$content .=$var_pasiva_usd;


//31 == new
//$var_dkewajiban_e_aud
$content .=$var_dkewajiban_e_aud;
$content .=$var_dkewajiban_e_eur;
$content .=$var_dkewajiban_e_hkd;
$content .=$var_dkewajiban_e_jpy;
$content .=$var_dkewajiban_e_sgd;
$content .=$var_dkewajiban_e_usd;


//32 === new
//$var_ekewajiban_e_aud

$content .=$var_ekewajiban_e_aud;
$content .=$var_ekewajiban_e_eur;
$content .=$var_ekewajiban_e_hkd;
$content .=$var_ekewajiban_e_jpy;
$content .=$var_ekewajiban_e_sgd;
$content .=$var_ekewajiban_e_usd;





//32  ==== old
#$content .=$var_tagihan_a_aud;
#$content .=$var_tagihan_a_eur;
#$content .=$var_tagihan_a_hkd;
#$content .=$var_tagihan_a_jpy;
#$content .=$var_tagihan_a_sgd;
#$content .=$var_tagihan_a_usd;


//33
$content .=$var_tagihan_b_aud;
$content .=$var_tagihan_b_eur;
$content .=$var_tagihan_b_hkd;
$content .=$var_tagihan_b_jpy;
$content .=$var_tagihan_b_sgd;
$content .=$var_tagihan_b_usd;

//34 == new
//$var_fkewajiban_e_aud
$content .=$var_fkewajiban_e_aud;
$content .=$var_fkewajiban_e_eur;
$content .=$var_fkewajiban_e_hkd;
$content .=$var_fkewajiban_e_jpy;
$content .=$var_fkewajiban_e_sgd;
$content .=$var_fkewajiban_e_usd;

#$content .=$var_tagihan_e_aud;
#$content .=$var_tagihan_e_eur;
#$content .=$var_tagihan_e_hkd;
#$content .=$var_tagihan_e_jpy;
#$content .=$var_tagihan_e_sgd;
#$content .=$var_tagihan_e_usd;


//35
$content .=$var_akewajiban_e_aud;
$content .=$var_akewajiban_e_eur;
$content .=$var_akewajiban_e_hkd;
$content .=$var_akewajiban_e_jpy;
$content .=$var_akewajiban_e_sgd;
$content .=$var_akewajiban_e_usd;
//39
$content .=$var_bkewajiban_e_aud;
$content .=$var_bkewajiban_e_eur;
$content .=$var_bkewajiban_e_hkd;
$content .=$var_bkewajiban_e_jpy;
$content .=$var_bkewajiban_e_sgd;
$content .=$var_bkewajiban_e_usd;




//52
$content .=$var_ckewajiban_e_aud;
$content .=$var_ckewajiban_e_eur;
$content .=$var_ckewajiban_e_hkd;
$content .=$var_ckewajiban_e_jpy;
$content .=$var_ckewajiban_e_sgd;
$content .=$var_ckewajiban_e_usd;

//61
$content .=$var_tagihan_c_aud;
$content .=$var_tagihan_c_eur;
$content .=$var_tagihan_c_hkd;
$content .=$var_tagihan_c_jpy;
$content .=$var_tagihan_c_sgd;
$content .=$var_tagihan_c_usd;


//65
$content .=$var_kewajiban_d_aud;
$content .=$var_kewajiban_d_eur;
$content .=$var_kewajiban_d_hkd;
$content .=$var_kewajiban_d_jpy;
$content .=$var_kewajiban_d_sgd;
$content .=$var_kewajiban_d_usd;
//67
$content .=$var_kewajiban_c_aud;
$content .=$var_kewajiban_c_eur;
$content .=$var_kewajiban_c_hkd;
$content .=$var_kewajiban_c_jpy;
$content .=$var_kewajiban_c_sgd;
$content .=$var_kewajiban_c_usd;
//69
$content .=$var_tagihan_d_aud;
$content .=$var_tagihan_d_eur;
$content .=$var_tagihan_d_hkd;
$content .=$var_tagihan_d_jpy;
$content .=$var_tagihan_d_sgd;
$content .=$var_tagihan_d_usd;
*/

$content .= $var_code_aud_10;
$content .= $var_code_eur_10;
$content .= $var_code_hkd_10;
$content .= $var_code_jpy_10;
$content .= $var_code_sgd_10;
$content .= $var_code_usd_10;

$content .= $var_code_aud_15;
$content .= $var_code_eur_15;
$content .= $var_code_hkd_15;
$content .= $var_code_jpy_15;
$content .= $var_code_sgd_15;
$content .= $var_code_usd_15;

$content .= $var_code_aud_29;
$content .= $var_code_eur_29;
$content .= $var_code_hkd_29;
$content .= $var_code_jpy_29;
$content .= $var_code_sgd_29;
$content .= $var_code_usd_29;

$content .= $var_code_aud_31;
$content .= $var_code_eur_31;
$content .= $var_code_hkd_31;
$content .= $var_code_jpy_31;
$content .= $var_code_sgd_31;
$content .= $var_code_usd_31;

$content .= $var_code_aud_32;
$content .= $var_code_eur_32;
$content .= $var_code_hkd_32;
$content .= $var_code_jpy_32;
$content .= $var_code_sgd_32;
$content .= $var_code_usd_32;

$content .= $var_code_aud_33;
$content .= $var_code_eur_33;
$content .= $var_code_hkd_33;
$content .= $var_code_jpy_33;
$content .= $var_code_sgd_33;
$content .= $var_code_usd_33;

$content .= $var_code_aud_34;
$content .= $var_code_eur_34;
$content .= $var_code_hkd_34;
$content .= $var_code_jpy_34;
$content .= $var_code_sgd_34;
$content .= $var_code_usd_34;

$content .= $var_code_aud_35;
$content .= $var_code_eur_35;
$content .= $var_code_hkd_35;
$content .= $var_code_jpy_35;
$content .= $var_code_sgd_35;
$content .= $var_code_usd_35;

$content .= $var_code_aud_39;
$content .= $var_code_eur_39;
$content .= $var_code_hkd_39;
$content .= $var_code_jpy_39;
$content .= $var_code_sgd_39;
$content .= $var_code_usd_39;

$content .= $var_code_aud_51;
$content .= $var_code_eur_51;
$content .= $var_code_hkd_51;
$content .= $var_code_jpy_51;
$content .= $var_code_sgd_51;
$content .= $var_code_usd_51;

$content .= $var_code_aud_52;
$content .= $var_code_eur_52;
$content .= $var_code_hkd_52;
$content .= $var_code_jpy_52;
$content .= $var_code_sgd_52;
$content .= $var_code_usd_52;

$content .= $var_code_aud_61;
$content .= $var_code_eur_61;
$content .= $var_code_hkd_61;
$content .= $var_code_jpy_61;
$content .= $var_code_sgd_61;
$content .= $var_code_usd_61;

$content .= $var_code_aud_65;
$content .= $var_code_eur_65;
$content .= $var_code_hkd_65;
$content .= $var_code_jpy_65;
$content .= $var_code_sgd_65;
$content .= $var_code_usd_65;

$content .= $var_code_aud_67;
$content .= $var_code_eur_67;
$content .= $var_code_hkd_67;
$content .= $var_code_jpy_67;
$content .= $var_code_sgd_67;
$content .= $var_code_usd_67;

$content .= $var_code_aud_69;
$content .= $var_code_eur_69;
$content .= $var_code_hkd_69;
$content .= $var_code_jpy_69;
$content .= $var_code_sgd_69;
$content .= $var_code_usd_69;




$fp = fopen("download/LHA401_$label_txtfile.txt","wb");
fwrite($fp,$header.$content);
fclose($fp);

/*

echo $var_akewajiban_e_usd."<br>";
echo $var_akewajiban_e_aud."<br>";
echo $var_akewajiban_e_eur."<br>";
echo $var_akewajiban_e_hkd."<br>";
echo $var_akewajiban_e_jpy."<br>";
echo $var_akewajiban_e_sgd."<br>";
echo $var_akewajiban_e_usd."<br>";

echo $var_bkewajiban_e_aud."<br>";
echo $var_bkewajiban_e_eur."<br>";
echo $var_bkewajiban_e_hkd."<br>";
echo $var_bkewajiban_e_jpy."<br>";
echo $var_bkewajiban_e_sgd."<br>";
echo $var_bkewajiban_e_usd."<br>";

echo $var_ckewajiban_e_aud."<br>";
echo $var_ckewajiban_e_eur."<br>";
echo $var_ckewajiban_e_hkd."<br>";
echo $var_ckewajiban_e_jpy."<br>";
echo $var_ckewajiban_e_sgd."<br>";
echo $var_ckewajiban_e_usd."<br>";

echo $kewajiban_e_aud_a."<br>";
echo $kewajiban_e_eur_a."<br>";
echo $kewajiban_e_hkd_a."<br>";
echo $kewajiban_e_jpy_a."<br>";
echo $kewajiban_e_sgd_a."<br>";
echo $kewajiban_e_usd_a."<br>";


echo $kewajiban_e_aud_b."<br>";
echo $kewajiban_e_eur_b."<br>";
echo $kewajiban_e_hkd_b."<br>";
echo $kewajiban_e_jpy_b."<br>";
echo $kewajiban_e_sgd_b."<br>";
echo $kewajiban_e_usd_b."<br>";

echo $kewajiban_e_aud_c."<br>";
echo $kewajiban_e_eur_c."<br>";
echo $kewajiban_e_hkd_c."<br>";
echo $kewajiban_e_jpy_c."<br>";
echo $kewajiban_e_sgd_c."<br>";
echo $kewajiban_e_usd_c."<br>";
*/

?>
<div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> List NOP (Net Open Position)
                            </div>
                            <!--<div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>

                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div>-->
                        </div>
                        <div class="portlet-body">
                            <h4 class="font-blue-chambray"><b>PT Bank MNC Internasional, Tbk</b></h4>
                            
                            
                            <div class="tabbable-line">
                                <!--<ul class="nav nav-tabs ">
                                    <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        NOP (Net Open Position)</a>
                                    </li>
                                  
                                    
                                </ul>-->
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
											<b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/NOP_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br><b>(Dalam Jutaan Rupiah)</b> </div> </b></h5>

                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/LHA401_$label_txtfile.txt";?>" class="btn btn-sm red" download> Download txt <i class="fa fa-arrow-circle-o-down"></i> </a> </div></b>  

                                            
</br>
</br>
                                      
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               <tr class="active">
                                                <td width="5%" align="center"><b>No </b></td>
                                                <td width="25%" align="center"><b>Keterangan</b></td>
                                                <td width="10%" align="center"><b> AUD </b></td>
                                                <td width="10%" align="center"><b> EUR </b></td>
                                                <td width="10%" align="center"><b> HKD </b></td>
                                                <td width="10%" align="center"><b> JPY </b></td>
                                                <td width="10%" align="center"><b> SGD </b></td>
                                                <td width="10%" align="center"><b> USD </b></td>
                                                <td width="10%" align="center"><b> Jumlah </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td width="5%" style="font-size:12px">A. </td>
                                                <td width="25%" style="font-size:12px">Neraca</td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px">  </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px">  </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">1 </td>
                                                <td width="25%" style="font-size:12px"><b>Aktiva Valas </b></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">- Aktiva Valas tidak termasuk giro pada bank lain
 </td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">- Giro pada bank lain </td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">2</td>
                                                <td width="25%" style="font-size:12px"><b>Pasiva Valas </b></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">3</td>
                                                <td width="25%" style="font-size:12px"><b>Selisih Aktiva dan Pasiva Valas (A.1 - A.2)</b></td>
                                               <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">Selisih Aktiva dan Pasiva Valas (Nilai Absolut)

</td>
                                               
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">B</td>
                                                <td width="25%" style="font-size:12px"><b>Rekening Administratif</b>


</td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">1</td>
                                                <td width="25%" style="font-size:12px"><b>Rekening Administratif Tagihan Valas
</b>

</td>											
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K25')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">a. Kontrak pembelian forward


</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">b. Kontrak pembelian futures


</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">c. Kontrak penjualan put options (bank sebagai writter)


</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">d. Kontrak pembelian put options (bank sebagai holder, khusus back to back option)



</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">e. Rekening Administratif Kewajiban Valas diluar kontrak penjualan forward, futures, dan option

</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">2</td>
                                                <td width="25%" style="font-size:12px"><b>Rekening Administratif Kewajiban Valas</b>

</td>											

                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">a. Kontrak pembelian forward


</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">b. Kontrak pembelian futures


</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">c.  Kontrak penjualan call options (bank sebagai writter)



</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">d. Kontrak pembelian put options (bank sebagai    holder, khusus back to back option)





</td>
                                                 <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px">e. Rekening Administratif Kewajiban Valas diluar "   kontrak penjualan forward, futures, dan option
"


</td>
                                                 <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K41')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>


                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right">3</td>
                                                <td width="25%" style="font-size:12px"><b>Selisih Rekening Administratif (B.1 - B.2)


 </b> 



</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">C</td>
                                                <td width="25%" style="font-size:12px"><b>Posisi Devisa Netto per Valuta


 </b> 

</td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px" align="right"></td>
                                                <td width="25%" style="font-size:12px"><b>(A.3 + B.3)
</b>

</td>											<td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">D</td>
                                                <td width="25%" style="font-size:12px"><b>Posisi Devisa Netto


 </b> 


</td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                <td width="10%" style="font-size:12px"> </td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px"></td>
                                                <td width="25%" style="font-size:12px"><b>(Nilai Absolut C)


 </b> 


</td>											<td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K49')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">E</td>
                                                <td width="25%" style="font-size:12px"><b>Modal dalam Rupiah


 </b> 



</td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('E51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('F51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('G51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('H51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('I51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('J51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo $objPHPExcel->getActiveSheet()->getCell('K51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">F</td>
                                                <td width="25%" style="font-size:12px"><b>% PDN terhadap modal (A/E) Neraca

 </b> 




</td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_aud/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_eur/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_hkd/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_jpy/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_sgd/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php echo round(abs(($a1_min_a2_usd/$modal_nilai_fix)*100),2)."%"; ?></td>
                                                <td width="10%" style="font-size:12px"> <?php //echo round(abs((($a1_min_a2_tot+($b1_min_b2_tot-($a3_plus_b3_aud+$a3_plus_b3_eur+$a3_plus_b3_hkd+$a3_plus_b3_jpy+$a3_plus_b3_sgd+$a3_plus_b3_usd)))*100/$modal_nilai_fix)),2)."%"; 


       $pdn=round(abs(($a1_min_a2_aud/$modal_nilai_fix)*100),2)+round(abs(($a1_min_a2_eur/$modal_nilai_fix)*100),2)+round(abs(($a1_min_a2_hkd/$modal_nilai_fix)*100),2)+round(abs(($a1_min_a2_jpy/$modal_nilai_fix)*100),2)+round(abs(($a1_min_a2_sgd/$modal_nilai_fix)*100),2)+round(abs(($a1_min_a2_usd/$modal_nilai_fix)*100),2);
       echo $pdn."%";



                                                ?></td>
                                                </tr>
                                                <tr>
                                                <td width="5%" style="font-size:12px">G</td>
                                                <td width="25%" style="font-size:12px"><b>% PDN terhadap modal (D/E) Neraca & Rek. Adm.
 </b> 

</td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_aud+$b1_b2_aud)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_eur+$b1_b2_eur)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_hkd+$b1_b2_hkd)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_jpy+$b1_b2_jpy)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_sgd+$b1_b2_sgd)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php echo round(abs(($a1_min_a2_usd+$b1_b2_usd)/$modal_nilai_fix)*100,2)."%"; ?> </td>
                                                <td width="10%" style="font-size:12px"><?php //echo round(abs(($a1_min_a2_tot+($b1_min_b2_tot-($a3_plus_b3_aud+$a3_plus_b3_eur+$a3_plus_b3_hkd+$a3_plus_b3_jpy+$a3_plus_b3_sgd+$a3_plus_b3_usd)))/$modal_nilai_fix)*100,2)."%"; 


        $pdn2=round(abs(($a1_min_a2_aud+$b1_b2_aud)/$modal_nilai_fix)*100,2)+round(abs(($a1_min_a2_eur+$b1_b2_eur)/$modal_nilai_fix)*100,2)+round(abs(($a1_min_a2_hkd+$b1_b2_hkd)/$modal_nilai_fix)*100,2)+round(abs(($a1_min_a2_jpy+$b1_b2_jpy)/$modal_nilai_fix)*100,2)+round(abs(($a1_min_a2_sgd+$b1_b2_sgd)/$modal_nilai_fix)*100,2)+round(abs(($a1_min_a2_usd+$b1_b2_usd)/$modal_nilai_fix)*100,2);
        echo $pdn2."%";


                                                ?> </td>
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

