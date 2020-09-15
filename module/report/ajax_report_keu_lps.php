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
//logActivity("generate KPMM",date('Y_m_d_H_i_s'));

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

############################### Query Nearaca PL ################################################################################################################# 
$query_neraca="SELECT SUM(Nilai) AS Jumlah_Nominal FROM (
SELECT a.kodegl,SUM(a.nominal)/1000000 AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_LKBLPS c ON c.LKBLPS_Level_3 = b.LKBLPS_Level_3
WHERE a.DataDate='$curr_tgl' ";

 
$neraca_add=" GROUP BY a.kodegl ,b.LKBLPS_Level_3,a.kodeproduct,a.kodecabang )AS tabel1 ";



// 10 LKBLPS101000001 Kas
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000001' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c10=$row_neraca['Jumlah_Nominal'];

//11 LKBLPS101000002 Penempatan pada Bank Indonesia
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c11=$row_neraca['Jumlah_Nominal'];

//LKBLPS101000003 Penempatan pada bank lain
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000003' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c12=$row_neraca['Jumlah_Nominal'];

//LKBLPS101000004 Tagihan spot dan derivatif
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c13=$row_neraca['Jumlah_Nominal'];

//15  LKBLPS101000007   a. Diukur pada nilai wajar melalui laporan laba/rugi
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000007' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c15=$row_neraca['Jumlah_Nominal'];

//16  LKBLPS101000009   b. Tersedia untuk dijual
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000009' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c16=$row_neraca['Jumlah_Nominal'];

//17  LKBLPS101000010   c. Dimiliki hingga jatuh tempo
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000010' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c17=$row_neraca['Jumlah_Nominal'];

//18  LKBLPS101000011  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000011' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c18=$row_neraca['Jumlah_Nominal'];

//19  LKBLPS101000012   
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000012' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c19=$row_neraca['Jumlah_Nominal'];

//20  LKBLPS101000013  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000013' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c20=$row_neraca['Jumlah_Nominal'];


//21  LKBLPS101000014 Tagihan akseptasi
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000014' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c21=$row_neraca['Jumlah_Nominal'];

//23  LKBLPS101000016 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000016' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c23=$row_neraca['Jumlah_Nominal'];

//24  LKBLPS101000017 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000017' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c24=$row_neraca['Jumlah_Nominal'];

//25  LKBLPS101000018 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000018' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c25=$row_neraca['Jumlah_Nominal'];

//26  LKBLPS101000019   d. Pinjaman diberikan dan piutang
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000019' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c26=$row_neraca['Jumlah_Nominal'];

//27  LKBLPS101000020
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000020' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c27=$row_neraca['Jumlah_Nominal'];

//28  LKBLPS101000021
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000021' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c28=$row_neraca['Jumlah_Nominal'];

//30  LKBLPS101000023
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000023' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c30=$row_neraca['Jumlah_Nominal'];

//31  LKBLPS101000024   b. Kredit
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000024' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c31=$row_neraca['Jumlah_Nominal'];

//32
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000025' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c32=$row_neraca['Jumlah_Nominal'];

//33  LKBLPS101000026 Aset tidak berwujud
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000026' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c33=$row_neraca['Jumlah_Nominal'];

//34  LKBLPS101000027 Akumulasi amortisasi aset tidak berwujud -/-
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000027' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c34=$row_neraca['Jumlah_Nominal'];

//35  LKBLPS101000028 Aset tetap dan inventaris
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000028' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c35=$row_neraca['Jumlah_Nominal'];

//36  LKBLPS101000029 Akumulasi penyusutan aset tetap dan inventaris -/-
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000029' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c36=$row_neraca['Jumlah_Nominal'];

//38  LKBLPS101000031
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000031' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c38=$row_neraca['Jumlah_Nominal'];

//39  LKBLPS101000032   b. Aset yang diambil alih
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000032' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c39=$row_neraca['Jumlah_Nominal'];

//40 LKBLPS101000033   c. Rekening tunda
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000033' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c40=$row_neraca['Jumlah_Nominal'];

//41
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000034' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c41=$row_neraca['Jumlah_Nominal'];

//42  *Note : value is showing, but not 0 LKBLPS101000035   
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000035' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c42=$row_neraca['Jumlah_Nominal'];

//43  LKBLPS101000036  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000036' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c43=$row_neraca['Jumlah_Nominal'];

//44  LKBLPS101000037
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000037' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c44=$row_neraca['Jumlah_Nominal'];

//45  LKBLPS101000038 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000038' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c45=$row_neraca['Jumlah_Nominal'];

//46    LKBLPS101000039
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000039' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c46=$row_neraca['Jumlah_Nominal'];
//47  *Note : value is showing, but not match 382593.062484 LKBLPS101000040  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS101000040' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c47=$row_neraca['Jumlah_Nominal'];
//51    LKBLPS102000001
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000001' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c51=abs($row_neraca['Jumlah_Nominal']);
//52  *Note : value is showing, but not match -602363.989447  LKBLPS102000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c52=abs($row_neraca['Jumlah_Nominal']);
//53    LKBLPS102000003
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000003' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c53=abs($row_neraca['Jumlah_Nominal']);

//54    LKBLPS102000004
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c54=abs($row_neraca['Jumlah_Nominal']);

//55    LKBLPS102000005
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000005' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c55=abs($row_neraca['Jumlah_Nominal']);

//56  *Note : value is showing, but not match -493422.532927. Perlu konfirm parameter yg ditarik yg mana  LKBLPS102000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000006' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c56=abs($row_neraca['Jumlah_Nominal']);
//57    LKBLPS102000007  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000007' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c57=abs($row_neraca['Jumlah_Nominal']);

//58      LKBLPS102000008
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c58=abs($row_neraca['Jumlah_Nominal']);

//59    LKBLPS102000009
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000009' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c59=abs($row_neraca['Jumlah_Nominal']);
//60    LKBLPS102000010
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000010' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c60=abs($row_neraca['Jumlah_Nominal']);
//61    LKBLPS102000011 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000011' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c61=abs($row_neraca['Jumlah_Nominal']);

//62    LKBLPS102000012 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000012' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c62=abs($row_neraca['Jumlah_Nominal']);

//64    LKBLPS102000014
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000014' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c64=abs($row_neraca['Jumlah_Nominal']);

//65    LKBLPS102000015
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000015' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c65=abs($row_neraca['Jumlah_Nominal']);

//66    LKBLPS102000016 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000016' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c66=abs($row_neraca['Jumlah_Nominal']);

//67  *Note : value is showing, but not match -676.610890. Perlu konfirm parameter yg ditarik yg mana LKBLPS102000017
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS102000017' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c67=abs($row_neraca['Jumlah_Nominal']);

//70 LKBLPS103000004
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c70=abs($row_neraca['Jumlah_Nominal']);

//71 LKBLPS103000005
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000005' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c71=$row_neraca['Jumlah_Nominal'];

//72 LKBLPS103000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000006' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c72=abs($row_neraca['Jumlah_Nominal']);

//74  LKBLPS103000008
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c74=abs($row_neraca['Jumlah_Nominal']);

//75 LKBLPS103000004
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c75=abs($row_neraca['Jumlah_Nominal']);

//76 LKBLPS103000005
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c76=abs($row_neraca['Jumlah_Nominal']);

//77 LKBLPS103000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c77=abs($row_neraca['Jumlah_Nominal']);

//81 LKBLPS103000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000016' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c81=abs($row_neraca['Jumlah_Nominal']);

//85 LKBLPS103000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000022' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c85=abs($row_neraca['Jumlah_Nominal']);

//90 LKBLPS103000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000029' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c90=abs($row_neraca['Jumlah_Nominal']);

//92 LKBLPS103000031
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000031' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c92=abs($row_neraca['Jumlah_Nominal']);

//93 LKBLPS103000033
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000033' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c93=abs($row_neraca['Jumlah_Nominal']);

//95 LKBLPS103000034
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000034' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c95=$row_neraca['Jumlah_Nominal'];

//96 LKBLPS103000035
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000035' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c96=abs($row_neraca['Jumlah_Nominal']);

#108 *Note : value is showing, but not match -139809.014676  LKBLPS201000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c108=abs($row_neraca['Jumlah_Nominal']);

//109 *Note : value is showing, but not match 0. parameter tidak tersedia di Referensi_gl_02  LKBLPS201000003 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000003' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c109=$row_neraca['Jumlah_Nominal'];

//111
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c111=$row_neraca['Jumlah_Nominal'];

//112  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000003' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c112=$row_neraca['Jumlah_Nominal'];

//118 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c118=abs($row_neraca['Jumlah_Nominal']);
//119 LKBLPS201000009 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000009' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c119=$row_neraca['Jumlah_Nominal']; 
//120 LKBLPS201000010   
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000010' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c120=abs($row_neraca['Jumlah_Nominal']);
//121 LKBLPS201000011 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000011' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c121=$row_neraca['Jumlah_Nominal'];
//122   LKBLPS201000012
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000012' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c122=$row_neraca['Jumlah_Nominal'];
//echo($query_neraca.$neraca_lps.$neraca_add);
//die();

//124 *Note : value is showing, but not match -3902.535870  LKBLPS201000017 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000014' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c124=abs($row_neraca['Jumlah_Nominal']);

//125   LKBLPS201000020
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000015' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c125=$row_neraca['Jumlah_Nominal'];
#126  LKBLPS201000016
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000016' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
 $c126=$row_neraca['Jumlah_Nominal'];

#127   LKBLPS201000017
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000017' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c127=abs($row_neraca['Jumlah_Nominal']);

//128   LKBLPS201000019
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000019' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c128=$row_neraca['Jumlah_Nominal'];
#129 LKBLPS201000018 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000018' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c129=$row_neraca['Jumlah_Nominal'];

#130   LKBLPS201000020
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000020' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c130=abs($row_neraca['Jumlah_Nominal']);

//132 LKBLPS201000022
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS201000022' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c132=abs($row_neraca['Jumlah_Nominal']);


//136 LKBLPS202000006
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000006' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c136=$row_neraca['Jumlah_Nominal'];

//137 LKBLPS202000007
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000007' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c137=$row_neraca['Jumlah_Nominal'];

//138 LKBLPS202000008
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c138=abs($row_neraca['Jumlah_Nominal']);

//139 LKBLPS202000009
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000009' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c139=$row_neraca['Jumlah_Nominal'];

//140 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000010' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c140=$row_neraca['Jumlah_Nominal'];

//142  
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000012' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c142=$row_neraca['Jumlah_Nominal'];

//143 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000013' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c143=$row_neraca['Jumlah_Nominal'];

//144 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000014' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c144=$row_neraca['Jumlah_Nominal'];

//145 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000015' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c145=$row_neraca['Jumlah_Nominal'];

//147 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000017' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c147=$row_neraca['Jumlah_Nominal'];

//148 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000018' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c148=$row_neraca['Jumlah_Nominal'];


//150 *Note : value is showing, but not match 16878.171092  LKBLPS202000025
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000020' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c150=$row_neraca['Jumlah_Nominal'];
//151 *Note : value is showing, but not match 52.526886 LKBLPS202000026
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000021' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c151=$row_neraca['Jumlah_Nominal'];
//152 *Note : value is showing, but not match 16687.515722  LKBLPS202000027
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000022' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c152=$row_neraca['Jumlah_Nominal'];
//153   LKBLPS203000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000023' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c153=$row_neraca['Jumlah_Nominal'];
//154
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000024' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c154=$row_neraca['Jumlah_Nominal'];
//155
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000025' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c155=$row_neraca['Jumlah_Nominal'];
//156
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000026' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c156=$row_neraca['Jumlah_Nominal'];
//157
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS202000027' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c157=$row_neraca['Jumlah_Nominal'];


//162   LKBLPS203000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS203000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c162=abs($row_neraca['Jumlah_Nominal']);
//163 #N/A    LKBLPS203000003
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS203000003' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c163=$row_neraca['Jumlah_Nominal'];
//164 #N/A    LKBLPS203000004
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS203000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c164=$row_neraca['Jumlah_Nominal'];

//169 LKBLPS204000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS204000002' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c169=$row_neraca['Jumlah_Nominal'];

//175 LKBLPS206000004
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS206000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c175=$row_neraca['Jumlah_Nominal'];

//187 LKBLPS204000002
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS206000023' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c187=$row_neraca['Jumlah_Nominal'];

#198	 LKBLPS103000039 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000039' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c198=$row_neraca['Jumlah_Nominal'];
#199	 LKBLPS103000040 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000040' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c199=$row_neraca['Jumlah_Nominal'];
#200	 LKBLPS103000041 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000041' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c200=$row_neraca['Jumlah_Nominal'];
#201	 LKBLPS103000042 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS103000042' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c201=$row_neraca['Jumlah_Nominal'];
#207	 LKBLPS104000004 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000004' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c207=$row_neraca['Jumlah_Nominal'];
#208	 LKBLPS104000005 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000005' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c208=$row_neraca['Jumlah_Nominal'];
#210	 LKBLPS104000007 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000007' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c210=$row_neraca['Jumlah_Nominal'];
#211	 LKBLPS104000008 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000008' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c211=$row_neraca['Jumlah_Nominal'];
#213	 LKBLPS104000010 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000010' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c213=abs($row_neraca['Jumlah_Nominal']);
#214	 LKBLPS104000011 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000011' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c214=$row_neraca['Jumlah_Nominal'];
#217	 LKBLPS104000014 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000014' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c117=$row_neraca['Jumlah_Nominal'];
#218	 LKBLPS104000015 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000015' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c218=$row_neraca['Jumlah_Nominal'];
#220	 LKBLPS104000017 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000017' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c220=$row_neraca['Jumlah_Nominal'];
#221	 LKBLPS104000018 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000018' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c221=$row_neraca['Jumlah_Nominal'];
#223	 LKBLPS104000020 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000020' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c223=abs($row_neraca['Jumlah_Nominal']);
#224	 LKBLPS104000021 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000021' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c224=$row_neraca['Jumlah_Nominal'];
#225	 LKBLPS104000022 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000022' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c225=abs($row_neraca['Jumlah_Nominal']);
#230	 LKBLPS104000027 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000027' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c230=$row_neraca['Jumlah_Nominal'];
#231	 LKBLPS104000028 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000028' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c231=$row_neraca['Jumlah_Nominal'];
#233	 LKBLPS104000030 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000030' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c233=$row_neraca['Jumlah_Nominal'];
#234	 LKBLPS104000031 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000031' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c234=$row_neraca['Jumlah_Nominal'];
#235	 LKBLPS104000032 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000032' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c235=$row_neraca['Jumlah_Nominal'];
#239	 LKBLPS104000035 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000035' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c239=$row_neraca['Jumlah_Nominal'];
#240	 LKBLPS104000036 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000036' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c240=$row_neraca['Jumlah_Nominal'];
#241	 LKBLPS104000037 
$neraca_lps=" AND b.LKBLPS_Level_3 ='LKBLPS104000037' ";
$result_neraca=odbc_exec($connection2, $query_neraca.$neraca_lps.$neraca_add);
$row_neraca=odbc_fetch_array($result_neraca);
$c241=abs($row_neraca['Jumlah_Nominal']);

############################################################################################################################################################ 


##### QUERY KPMM ########################################################################################################################################## 
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM (
SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_KPMM c ON c.KPMM_Level_3 = b.KPMM_Level_3
WHERE a.DataDate='$curr_tgl'  ";



$q_add2="  GROUP BY a.kodegl ,b.KPMM_Level_3 )AS tabel1 ";

 $curr_mon=date('n',strtotime($tanggal));
 $curr_year=date('Y',strtotime($tanggal));
# +++++++++++ QUERY ATMR ++++++++++++++
$query_atmr= " select * from Master_ATMR WHERE  Month(DataDate)='$curr_mon' and Year(DataDate)='$curr_year'  ";
//echo $query_atmr;
//die();

$result_atmr=odbc_exec($connection2, $query_atmr);
$rowAtmr=odbc_fetch_array($result_atmr);
$atmr_kredit=$rowAtmr['ATMR_Kredit'];
$atmr_pasar=$rowAtmr['ATMR_Pasar'];
$atmr_operasional=$rowAtmr['ATMR_Operasional'];


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







#--II : 3 Cadangan Umum PPA atas aset produktif yang wajib dibentuk
$query=" SELECT SUM (NILAI)/1000000 AS Jumlah_Nominal FROM 
(
 SELECT SUM(PPA) AS NILAI FROM DM_KPMM_PPA_LBU WITH (NOLOCK)
 WHERE JENIS='PPA ASET PRODUKTIF' AND KOL='1' AND DATADATE='$curr_tgl'
) AS TABEL1 ";

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

$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m33=abs($row['Total_Nilai']);

#--1.2.2.2.6 PPA aset non produktif yang wajib dibentuk (*Note : dibuat positive)
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
$result=odbc_exec($connection2, $query);
$row=odbc_fetch_array($result);
$m35=abs($row['Total_Nilai']);





// Create new PHPExcel object
$objPHPExcel = new PHPExcel();



// SHEET KE 1 ======================================================================================

$objPHPExcel->setActiveSheetIndex(0);

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

 $objPHPExcel->getActiveSheet()->getStyle('A2:C8')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A48:C48')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A97:C106')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A190:C196')->applyFromArray($styleArrayFontBold);
// $objPHPExcel->getActiveSheet()->getStyle('A2:N3')->applyFromArray($styleArrayAlignment1);
// $objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayAlignment1);

// $objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



 $objPHPExcel->getActiveSheet()->getStyle('B1:B8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('B1:B8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('B48')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B48')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B97')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B97')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('B99:B104')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B99:B104')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('B190:B195')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B190:B195')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);




//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B8:C97')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B105:C188')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B196:C241')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('B2:C5')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B6:C8')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B99:C101')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B102:C104')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B190:C192')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B193:C195')->applyFromArray($styleArrayBorder2);
//FILL COLOR

 $objPHPExcel->getActiveSheet()->getStyle('A1:D245')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');


//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(75);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(5);



// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);

// $objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:N1');






$objPHPExcel->getActiveSheet()->setCellValue('B3', 'NERACA BANK UMUM');
$objPHPExcel->getActiveSheet()->setCellValue('B4', "(Dalam jutaan rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('B6', 'PT. BANK MNC INTERNASIONAL, TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B7', '485000');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'POS-POS');
$objPHPExcel->getActiveSheet()->setCellValue('C8', "$label_tgl");

$objPHPExcel->getActiveSheet()->setCellValue('B9', 'ASET');


#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Kas');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Penempatan pada Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Penempatan pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Tagihan spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B15', ' a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B16', ' b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('B17', ' c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('B18', ' d. Pinjaman diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Surat berharga yang dijual dengan janji dibeli kembali (repo)');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Tagihan atas surat berharga yang dibeli dengan janji dijual kembali (reverse repo)');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Tagihan akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B23', ' a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B24', ' b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('B25', ' c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('B26', ' d. Pinjaman diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'Pembiayaan Syariah');
$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Penyertaan');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Cadangan kerugian penurunan nilai aset keuangan -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B30', ' a. Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B31', ' b. Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B32', ' c. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'Aset tidak berwujud');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Akumulasi amortisasi aset tidak berwujud -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Aset tetap dan inventaris');
$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Akumulasi penyusutan aset tetap dan inventaris -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Aset Non Produktif');
$objPHPExcel->getActiveSheet()->setCellValue('B38', ' a. Properti terbengkalai');
$objPHPExcel->getActiveSheet()->setCellValue('B39', ' b. Aset yang diambil alih');
$objPHPExcel->getActiveSheet()->setCellValue('B40', ' c. Rekening tunda');
$objPHPExcel->getActiveSheet()->setCellValue('B41', ' d. Aset antar kantor:');
$objPHPExcel->getActiveSheet()->setCellValue('B42', '    a. Melakukan kegiatan operasional di Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B43', '    b. Melakukan kegiatan operasional di luar Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B44', 'Cadangan kerugian penurunan nilai aset lainnya -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'Sewa pembiayaan');
$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Aset pajak tangguhan');
$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Aset Lainnya');

$objPHPExcel->getActiveSheet()->setCellValue('B48', 'TOTAL ASET');

$objPHPExcel->getActiveSheet()->setCellValue('B50', 'KEWAJIBAN DAN MODAL');
$objPHPExcel->getActiveSheet()->setCellValue('B51', 'Giro');
$objPHPExcel->getActiveSheet()->setCellValue('B52', 'Tabungan');
$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Simpanan berjangka');
$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Dana investasi revenue sharing');
$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Pinjaman Dari Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Pinjaman Dari bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('B57', 'Liabilitas spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B58', 'Utang surat berharga yang dijual dengan janji dibeli kembali (repo)');
$objPHPExcel->getActiveSheet()->setCellValue('B59', 'Utang akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('B60', 'Surat berharga yang diterbitkan');
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'Pinjaman yang diterima');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Setoran jaminan');
$objPHPExcel->getActiveSheet()->setCellValue('B63', 'Kewajiban antarkantor');
$objPHPExcel->getActiveSheet()->setCellValue('B64', ' a. Melakukan kegiatan operasional di Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B65', ' b. Melakukan kegiatan operasional di luar Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B66', 'Kewajiban pajak tangguhan');
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'Kewajiban Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'Dana investasi (profit sharing)');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'Modal disetor');
$objPHPExcel->getActiveSheet()->setCellValue('B70', ' a. Modal dasar');
$objPHPExcel->getActiveSheet()->setCellValue('B71', ' b. Modal yang belum disetor -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B72', ' c. Saham yang dibeli kembali (treasury stock) -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B73', 'Tambahan modal disetor');
$objPHPExcel->getActiveSheet()->setCellValue('B74', ' a. Agio');
$objPHPExcel->getActiveSheet()->setCellValue('B75', ' b. Disagio -/-');
$objPHPExcel->getActiveSheet()->setCellValue('B76', ' c. Modal sumbangan');
$objPHPExcel->getActiveSheet()->setCellValue('B77', ' d. Dana setoran modal');
$objPHPExcel->getActiveSheet()->setCellValue('B78', ' e. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B79', 'Pendapatan (kerugian) komprehensif lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B80', ' a. Penyesuaian akibat penjabaran Laporan Keuangan dalam mata uang asing');

$objPHPExcel->getActiveSheet()->setCellValue('B81', ' b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('B82', ' c. Bagian efektif lindung arus kas');
$objPHPExcel->getActiveSheet()->setCellValue('B83', ' d. Keuntungan revaluasi asset tetap');
$objPHPExcel->getActiveSheet()->setCellValue('B84', ' e. Bagian pendapatan komprehensif lain dari entitas asosiasi');
$objPHPExcel->getActiveSheet()->setCellValue('B85', ' f. Keuntungan (kerugian) aktuarial program imbalan pasti ');

$objPHPExcel->getActiveSheet()->setCellValue('B86', ' g. Pajak penghasilan terkait dengan penghasilan komprehensif lain');
$objPHPExcel->getActiveSheet()->setCellValue('B87', ' h. Lainnya');

$objPHPExcel->getActiveSheet()->setCellValue('B88', 'Selisih kuasi reorganisasi');
$objPHPExcel->getActiveSheet()->setCellValue('B89', 'Selisih restrukturisasi entitas sepengendali');
$objPHPExcel->getActiveSheet()->setCellValue('B90', 'Ekuitas Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B91', 'Cadangan');
$objPHPExcel->getActiveSheet()->setCellValue('B92', ' a. Cadangan umum');
$objPHPExcel->getActiveSheet()->setCellValue('B93', ' b. Cadangan tujuan');
$objPHPExcel->getActiveSheet()->setCellValue('B94', 'Laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B95', ' a. Tahun-tahun lalu');
$objPHPExcel->getActiveSheet()->setCellValue('B96', ' b. Tahun berjalan');
$objPHPExcel->getActiveSheet()->setCellValue('B97', 'TOTAL KEWAJIBAN DAN MODAL');





#### VALUE FOR NERACA ------------------------
// for ($i=10; $i <=91 ; $i++) { 
//   $objPHPExcel->getActiveSheet()->setCellValue("C$i", "$"."c".$i);
// }

$objPHPExcel->getActiveSheet()->setCellValue("C10", $c10);
$objPHPExcel->getActiveSheet()->setCellValue("C11", $c11);
$objPHPExcel->getActiveSheet()->setCellValue("C12", $c12);
$objPHPExcel->getActiveSheet()->setCellValue("C13", $c13);
$objPHPExcel->getActiveSheet()->setCellValue("C14", $c14);
$objPHPExcel->getActiveSheet()->setCellValue("C15", $c15);
$objPHPExcel->getActiveSheet()->setCellValue("C16", $c16);
$objPHPExcel->getActiveSheet()->setCellValue("C17", $c17);
$objPHPExcel->getActiveSheet()->setCellValue("C18", $c18);
$objPHPExcel->getActiveSheet()->setCellValue("C19", $c19);
$objPHPExcel->getActiveSheet()->setCellValue("C20", $c20);
$objPHPExcel->getActiveSheet()->setCellValue("C21", $c21);
$objPHPExcel->getActiveSheet()->setCellValue("C22", $c22);
$objPHPExcel->getActiveSheet()->setCellValue("C23", $c23);
$objPHPExcel->getActiveSheet()->setCellValue("C24", $c24);
$objPHPExcel->getActiveSheet()->setCellValue("C25", $c25);
$objPHPExcel->getActiveSheet()->setCellValue("C26", $c26);
$objPHPExcel->getActiveSheet()->setCellValue("C27", $c27);
$objPHPExcel->getActiveSheet()->setCellValue("C28", $c28);
$objPHPExcel->getActiveSheet()->setCellValue("C29", $c29);
$objPHPExcel->getActiveSheet()->setCellValue("C30", $c30);
$objPHPExcel->getActiveSheet()->setCellValue("C31", $c31);
$objPHPExcel->getActiveSheet()->setCellValue("C32", $c32);
$objPHPExcel->getActiveSheet()->setCellValue("C33", $c33);
$objPHPExcel->getActiveSheet()->setCellValue("C34", $c34);
$objPHPExcel->getActiveSheet()->setCellValue("C35", $c35);
$objPHPExcel->getActiveSheet()->setCellValue("C36", $c36);
$objPHPExcel->getActiveSheet()->setCellValue("C37", $c37);
$objPHPExcel->getActiveSheet()->setCellValue("C38", $c38);
$objPHPExcel->getActiveSheet()->setCellValue("C39", $c39);
$objPHPExcel->getActiveSheet()->setCellValue("C40", $c40);
$objPHPExcel->getActiveSheet()->setCellValue("C41", $c41);
$objPHPExcel->getActiveSheet()->setCellValue("C42", $c42);
$objPHPExcel->getActiveSheet()->setCellValue("C43", $c43);
$objPHPExcel->getActiveSheet()->setCellValue("C44", $c44);
$objPHPExcel->getActiveSheet()->setCellValue("C45", $c45);
$objPHPExcel->getActiveSheet()->setCellValue("C46", $c46);
$objPHPExcel->getActiveSheet()->setCellValue("C47", $c47);
$objPHPExcel->getActiveSheet()->setCellValue("C48", $c48);
$objPHPExcel->getActiveSheet()->setCellValue("C49", $c49);
$objPHPExcel->getActiveSheet()->setCellValue("C50", $c50);
$objPHPExcel->getActiveSheet()->setCellValue("C51", $c51);
$objPHPExcel->getActiveSheet()->setCellValue("C52", $c52);
$objPHPExcel->getActiveSheet()->setCellValue("C53", $c53);
$objPHPExcel->getActiveSheet()->setCellValue("C54", $c54);
$objPHPExcel->getActiveSheet()->setCellValue("C55", $c55);
$objPHPExcel->getActiveSheet()->setCellValue("C56", $c56);
$objPHPExcel->getActiveSheet()->setCellValue("C57", $c57);
$objPHPExcel->getActiveSheet()->setCellValue("C58", $c58);
$objPHPExcel->getActiveSheet()->setCellValue("C59", $c59);
$objPHPExcel->getActiveSheet()->setCellValue("C60", $c60);
$objPHPExcel->getActiveSheet()->setCellValue("C61", $c61);
$objPHPExcel->getActiveSheet()->setCellValue("C62", $c62);
$objPHPExcel->getActiveSheet()->setCellValue("C63", $c63);
$objPHPExcel->getActiveSheet()->setCellValue("C64", $c64);
$objPHPExcel->getActiveSheet()->setCellValue("C65", $c65);
$objPHPExcel->getActiveSheet()->setCellValue("C66", $c66);
$objPHPExcel->getActiveSheet()->setCellValue("C67", $c67);
$objPHPExcel->getActiveSheet()->setCellValue("C68", $c68);
$objPHPExcel->getActiveSheet()->setCellValue("C69", $c69);
$objPHPExcel->getActiveSheet()->setCellValue("C70", $c70);
$objPHPExcel->getActiveSheet()->setCellValue("C71", $c71);
$objPHPExcel->getActiveSheet()->setCellValue("C72", $c72);
$objPHPExcel->getActiveSheet()->setCellValue("C73", $c73);
$objPHPExcel->getActiveSheet()->setCellValue("C74", $c74);
$objPHPExcel->getActiveSheet()->setCellValue("C75", $c75);
$objPHPExcel->getActiveSheet()->setCellValue("C76", $c76);
$objPHPExcel->getActiveSheet()->setCellValue("C77", $c77);
$objPHPExcel->getActiveSheet()->setCellValue("C78", $c78);
$objPHPExcel->getActiveSheet()->setCellValue("C79", $c79);
$objPHPExcel->getActiveSheet()->setCellValue("C80", $c80);
$objPHPExcel->getActiveSheet()->setCellValue("C81", $c81);
$objPHPExcel->getActiveSheet()->setCellValue("C82", $c82);
$objPHPExcel->getActiveSheet()->setCellValue("C83", $c83);
$objPHPExcel->getActiveSheet()->setCellValue("C84", $c84);
$objPHPExcel->getActiveSheet()->setCellValue("C85", $c85);
$objPHPExcel->getActiveSheet()->setCellValue("C86", $c86);
$objPHPExcel->getActiveSheet()->setCellValue("C87", $c87);
$objPHPExcel->getActiveSheet()->setCellValue("C88", $c88);
$objPHPExcel->getActiveSheet()->setCellValue("C89", $c89);
$objPHPExcel->getActiveSheet()->setCellValue("C90", $c90);
$objPHPExcel->getActiveSheet()->setCellValue("C91", $c91);
$objPHPExcel->getActiveSheet()->setCellValue("C92", $c92);
$objPHPExcel->getActiveSheet()->setCellValue("C93", $c93);
$objPHPExcel->getActiveSheet()->setCellValue("C94", $c94);
$objPHPExcel->getActiveSheet()->setCellValue("C95", $c95);
$objPHPExcel->getActiveSheet()->setCellValue("C96", $c96);
$objPHPExcel->getActiveSheet()->setCellValue("C97", $c97);
$objPHPExcel->getActiveSheet()->setCellValue("C98", $c98);
//103-164

// $objPHPExcel->getActiveSheet()->setCellValue("C101", $c101);
// $objPHPExcel->getActiveSheet()->setCellValue("C102", $c102);
// $objPHPExcel->getActiveSheet()->setCellValue("C103", $c103);
// $objPHPExcel->getActiveSheet()->setCellValue("C104", $c104);
// $objPHPExcel->getActiveSheet()->setCellValue("C105", $c105);
// $objPHPExcel->getActiveSheet()->setCellValue("C106", $c106);
$objPHPExcel->getActiveSheet()->setCellValue("C107", $c107);
$objPHPExcel->getActiveSheet()->setCellValue("C108", $c108);
$objPHPExcel->getActiveSheet()->setCellValue("C109", $c109);
$objPHPExcel->getActiveSheet()->setCellValue("C110", $c110);
$objPHPExcel->getActiveSheet()->setCellValue("C111", $c111);
$objPHPExcel->getActiveSheet()->setCellValue("C112", $c112);
$objPHPExcel->getActiveSheet()->setCellValue("C113", $c113);
$objPHPExcel->getActiveSheet()->setCellValue("C114", $c114);
$objPHPExcel->getActiveSheet()->setCellValue("C115", $c115);
$objPHPExcel->getActiveSheet()->setCellValue("C116", $c116);
$objPHPExcel->getActiveSheet()->setCellValue("C117", $c117);
$objPHPExcel->getActiveSheet()->setCellValue("C118", $c118);
$objPHPExcel->getActiveSheet()->setCellValue("C119", $c119);
$objPHPExcel->getActiveSheet()->setCellValue("C120", $c120);
$objPHPExcel->getActiveSheet()->setCellValue("C121", $c121);
$objPHPExcel->getActiveSheet()->setCellValue("C122", $c122);
$objPHPExcel->getActiveSheet()->setCellValue("C123", $c123);
$objPHPExcel->getActiveSheet()->setCellValue("C124", $c124);
$objPHPExcel->getActiveSheet()->setCellValue("C125", $c125);
$objPHPExcel->getActiveSheet()->setCellValue("C126", $c126);
$objPHPExcel->getActiveSheet()->setCellValue("C127", $c127);
$objPHPExcel->getActiveSheet()->setCellValue("C128", $c128);
$objPHPExcel->getActiveSheet()->setCellValue("C129", $c129);
$objPHPExcel->getActiveSheet()->setCellValue("C130", $c130);
$objPHPExcel->getActiveSheet()->setCellValue("C131", $c131);
$objPHPExcel->getActiveSheet()->setCellValue("C132", $c132);
$objPHPExcel->getActiveSheet()->setCellValue("C133", $c133);
$objPHPExcel->getActiveSheet()->setCellValue("C134", $c134);
$objPHPExcel->getActiveSheet()->setCellValue("C135", $c135);
$objPHPExcel->getActiveSheet()->setCellValue("C136", $c136);
$objPHPExcel->getActiveSheet()->setCellValue("C137", $c137);
$objPHPExcel->getActiveSheet()->setCellValue("C138", $c138);
$objPHPExcel->getActiveSheet()->setCellValue("C139", $c139);
$objPHPExcel->getActiveSheet()->setCellValue("C140", $c140);
$objPHPExcel->getActiveSheet()->setCellValue("C141", $c141);
$objPHPExcel->getActiveSheet()->setCellValue("C142", $c142);
$objPHPExcel->getActiveSheet()->setCellValue("C143", $c143);
$objPHPExcel->getActiveSheet()->setCellValue("C144", $c144);
$objPHPExcel->getActiveSheet()->setCellValue("C145", $c145);
$objPHPExcel->getActiveSheet()->setCellValue("C146", $c146);
$objPHPExcel->getActiveSheet()->setCellValue("C147", $c147);
$objPHPExcel->getActiveSheet()->setCellValue("C148", $c148);
$objPHPExcel->getActiveSheet()->setCellValue("C149", $c149);
$objPHPExcel->getActiveSheet()->setCellValue("C150", $c150);
$objPHPExcel->getActiveSheet()->setCellValue("C151", $c151);
$objPHPExcel->getActiveSheet()->setCellValue("C152", $c152);
$objPHPExcel->getActiveSheet()->setCellValue("C153", $c153);
$objPHPExcel->getActiveSheet()->setCellValue("C154", $c154);
$objPHPExcel->getActiveSheet()->setCellValue("C155", $c155);
$objPHPExcel->getActiveSheet()->setCellValue("C156", $c156);
$objPHPExcel->getActiveSheet()->setCellValue("C157", $c157);
$objPHPExcel->getActiveSheet()->setCellValue("C158", $c158);
$objPHPExcel->getActiveSheet()->setCellValue("C159", $c159);
$objPHPExcel->getActiveSheet()->setCellValue("C160", $c160);
$objPHPExcel->getActiveSheet()->setCellValue("C161", $c161);
$objPHPExcel->getActiveSheet()->setCellValue("C162", $c162);
$objPHPExcel->getActiveSheet()->setCellValue("C163", $c163);
$objPHPExcel->getActiveSheet()->setCellValue("C164", $c164);
$objPHPExcel->getActiveSheet()->setCellValue("C165", $c165);
$objPHPExcel->getActiveSheet()->setCellValue("C166", $c166);
$objPHPExcel->getActiveSheet()->setCellValue("C167", $c167);
$objPHPExcel->getActiveSheet()->setCellValue("C168", $c168);
$objPHPExcel->getActiveSheet()->setCellValue("C169", $c169);
$objPHPExcel->getActiveSheet()->setCellValue("C170", $c170);
$objPHPExcel->getActiveSheet()->setCellValue("C171", $c171);
$objPHPExcel->getActiveSheet()->setCellValue("C172", $c172);
$objPHPExcel->getActiveSheet()->setCellValue("C173", $c173);
$objPHPExcel->getActiveSheet()->setCellValue("C174", $c174);
//$objPHPExcel->getActiveSheet()->setCellValue("C111", $c111);
//$objPHPExcel->getActiveSheet()->setCellValue("C112", $c112);
//175-216
$objPHPExcel->getActiveSheet()->setCellValue("C175", $c175);
$objPHPExcel->getActiveSheet()->setCellValue("C176", $c176);
$objPHPExcel->getActiveSheet()->setCellValue("C177", $c177);
$objPHPExcel->getActiveSheet()->setCellValue("C178", $c178);
$objPHPExcel->getActiveSheet()->setCellValue("C179", $c179);
$objPHPExcel->getActiveSheet()->setCellValue("C180", $c180);
$objPHPExcel->getActiveSheet()->setCellValue("C181", $c181);
$objPHPExcel->getActiveSheet()->setCellValue("C182", $c182);
$objPHPExcel->getActiveSheet()->setCellValue("C183", $c183);
$objPHPExcel->getActiveSheet()->setCellValue("C184", $c184);
$objPHPExcel->getActiveSheet()->setCellValue("C185", $c185);
$objPHPExcel->getActiveSheet()->setCellValue("C186", $c186);
$objPHPExcel->getActiveSheet()->setCellValue("C187", $c187);
// $objPHPExcel->getActiveSheet()->setCellValue("C188", $c188);
// $objPHPExcel->getActiveSheet()->setCellValue("C189", $c189);
// $objPHPExcel->getActiveSheet()->setCellValue("C190", $c190);
// $objPHPExcel->getActiveSheet()->setCellValue("C191", $c191);
// $objPHPExcel->getActiveSheet()->setCellValue("C192", $c192);
// $objPHPExcel->getActiveSheet()->setCellValue("C193", $c193);
// $objPHPExcel->getActiveSheet()->setCellValue("C194", $c194);
// $objPHPExcel->getActiveSheet()->setCellValue("C195", $c195);
$objPHPExcel->getActiveSheet()->setCellValue("C196", $c196);
$objPHPExcel->getActiveSheet()->setCellValue("C197", $c197);
$objPHPExcel->getActiveSheet()->setCellValue("C198", $c198);
$objPHPExcel->getActiveSheet()->setCellValue("C199", $c199);

$objPHPExcel->getActiveSheet()->setCellValue("C200", $c200);
$objPHPExcel->getActiveSheet()->setCellValue("C201", $c201);
$objPHPExcel->getActiveSheet()->setCellValue("C202", $c202);
$objPHPExcel->getActiveSheet()->setCellValue("C203", $c203);
$objPHPExcel->getActiveSheet()->setCellValue("C204", $c204);
$objPHPExcel->getActiveSheet()->setCellValue("C205", $c205);
$objPHPExcel->getActiveSheet()->setCellValue("C206", $c206);
$objPHPExcel->getActiveSheet()->setCellValue("C207", $c207);
$objPHPExcel->getActiveSheet()->setCellValue("C208", $c208);
$objPHPExcel->getActiveSheet()->setCellValue("C209", $c209);
$objPHPExcel->getActiveSheet()->setCellValue("C210", $c210);
$objPHPExcel->getActiveSheet()->setCellValue("C211", $c211);
$objPHPExcel->getActiveSheet()->setCellValue("C212", $c212);
$objPHPExcel->getActiveSheet()->setCellValue("C213", $c213);
$objPHPExcel->getActiveSheet()->setCellValue("C214", $c214);
$objPHPExcel->getActiveSheet()->setCellValue("C215", $c215);
$objPHPExcel->getActiveSheet()->setCellValue("C216", $c216);
$objPHPExcel->getActiveSheet()->setCellValue("C217", $c217);
$objPHPExcel->getActiveSheet()->setCellValue("C218", $c218);
$objPHPExcel->getActiveSheet()->setCellValue("C219", $c219);

$objPHPExcel->getActiveSheet()->setCellValue("C220", $c220);
$objPHPExcel->getActiveSheet()->setCellValue("C221", $c221);
$objPHPExcel->getActiveSheet()->setCellValue("C222", $c222);
$objPHPExcel->getActiveSheet()->setCellValue("C223", $c223);
$objPHPExcel->getActiveSheet()->setCellValue("C224", $c224);
$objPHPExcel->getActiveSheet()->setCellValue("C225", $c225);
$objPHPExcel->getActiveSheet()->setCellValue("C226", $c226);
$objPHPExcel->getActiveSheet()->setCellValue("C227", $c227);
$objPHPExcel->getActiveSheet()->setCellValue("C228", $c228);
$objPHPExcel->getActiveSheet()->setCellValue("C229", $c229);
$objPHPExcel->getActiveSheet()->setCellValue("C230", $c230);
$objPHPExcel->getActiveSheet()->setCellValue("C231", $c231);
$objPHPExcel->getActiveSheet()->setCellValue("C232", $c232);
$objPHPExcel->getActiveSheet()->setCellValue("C233", $c233);
$objPHPExcel->getActiveSheet()->setCellValue("C234", $c234);
$objPHPExcel->getActiveSheet()->setCellValue("C235", $c235);
$objPHPExcel->getActiveSheet()->setCellValue("C236", $c236);
$objPHPExcel->getActiveSheet()->setCellValue("C237", $c237);
$objPHPExcel->getActiveSheet()->setCellValue("C238", $c238);
$objPHPExcel->getActiveSheet()->setCellValue("C239", $c239);

$objPHPExcel->getActiveSheet()->setCellValue("C240", $c240);
$objPHPExcel->getActiveSheet()->setCellValue("C241", $c241);




$objPHPExcel->getActiveSheet()->setCellValue('B99', 'LAPORAN LABA RUGI BANK UMUM - LRBU');
$objPHPExcel->getActiveSheet()->setCellValue('B100', '(Dalam Jutaan Rupiah)');
$objPHPExcel->getActiveSheet()->setCellValue('B102', 'PT BANK MNC INTERNASIONAL, Tbk');
$objPHPExcel->getActiveSheet()->setCellValue('B103', '485000');
$objPHPExcel->getActiveSheet()->setCellValue('B104', 'POS-POS');


$objPHPExcel->getActiveSheet()->setCellValue('B105', 'PENDAPATAN DAN BEBAN OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B106', 'A. Pendapatan dan Beban Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B107', 'a. Pendapatan Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B108', '   i.  Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B109', '   ii. Valuta Asing');
$objPHPExcel->getActiveSheet()->setCellValue('B110', 'b. Beban Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B111', '   i.  Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B112', '   ii. Valuta Asing');
$objPHPExcel->getActiveSheet()->setCellValue('B113', 'Pendapatan (Beban) Bunga Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B114', '');
$objPHPExcel->getActiveSheet()->setCellValue('B115', 'B. Pendapatan dan Beban Operasional selain Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B116', '1. Pendapatan Operasional Selain Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B117', 'a. Peningkatan nilai wajar aset keuangan (mark to market)');
$objPHPExcel->getActiveSheet()->setCellValue('B118', '   i.   Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B119', '   ii.  Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B120', '   iii. Spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B121', '   iv.  Aset keuangan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B122', 'b. Penurunan nilai wajar kewajiban keuangan (mark to market)');
$objPHPExcel->getActiveSheet()->setCellValue('B123', 'c. Keuntungan penjualan aset keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('B124', '   i.   Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B125', '   ii.  Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B126', '   iii. Aset keuangan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B127', 'd. Keuntungan transaksi spot dan derivatif (realised)');
$objPHPExcel->getActiveSheet()->setCellValue('B128', 'e. Keuntungan dari penyertaan dengan equity method,');
$objPHPExcel->getActiveSheet()->setCellValue('B129', 'f. Dividen');
$objPHPExcel->getActiveSheet()->setCellValue('B130', 'g. komisi/provisi/fee dan administrasi');
$objPHPExcel->getActiveSheet()->setCellValue('B131', 'h. Koreksi atas cadangan kerugian penurunan nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B132', 'i. Pendapatan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B133', '');
$objPHPExcel->getActiveSheet()->setCellValue('B134', '2. Beban Operasional Selain Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B135', 'a. Penurunan nilai wajar aset keuangan (mark to market)');
$objPHPExcel->getActiveSheet()->setCellValue('B136', '   i.   Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B137', '   ii.  Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B138', '   iii. Spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B139', '   iv.  Aset keuangan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B140', 'b. Peningkatan nilai wajar kewajiban keuangan (mark to market)');
$objPHPExcel->getActiveSheet()->setCellValue('B141', 'c. Kerugian penjualan aset keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('B142', '   i.   Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B143', '   ii.  Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B144', '   iii. Aset keuangan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B145', 'd. Kerugian transaksi spot dan derivatif (realised)');
$objPHPExcel->getActiveSheet()->setCellValue('B146', 'e. Kerugian penurunan nilai aset keuangan (impairment)');
$objPHPExcel->getActiveSheet()->setCellValue('B147', '   i.   Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B148', '   ii.  Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B149', '   iii. Pembiayaan syariah');
$objPHPExcel->getActiveSheet()->setCellValue('B150', '   iv.  Aset keuangan lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B151', 'f.  Kerugian risiko operasional');
$objPHPExcel->getActiveSheet()->setCellValue('B152', 'g. Kerugian dari penyertaan dengan equity method');
$objPHPExcel->getActiveSheet()->setCellValue('B153', 'h. Komisi/provisi/fee dan administrasi');
$objPHPExcel->getActiveSheet()->setCellValue('B154', 'i. Kerugian penurunan nilai aset lainnya (non keuangan)');
$objPHPExcel->getActiveSheet()->setCellValue('B155', 'j. Beban tenaga kerja');
$objPHPExcel->getActiveSheet()->setCellValue('B156', 'k. Beban promosi');
$objPHPExcel->getActiveSheet()->setCellValue('B157', 'l. Beban lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B158', 'Pendapatan (Beban) Operasional Selain Bunga Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B159', 'LABA (RUGI) OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B160', '');
$objPHPExcel->getActiveSheet()->setCellValue('B161', 'PENDAPATAN DAN BEBAN NON OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B162', 'Keuntungan (kerugian) penjualan aset tetap dan inventaris');
$objPHPExcel->getActiveSheet()->setCellValue('B163', 'Keuntungan (kerugian) penjabaran transaksi valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B164', 'Pendapatan (beban) non operasional lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B165', 'LABA (RUGI) NON OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B166', 'LABA (RUGI) TAHUN BERJALAN ');
$objPHPExcel->getActiveSheet()->setCellValue('B167', 'Pajak penghasilan');
$objPHPExcel->getActiveSheet()->setCellValue('B168', ' a. Taksiran pajak tahun berjalan');
$objPHPExcel->getActiveSheet()->setCellValue('B169', ' b. Pendapatan (beban) pajak tangguhan');
$objPHPExcel->getActiveSheet()->setCellValue('B170', 'LABA (RUGI) BERSIH');

$objPHPExcel->getActiveSheet()->setCellValue('B171', '');
$objPHPExcel->getActiveSheet()->setCellValue('B172', 'PENGHASILAN KOMPREHENSIF LAIN');
$objPHPExcel->getActiveSheet()->setCellValue('B173', '1. Pos-pos yang tidak akan direklasifikasi ke laba rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B174', 'a. Keuntungan Revaluasi aset tetap');
$objPHPExcel->getActiveSheet()->setCellValue('B175', 'b. Keuntungan (kerugian) aktuarial program imbalan pasti');
$objPHPExcel->getActiveSheet()->setCellValue('B176', 'c. Bagian pendapatan komprehensif lain dari entitas asosiasi');
$objPHPExcel->getActiveSheet()->setCellValue('B177', 'd. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B178', 'e. Pajak penghasilan terkait pos-pos yang tidak akan direklasifikasi ke laba rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B179', '2. Pos-pos yang akan direklasifikasi ke laba rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B180', 'a. Penyesuaian akibat penjabaran laporan keuangan dalam mata uang asing');
$objPHPExcel->getActiveSheet()->setCellValue('B181', 'b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('B182', 'c. Bagian efektif dari lindung nilai arus kas');
$objPHPExcel->getActiveSheet()->setCellValue('B183', 'd. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B184', 'e. Pajak penghasilan terkait pos-pos yang  akan direklasifikasi ke laba rugi');
$objPHPExcel->getActiveSheet()->setCellValue('B185', 'PENGHASILAN KOMPREHENSIF LAIN TAHUN BERJALAN - NET PAJAK PENGHASILAN TERKAIT');
$objPHPExcel->getActiveSheet()->setCellValue('B186', 'TOTAL LABA (RUGI) KOMPREHENSIF TAHUN BERJALAN');
$objPHPExcel->getActiveSheet()->setCellValue('B187', 'TRANSFER LABA (RUGI) KE KANTOR PUSAT');


//$objPHPExcel->getActiveSheet()->setCellValue('B171', 'Transfer laba (rugi) ke kantor pusat');



$objPHPExcel->getActiveSheet()->setCellValue('B191', 'KOMITMEN & KONTINJENSI BANK UMUM - KKBU');
$objPHPExcel->getActiveSheet()->setCellValue('B192', '(Dalam Jutaan Rupiah)');
$objPHPExcel->getActiveSheet()->setCellValue('B193', 'PT BANK MNC INTERNASIONAL Tbk');
$objPHPExcel->getActiveSheet()->setCellValue('B194', '485000');
$objPHPExcel->getActiveSheet()->setCellValue('B195', 'POS-POS');

$objPHPExcel->getActiveSheet()->setCellValue('B196', ' TAGIHAN KOMITMEN ');

$objPHPExcel->getActiveSheet()->setCellValue('B197', ' 1. Fasilitas pinjaman yang belum ditarik ');
$objPHPExcel->getActiveSheet()->setCellValue('B198', '    a. Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B199', '    b. Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B200', ' 2. Posisi pembelian spot dan derivatif yang masih berjalan ');
$objPHPExcel->getActiveSheet()->setCellValue('B201', ' 3. Lainnya ');
$objPHPExcel->getActiveSheet()->setCellValue('B202', '');
$objPHPExcel->getActiveSheet()->setCellValue('B203', ' KEWAJIBAN KOMITMEN ');
$objPHPExcel->getActiveSheet()->setCellValue('B204', ' 1. Fasilitas kredit kepada nasabah yang belum ditarik ');
$objPHPExcel->getActiveSheet()->setCellValue('B205', '      a. BUMN ');
$objPHPExcel->getActiveSheet()->setCellValue('B206', '         i.  Committed ');
$objPHPExcel->getActiveSheet()->setCellValue('B207', '             - Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B208', '             - Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B209', '         ii. Uncommitted ');
$objPHPExcel->getActiveSheet()->setCellValue('B210', '             - Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B211', '             - Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B212', '      b. Lainnya ');
$objPHPExcel->getActiveSheet()->setCellValue('B213', '         i.  Committed ');
$objPHPExcel->getActiveSheet()->setCellValue('B214', '         ii. Uncommitted ');
$objPHPExcel->getActiveSheet()->setCellValue('B215', ' 2. Fasilitas kredit kepada bank lain yang belum ditarik ');
$objPHPExcel->getActiveSheet()->setCellValue('B216', '      a. Committed ');
$objPHPExcel->getActiveSheet()->setCellValue('B217', '         i. Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B218', '         ii. Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B219', '      b. Uncommitted ');
$objPHPExcel->getActiveSheet()->setCellValue('B220', '            i. Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B221', '            ii. Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B222', '3. Irrevocable L/C yang masih berjalan ');
$objPHPExcel->getActiveSheet()->setCellValue('B223', '      a. L/C luar negeri');
$objPHPExcel->getActiveSheet()->setCellValue('B224', '      b. L/C dalam negeri');
$objPHPExcel->getActiveSheet()->setCellValue('B225', '4. Posisi penjualan spot dan derivatif yang masih berjalan');
$objPHPExcel->getActiveSheet()->setCellValue('B226', '5. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B227', '');
$objPHPExcel->getActiveSheet()->setCellValue('B228', ' TAGIHAN KONTINJENSI ');
$objPHPExcel->getActiveSheet()->setCellValue('B229', ' 1. Garansi yang diterima ');
$objPHPExcel->getActiveSheet()->setCellValue('B230', '      a. Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B231', '      b. Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B232', ' 2. Pendapatan bunga dalam penyelesaian ');
$objPHPExcel->getActiveSheet()->setCellValue('B233', '      a. Bunga kredit yang diberikan ');
$objPHPExcel->getActiveSheet()->setCellValue('B234', '      b. Bunga lainnya ');
$objPHPExcel->getActiveSheet()->setCellValue('B235', ' 3. Lainnya ');
$objPHPExcel->getActiveSheet()->setCellValue('B236', '');
$objPHPExcel->getActiveSheet()->setCellValue('B237', ' KEWAJIBAN KONTIJENSI ');
$objPHPExcel->getActiveSheet()->setCellValue('B238', ' 1. Garansi yang diberikan ');
$objPHPExcel->getActiveSheet()->setCellValue('B239', '      a. Rupiah ');
$objPHPExcel->getActiveSheet()->setCellValue('B240', '      b. Valuta asing ');
$objPHPExcel->getActiveSheet()->setCellValue('B241', ' 2. Lainnya ');


for ($i=10;$i<=241;$i++) {

	if ($i=='45' || $i=='49' || $i=='50' || $i=='68' || $i=='78' || $i=='80' || $i=='82' || $i=='83' || $i=='84' || $i=='86' || $i=='88' || $i=='89'|| $i=='98' || $i=='99' || $i=='100' || $i=='101' || $i=='102' || $i=='103' || $i=='104' || $i=='105' || $i=='106' || $i=='114'|| $i=='115' || $i=='133' || $i=='160' || $i=='161' || $i=='171' || $i=='172' || $i=='174'|| $i=='176' || $i=='177' || $i=='178'|| $i=='188' || $i=='189' || $i=='190' || $i=='191' || $i=='192' || $i=='193' || $i=='194' || $i=='195' || $i=='202' || $i=='227' || $i=='236' ){

	} else {

    		$colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    		if ($colB == NULL || $colB == '') {
        		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    		}
    	}
}


  
$objPHPExcel->getActiveSheet()->getStyle('C10:C48')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C51:C97')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C107:C187')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C196:C241')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');




# Giro Jumlah Nominal --> Kolom D
// $index=10;
// foreach ($giro_nominal as $nilai ) {
//   $objPHPExcel->getActiveSheet()->setCellValue("D$index", floatval($nilai) );

// $index++;

// }





$objPHPExcel->getActiveSheet()->setTitle('Neraca PL');




##################  sheet ke 2
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
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A1:P1');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:P2'); 
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:P3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('M4:P4');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A5:L7');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('M5:N6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('O5:P6');


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A64:F65');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K64:L65');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('M64:N64');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('G64:H64');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('I64:J64');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('O64:P64');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K66:L66');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('K72:L72');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A77:G77');

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




// SHEET KE 3 ======================================================================================
// Create new PHPExcel object
//$objPHPExcel = new PHPExcel();
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);

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

 $objPHPExcel->getActiveSheet()->getStyle('B2:H9')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('B101:H108')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('B128:H130')->applyFromArray($styleArrayFontBold);

// $objPHPExcel->getActiveSheet()->getStyle('A8:N9')->applyFromArray($styleArrayFontBold);



 $objPHPExcel->getActiveSheet()->getStyle('B2:H9')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('B2:H9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


 $objPHPExcel->getActiveSheet()->getStyle('B101:F105')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('B101:F105')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('B106:F108')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('B106:F108')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('C122:H122')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('C122:H122')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B128:F130')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('C128:F130')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

 $objPHPExcel->getActiveSheet()->getStyle('B8:H99')->applyFromArray($styleArrayBorder1);
 $objPHPExcel->getActiveSheet()->getStyle('B2:H5')->applyFromArray($styleArrayBorder2);
 $objPHPExcel->getActiveSheet()->getStyle('B6:H7')->applyFromArray($styleArrayBorder2);
 $objPHPExcel->getActiveSheet()->getStyle('B101:F105')->applyFromArray($styleArrayBorder2);

$objPHPExcel->getActiveSheet()->getStyle('B106:F119')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B122:H125')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B128:F131')->applyFromArray($styleArrayBorder1);
//FILL COLOR

 $objPHPExcel->getActiveSheet()->getStyle('A1:I140')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
 //$objPHPExcel->getActiveSheet()->getStyle('I1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
 //$objPHPExcel->getActiveSheet()->getStyle('A53:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

// $objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
// $objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');

//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(75);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(5);



// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(2);

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B3:H3');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B4:H4');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B6:H6');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B7:H7');


 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('B8:B9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('C8:C9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('D8:D9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('E8:E9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('F8:F9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('G8:G9');
 $objPHPExcel->setActiveSheetIndex(2)->mergeCells('H8:H9');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B102:F102');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B103:F103');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B104:F104');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B106:B108');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C107:D107');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('E107:F107');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('B128:B130');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C129:D129');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('E129:F129');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C106:F106');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C128:F128');




$objPHPExcel->getActiveSheet()->getStyle('D10:D51')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('B3', 'KUALITAS ASET PRODUKTIF BANK UMUM - KAPBU');
$objPHPExcel->getActiveSheet()->setCellValue('B4', "(Dalam Jutaan Rupiah)");
//$objPHPExcel->getActiveSheet()->setCellValue('A3', 'RUPIAH');

$objPHPExcel->getActiveSheet()->setCellValue('B6', '31-Jan-16');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'POS-POS');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'L');
$objPHPExcel->getActiveSheet()->setCellValue('D8', 'DPK');
$objPHPExcel->getActiveSheet()->setCellValue('E8', 'KL');
$objPHPExcel->getActiveSheet()->setCellValue('F8', "D");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "M");
$objPHPExcel->getActiveSheet()->setCellValue('H8', "Jumlah");

#############  HEADER ################

$objPHPExcel->getActiveSheet()->setCellValue('B10', 'I. PIHAK TERKAIT');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Penempatan pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B14', '');
$objPHPExcel->getActiveSheet()->setCellValue('B15', "Tagihan spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "b. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Surat berharga");

$objPHPExcel->getActiveSheet()->setCellValue('B20', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B22', '');
$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Surat berharga yang dijual dengan janji dibeli kembali (Repo)');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B25', "b. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "Tagihan atas surat berharga dibeli dengan janji dijual kembali (Reverse Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "b. Valuta asing");

$objPHPExcel->getActiveSheet()->setCellValue('B30', '');
$objPHPExcel->getActiveSheet()->setCellValue('B31', 'Tagihan akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('B32', '');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'a. Debitur Usaha Mikro, Kecil dan Menengah (UMKM)');
$objPHPExcel->getActiveSheet()->setCellValue('B35', "     i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "     ii. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "b. Bukan Debitur UMKM");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "     i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "     ii. Valuta asing");


$objPHPExcel->getActiveSheet()->setCellValue('B40', 'c. Kredit yang direstrukturisasi');
$objPHPExcel->getActiveSheet()->setCellValue('B41', '     i.  Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B42', '     ii. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B43', 'd. Kredit properti');
$objPHPExcel->getActiveSheet()->setCellValue('B44', '');
$objPHPExcel->getActiveSheet()->setCellValue('B45', "Penyertaan");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "Penyertaan modal sementara");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "");
$objPHPExcel->getActiveSheet()->setCellValue('B49', "Komitmen dan kontinjensi");


$objPHPExcel->getActiveSheet()->setCellValue('B50', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B51', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B52', '     ii. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B53', '');
$objPHPExcel->getActiveSheet()->setCellValue('B54', '');
$objPHPExcel->getActiveSheet()->setCellValue('B55', "II. PIHAK TIDAK TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "Penempatan pada bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "b. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "");

$objPHPExcel->getActiveSheet()->setCellValue('B60', 'Tagihan spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B63', '');
$objPHPExcel->getActiveSheet()->setCellValue('B64', 'Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B65', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B66', "b. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B67', "");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "Surat berharga yang dijual dengan janji dibeli kembali (Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B69', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B71', '');
$objPHPExcel->getActiveSheet()->setCellValue('B72', 'Tagihan atas surat berharga dibeli dengan janji dijual kembali (Reverse Repo)');
$objPHPExcel->getActiveSheet()->setCellValue('B73', 'a. Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B74', 'b. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B75', "");
$objPHPExcel->getActiveSheet()->setCellValue('B76', "Tagihan akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('B77', "");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "a. Debitur Usaha Mikro, Kecil dan Menengah (UMKM)");
$objPHPExcel->getActiveSheet()->setCellValue('B80', '     i.  Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B81', '     ii. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B82', 'b. Bukan Debitur UMKM');
$objPHPExcel->getActiveSheet()->setCellValue('B83', '     i.  Rupiah');
$objPHPExcel->getActiveSheet()->setCellValue('B84', '     ii. Valuta asing');
$objPHPExcel->getActiveSheet()->setCellValue('B85', "c. Kredit yang direstrukturisasi");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "     i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "     ii. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "d. Kredit properti");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "");
$objPHPExcel->getActiveSheet()->setCellValue('B90', 'Penyertaan');
$objPHPExcel->getActiveSheet()->setCellValue('B91', '');
$objPHPExcel->getActiveSheet()->setCellValue('B92', 'Penyertaan modal sementara');
$objPHPExcel->getActiveSheet()->setCellValue('B93', '');
$objPHPExcel->getActiveSheet()->setCellValue('B94', 'Komitmen dan kontinjensi');
$objPHPExcel->getActiveSheet()->setCellValue('B95', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B96', "b. Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('B97', "");
$objPHPExcel->getActiveSheet()->setCellValue('B98', "Aset yang diambil alih");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "");



###### CADANGAN PENYISIHAN KERUGIAN ############

$objPHPExcel->getActiveSheet()->setCellValue('B102', "Sunday, January 31, 2016        ");
$objPHPExcel->getActiveSheet()->setCellValue('B103', "CADANGAN PENYISIHAN KERUGIAN        ");
$objPHPExcel->getActiveSheet()->setCellValue('B104', "(Dalam Jutaan Rupiah)       ");

$objPHPExcel->getActiveSheet()->setCellValue('B106', "POS-POS");
$objPHPExcel->getActiveSheet()->setCellValue('C107', "CKPN");
$objPHPExcel->getActiveSheet()->setCellValue('E107', "PPA wajib dibentuk");

$objPHPExcel->getActiveSheet()->setCellValue('C108', "Individual");
$objPHPExcel->getActiveSheet()->setCellValue('D108', "Kolektif");
$objPHPExcel->getActiveSheet()->setCellValue('E108', "Umum");
$objPHPExcel->getActiveSheet()->setCellValue('F108', "Khusus");

$objPHPExcel->getActiveSheet()->setCellValue('B110', "Penempatan pada bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B111', "Tagihan spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B112', "Surat berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B113', "Surat berharga yang dijual dengan janji dibeli kembali (Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B114', "Tagihan atas surat berharga yang dibeli dengan janji dijual kembali (Reverse Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B115', "Tagihan akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('B116', "Kredit ");
$objPHPExcel->getActiveSheet()->setCellValue('B117', "Penyertaan ");
$objPHPExcel->getActiveSheet()->setCellValue('B118', "Penyertaan modal sementara");
$objPHPExcel->getActiveSheet()->setCellValue('B119', "Transaksi rekening administratif");


$objPHPExcel->getActiveSheet()->setCellValue('B122', "I. PIHAK TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('C122', "L");
$objPHPExcel->getActiveSheet()->setCellValue('D122', "DPK");
$objPHPExcel->getActiveSheet()->setCellValue('E122', "KL");
$objPHPExcel->getActiveSheet()->setCellValue('F122', "D");
$objPHPExcel->getActiveSheet()->setCellValue('G122', "M");
$objPHPExcel->getActiveSheet()->setCellValue('H122', "Jumlah");


$objPHPExcel->getActiveSheet()->setCellValue('B123', "Tagihan Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B124', "II. PIHAK TIDAK TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('B125', "Tagihan Lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B128', "POS-POS");
$objPHPExcel->getActiveSheet()->setCellValue('C129', "CKPN");
$objPHPExcel->getActiveSheet()->setCellValue('E129', "PPA wajib dibentuk");

$objPHPExcel->getActiveSheet()->setCellValue('C130', "Individual");
$objPHPExcel->getActiveSheet()->setCellValue('D130', "Kolektif");
$objPHPExcel->getActiveSheet()->setCellValue('E130', "Umum");
$objPHPExcel->getActiveSheet()->setCellValue('F130', "Khusus");
$objPHPExcel->getActiveSheet()->setCellValue('B131', "Tagihan Lainnya");



$objPHPExcel->getActiveSheet()->setTitle('KAP');








// SHEET KE 4 ======================================================================================
// Create new PHPExcel object
//$objPHPExcel = new PHPExcel();
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3);

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

 $objPHPExcel->getActiveSheet()->getStyle('A1:C11')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A21:C21')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A32:C33')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('B42:C42')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('B46:C46')->applyFromArray($styleArrayFontBold);

 $objPHPExcel->getActiveSheet()->getStyle('B66:B69')->applyFromArray($styleArrayFontBold);
 $objPHPExcel->getActiveSheet()->getStyle('A86:C86')->applyFromArray($styleArrayFontBold);



 $objPHPExcel->getActiveSheet()->getStyle('A1:C9')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
 $objPHPExcel->getActiveSheet()->getStyle('A1:C9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);





//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B8:C91')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B2:C7')->applyFromArray($styleArrayBorder2);



//FILL COLOR

$objPHPExcel->getActiveSheet()->getStyle('A1:D100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');


// $objPHPExcel->getActiveSheet()->getStyle('A8:N9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
// $objPHPExcel->getActiveSheet()->getStyle('A52:N52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
//DIMENSION D

 $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
 $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(75);
 $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
 $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(5);



// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(3);

 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('B3:C3');
 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('B4:C4');
 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('B5:C5');
 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('B6:C6');
 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('B8:B9');
 $objPHPExcel->setActiveSheetIndex(3)->mergeCells('C8:C9');



$objPHPExcel->getActiveSheet()->getStyle('C11:C91')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('B3', ' PT. BANK MNC INTERNASIONAL Tbk ');
$objPHPExcel->getActiveSheet()->setCellValue('B4', "485000");
$objPHPExcel->getActiveSheet()->setCellValue('B5', 'INFORMASI TAMBAHAN BANK UMUM - ITBU');
$objPHPExcel->getActiveSheet()->setCellValue('B6', '(Dalam Jutaan Rupiah) ');


$objPHPExcel->getActiveSheet()->setCellValue('B8', 'POS-POS');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Sunday, January 31, 2016');

$objPHPExcel->getActiveSheet()->setCellValue('B10', "I. PENEMPATAN DAN KEWAJIBAN PADA BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "A. PENEMPATAN PADA BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "1. Bank Persero");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "2. BUSN Devisa");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "3. BUSN Non Devisa");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "4. BPD");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "5. Bank Campuran");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "6. Bank Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "7. BPR/BPRS");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "8. Bank di Luar Negeri");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "9. Lain-Lain");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "B. KEWAJIBAN PADA BANK LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "1. Bank Persero");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "2. BUSN Devisa");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "3. BUSN Non Devisa");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "4. BPD");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "5. Bank Campuran");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "6. Bank Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "7. BPR/BPRS");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "8. Bank di Luar Negeri");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "9. Lain-Lain");

$objPHPExcel->getActiveSheet()->setCellValue('B31', "");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "II. PORTFOLIO KREDIT YANG DIBERIKAN");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "A. BERDASARKAN GOLONGAN DEBITUR");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "1. Badan/Lembaga pemerintah");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "2. Pemerintah daerah");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "3. BUMN");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "4. BUMD");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "5. BUMS (Swasta)");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "6. BPR/BPRS");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "7. Perorangan");
$objPHPExcel->getActiveSheet()->setCellValue('B41', "8. Lain-Lain");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "B. BERDASARKAN JENIS PENGGUNAAN");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "1. Kredit Konsumsi ");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "2. Kredit Modal Kerja ");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "3. Kredit Investasi");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "C. BERDASARKAN SEKTOR INDUSTRI");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "1. Pertanian, perburuan dan kehutanan");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "2. Perikanan");
$objPHPExcel->getActiveSheet()->setCellValue('B49', "3. Pertambangan dan penggalian");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "4. Industri pengolahan");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "5. Listrik, gas dan air");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "6. Konstruksi");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "7. Perdagangan besar dan eceran");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "8. Penyediaan akomodasi dan penyediaan makan minum");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "9. Transportasi, pergudangan dan komunikasi");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "10. Perantara keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "11. Real estate, usaha persewaan, dan jasa perusahaan");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "12. Administrasi pemerintahan, pertahanan dan jaminan sosial wajib");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "13. Jasa pendidikan");
$objPHPExcel->getActiveSheet()->setCellValue('B60', "14. Jasa kesehatan dan kegiatan sosial");
$objPHPExcel->getActiveSheet()->setCellValue('B61', "15. Jasa kemasyarakatan, sosial budaya, hiburan dan perorangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "16. Jasa perorangan yang melayani rumah tangga");
$objPHPExcel->getActiveSheet()->setCellValue('B63', "17. Rumah tangga");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "18. Bukan lapangan usaha lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "");
$objPHPExcel->getActiveSheet()->setCellValue('B66', "III. COUNTER RATE DANA PIHAK KETIGA (%)");
$objPHPExcel->getActiveSheet()->setCellValue('B67', "A. GIRO (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "B. TABUNGAN (Rupiah)");
$objPHPExcel->getActiveSheet()->setCellValue('B69', "C. DEPOSITO BERJANGKA");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "1. 1 Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B71', "     a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B72', "     b. US$");
$objPHPExcel->getActiveSheet()->setCellValue('B73', "2. 3 Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B74', "     a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B75', "     b. US$");
$objPHPExcel->getActiveSheet()->setCellValue('B76', "3. 6 Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B77', "     a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "     b. US$");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "4. 12 Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B80', "     a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B81', "     b. US$");
$objPHPExcel->getActiveSheet()->setCellValue('B82', "5. 24 Bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B83', "     a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "     b. US$");
$objPHPExcel->getActiveSheet()->setCellValue('B85', "");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "IV. LIKUIDITAS");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "1. Rasio (%) Giro Wajib Minimum (GWM) akhir bulan");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "     a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "     b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B90', "2. Rasio (%) Aset Lancar 1 bulan ke depan/ Kewajiban Lancar 1 bulan ke depan");
$objPHPExcel->getActiveSheet()->setCellValue('B91', "3. Rasio (%) Kas/ Kewajiban Lancar 1 bulan ke depan");








#############  HEADER ################









$objPHPExcel->getActiveSheet()->setTitle('Informasi Tambahan');













/*
var_dump($giro_nominal);
echo "<br><br>";
var_dump($tabungan_nominal);
die();
*/






$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/KPMM_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/KPMM_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>

<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> Keuangan LPS 
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
                            
                            <div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/KPMM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> </div>
                            <div class="tabbable-line">
                               <ul class="nav nav-tabs ">
                                <li class="active" >
                                        <a href="#tab_15_1" data-toggle="tab">
                                        Neraca PL </a>
                                    </li>
                                    <li >
                                        <a href="#tab_15_2" data-toggle="tab">
                                        KPMM </a>
                                    </li>
                                   
                                    <li >
                                        <a href="#tab_15_3" data-toggle="tab">
                                        KAP </a>
                                    </li>
                                    <li >
                                        <a href="#tab_15_4" data-toggle="tab">
                                        Informasi Tambahan </a>
                                    </li>
                                </ul>
                               
                                <div class="tab-content">
                                <div class="tab-pane active" id="tab_15_1">
                                     
                                      <h3> <b>Neraca PL </b></h3>
                                       
                                         
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>NERACA BANK UMUM</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>PT Bank MNC International, Tbk</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="80%" align="center"  ><b>POS-POS</b></td>
                                                <td width="20%" align="center"  ><b><?php echo $label_tgl; ?></b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(0);
                                                for ($i=9; $i<=97 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="80%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="20%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>

                                            <br>

                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>LAPORAN LABA RUGI BANK UMUM - LRBU
</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>PT Bank MNC International, Tbk</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="80%" align="center"  ><b>POS-POS</b></td>
                                                <td width="20%" align="center"  ><b><?php echo $label_tgl; ?></b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(0);
                                                for ($i=105; $i<=187 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="80%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="20%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>


                                            <br>

                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>KOMITMEN & KONTINJENSI BANK UMUM - KKBU
</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>PT. Bank MNC International, Tbk</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="80%" align="center"  ><b>POS-POS</b></td>
                                                <td width="20%" align="center"  ><b><?php echo $label_tgl; ?></b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                //$objPHPExcel->setActiveSheetIndex(0);
                                                for ($i=196; $i<=241 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="80%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="20%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>



                                        </div>






                                    </div>















                                    <div class="tab-pane " id="tab_15_2">
                                      <h3> <b>KPMM </b></h3>
                                            <h5>
                                            <b>
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
                                               $objPHPExcel->setActiveSheetIndex(1);
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







                                    <div class="tab-pane " id="tab_15_3">
                                     
                                      <h3> <b>KAP </b></h3>
                                       
                                         
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="7" ><b>PT. Bank MNC International, Tbk</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="7" ><b>KUALITAS ASET PRODUKTIF BANK UMUM - KAPBU</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="40%" align="center"  ><b>POS-POS</b></td>
                                                <td width="10%" align="center"  ><b>L</b></td>
                                                <td width="10%" align="center"  ><b>DPK</b></td>
                                                <td width="10%" align="center"  ><b>KL</b></td>
                                                <td width="10%" align="center"  ><b>D</b></td>
                                                <td width="10%" align="center"  ><b>M</b></td>
                                                <td width="10%" align="center"  ><b>Jumlah</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(2);
                                                for ($i=10; $i<=98 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="40%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>


                                            <br>
                                            <br>

                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="5" ><b>PT. Bank MNC International, Tbk</b> <br>
                                                  <b>CADANGAN PENYISIHAN KERUGIAN                
</b>
                                                </td>
                                                </tr>
                                                
                                                <tr class="active">
                                                <td width="40%" align="center" rowspan="2" ><b>POS-POS</b></td>
                                                <td width="30%" align="center" colspan="2" ><b>CKPN</b></td>
                                                <td width="30%" align="center" colspan="2" ><b>PPA wajib dibentuk</b></td>
                                              
                                                </tr>
                                                <tr class="active">
                                                
                                                <td width="15%" align="center"  ><b>Individual</b></td>
                                                <td width="15%" align="center"  ><b>Kolektif</b></td>
                                                <td width="15%" align="center"  ><b>Umum</b></td>
                                                <td width="15%" align="center"  ><b>Khusus</b></td>
                                               
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(2);
                                                for ($i=110; $i<=119 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="40%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>

                                            <br>
                                            <br>

                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                              
                                                <tr >
                                                <td width="40%" align="left"  ><b>I. PIHAK TERKAIT</b></td>
                                                <td width="10%" align="center"  ><b>L</b></td>
                                                <td width="10%" align="center"  ><b>DPK</b></td>
                                                <td width="10%" align="center"  ><b>KL</b></td>
                                                <td width="10%" align="center"  ><b>D</b></td>
                                                <td width="10%" align="center"  ><b>M</b></td>
                                                <td width="10%" align="center"  ><b>Jumlah</b></td>
                                                </tr>
                                                
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(2);
                                                for ($i=123; $i<=125 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="40%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="10%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>
                                            <br>
                                            <br>

                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                
                                                
                                                
                                                <tr class="active">
                                                <td width="40%" align="center" rowspan="2" ><b>POS-POS</b></td>
                                                <td width="30%" align="center" colspan="2" ><b>CKPN</b></td>
                                                <td width="30%" align="center" colspan="2" ><b>PPA wajib dibentuk</b></td>
                                              
                                                </tr>
                                                <tr class="active">
                                                
                                                <td width="15%" align="center"  ><b>Individual</b></td>
                                                <td width="15%" align="center"  ><b>Kolektif</b></td>
                                                <td width="15%" align="center"  ><b>Umum</b></td>
                                                <td width="15%" align="center"  ><b>Khusus</b></td>
                                               
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(2);
                                                for ($i=131; $i<=131 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="40%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
                                              
                                            </tr>
                                               <?php
                                             }
                                               ?>
                                                 
                                                </tbody>
                                            </table>


                                        </div>





                                    </div>

                                    <div class="tab-pane " id="tab_15_4">
                                     
                                      <h3> <b>Informasi Tambahan </b></h3>
                                       
                                         
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>PT. Bank MNC International, Tbk</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="100%" align="center" colspan="2" ><b>INFORMASI TAMBAHAN BANK UMUM - ITBU</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="80%" align="center"  ><b>POS-POS</b></td>
                                                <td width="20%" align="center"  ><b>31 January 2017</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                <?php
                                                $objPHPExcel->setActiveSheetIndex(3);
                                                for ($i=10; $i<=92 ; $i++) { 
                                               

                                                ?>
                                            <tr>
                                              <td align="left" width="80%" > <b> <?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                              <td width="20%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                              
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

