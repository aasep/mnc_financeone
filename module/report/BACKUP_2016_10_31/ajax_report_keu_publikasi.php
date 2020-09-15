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
logActivity("generate publikasi",date('Y_m_d_H_i_s'));
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
$prev_tgl=$tahun."-12-31";
$var_prev_tgl=date('Y-m-d',strtotime(date('Y-m-d',strtotime($prev_tgl))." -1 year "));;


//echo $var_tgl."<br>";
//echo $var_prev_tgl."<br>";
//die();


############################  QUERY BS #########################################

$query =" select sum (a.Nominal) as Nominal from DM_Journal a
left join Referensi_GL_02 b on a.KodeGL=b.GLNO
left join referensi_publikasi c on b.PBLKS_Level_3=c.PBLKS_Level_3 
where a.DataDate ='$var_tgl' ";
$query2 =" select sum (a.Nominal) as Nominal from DM_Journal a
left join Referensi_GL_02 b on a.KodeGL=b.GLNO
left join referensi_publikasi c on b.PBLKS_Level_3=c.PBLKS_Level_3 
where a.DataDate ='$var_prev_tgl' ";



//PBLKS101000001  Kas  sandi lbu = 100
$query_add= " and c.PBLKS_Level_3='PBLKS101000001' and b.Sandi_LBU='100' ";

//echo $query.$query_add;
//die();
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f11=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g11=$row2['Nominal'];
//PBLKS101000002  Penempatan pada Bank indonesia    120
$query_add= " and c.PBLKS_Level_3='PBLKS101000002' and b.Sandi_LBU='120'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f12=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g12=$row2['Nominal'];


//PBLKS101000003  Penempatan pada bank lain      130
$query_add= " and c.PBLKS_Level_3='PBLKS101000003' and b.Sandi_LBU='130' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f13=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g13=$row2['Nominal'];
//PBLKS101000004  Tagihan spot dan derivatif    135
$query_add= " and c.PBLKS_Level_3='PBLKS101000004' and b.Sandi_LBU='135'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f14=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g14=$row2['Nominal'];
//PBLKS101000005  Surat berharga   
$query_add= " and c.PBLKS_Level_3='PBLKS101000005' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f15=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g15=$row2['Nominal'];
//PBLKS101000006  Diukur pada nilai wajar melalui laporan laba/rugi
$query_add= " and c.PBLKS_Level_3='PBLKS101000006' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f16=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g16=$row2['Nominal'];
//PBLKS101000007  Diperdagangkan
$query_add= " and c.PBLKS_Level_3='PBLKS101000007' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$fx=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$gx=$row2['Nominal'];
//PBLKS101000008  Ditetapkan untuk diukur pada nilai wajar
$query_add= " and c.PBLKS_Level_3='PBLKS101000008' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$fx=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$gx=$row2['Nominal'];
//PBLKS101000009  Tersedia untuk dijual     143
$query_add= " and c.PBLKS_Level_3='PBLKS101000009' and b.Sandi_LBU='143' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f17=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g17=$row2['Nominal'];
//PBLKS101000010  Dimiliki hingga jatuh tempo     144
$query_add= " and c.PBLKS_Level_3='PBLKS101000010' and b.Sandi_LBU='144'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f18=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g18=$row2['Nominal'];
//BLKS101000011  Pinjaman yang diberikan dan piutang    145
$query_add= " and c.PBLKS_Level_3='PBLKS101000011' and b.Sandi_LBU='145' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f19=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g19=$row2['Nominal'];
//PBLKS101000012  Surat berharga yang dijual dengan janji dibeli kembali (repo)  160
$query_add= " and c.PBLKS_Level_3='PBLKS101000012' and b.Sandi_LBU='160'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f20=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g20=$row2['Nominal'];
//PBLKS101000013  Tagihan atas surat berharga yang dibeli dengan janji dijual kembali (reverse repo)    164
$query_add= " and c.PBLKS_Level_3='PBLKS101000013' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f21=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g21=$row2['Nominal'];
//PBLKS101000014  Tagihan akseptasi    166
$query_add= " and c.PBLKS_Level_3='PBLKS101000014' and b.Sandi_LBU='166' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f22=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g22=$row2['Nominal'];
//PBLKS101000015  Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS101000015' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f23=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g23=$row2['Nominal'];
//PBLKS101000016       Diukur pada nilai wajar melalui laporan laba/rugi
$query_add= " and c.PBLKS_Level_3='PBLKS101000016' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f24=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g24=$row2['Nominal'];
//PBLKS101000017       Tersedia untuk dijual    172
$query_add= " and c.PBLKS_Level_3='PBLKS101000017' and b.Sandi_LBU='172'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f25=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g25=$row2['Nominal'];
//PBLKS101000018       Dimiliki hingga jatuh tempo   173
$query_add= " and c.PBLKS_Level_3='PBLKS101000018' and b.Sandi_LBU='173' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f26=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g26=$row2['Nominal'];
//PBLKS101000019       Pinjaman yang diberikan dan piutang    175
$query_add= " and c.PBLKS_Level_3='PBLKS101000019' and b.Sandi_LBU='175'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f27=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g27=$row2['Nominal'];
//PBLKS101000020  Pembiayaan syariah ¹    174
$query_add= " and c.PBLKS_Level_3='PBLKS101000020' and b.Sandi_LBU='174'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f28=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g128=$row2['Nominal'];
//PBLKS101000021  Penyertaan    200
$query_add= " and c.PBLKS_Level_3='PBLKS101000021' and b.Sandi_LBU='200' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f29=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g29=$row2['Nominal'];
//PBLKS101000022  Cadangan kerugian penurunan nilai aset keuangan -/-
$query_add= " and c.PBLKS_Level_3='PBLKS101000022' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f30=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g30=$row2['Nominal'];
//PBLKS101000023       Surat berharga    201
$query_add= " and c.PBLKS_Level_3='PBLKS101000023' and b.Sandi_LBU='201' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f31=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g31=$row2['Nominal'];
//PBLKS101000024       Kredit   202 
$query_add= " and c.PBLKS_Level_3='PBLKS101000024' and b.Sandi_LBU='202' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f32=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g32=$row2['Nominal'];
//PBLKS101000025       Lainnya    206
$query_add= " and c.PBLKS_Level_3='PBLKS101000025' and b.Sandi_LBU='206' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f33=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g33=$row2['Nominal'];
//PBLKS101000026  Aset tidak berwujud     212 
$query_add= " and c.PBLKS_Level_3='PBLKS101000026' and b.Sandi_LBU='212' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f34=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g34=$row2['Nominal'];
//PBLKS101000027  Akumulasi amortisasi aset tidak berwujud -/-   213
$query_add= " and c.PBLKS_Level_3='PBLKS101000027' and b.Sandi_LBU='213' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f35=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g35=$row2['Nominal'];
//PBLKS101000028  Aset tetap dan inventaris     214
$query_add= " and c.PBLKS_Level_3='PBLKS101000028' and b.Sandi_LBU='214' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f36=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g36=$row2['Nominal'];
//PBLKS101000029  Akumulasi penyusutan aset tetap dan inventaris -/-   215
$query_add= " and c.PBLKS_Level_3='PBLKS101000029' and b.Sandi_LBU='215' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f37=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g37=$row2['Nominal'];
//PBLKS101000030  Aset non produktif   
$query_add= " and c.PBLKS_Level_3='PBLKS101000030' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f38=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g38=$row2['Nominal'];
//PBLKS101000031       Properti terbengkalai    217
$query_add= " and c.PBLKS_Level_3='PBLKS101000031' and b.Sandi_LBU='217' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f39=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g39=$row2['Nominal'];
//PBLKS101000032       Aset yang diambil alih   218
$query_add= " and c.PBLKS_Level_3='PBLKS101000032' and b.Sandi_LBU='218'";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f40=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g40=$row2['Nominal'];
//PBLKS101000033       Rekening tunda   219  
$query_add= " and c.PBLKS_Level_3='PBLKS101000033' and b.Sandi_LBU='219' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f41=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g41=$row2['Nominal'];

//d. Aset antarkantor ²⁾ 

//i. Melakukan kegiatan operasional di Indonesia          223

//ii. Melakukan kegiatan operasional di luar Indonesia            224

//Cadangan kerugian penurunan nilai aset non keuangan -/-         Form 21
######################################################
//PBLKS101000038  Sewa pembiayaan ¹          227
$query_add= " and c.PBLKS_Level_3='PBLKS101000038' and b.Sandi_LBU='227' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f46=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g46=$row2['Nominal'];
//PBLKS101000039  Aset pajak tangguhan            228
$query_add= " and c.PBLKS_Level_3='PBLKS101000039' and b.Sandi_LBU='228' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f47=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g47=$row2['Nominal'];
//PBLKS101000040  Aset lainnya            230
$query_add= " and c.PBLKS_Level_3='PBLKS101000040' and b.Sandi_LBU='230' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f48=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g48=$row2['Nominal'];
//TOTAL ASET          290

//PBLKS102000001  Giro           300
$query_add= " and c.PBLKS_Level_3='PBLKS102000001' and b.Sandi_LBU='300' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f53=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g53=$row2['Nominal'];
//PBLKS102000002  Tabungan            320
$query_add= " and c.PBLKS_Level_3='PBLKS102000002' and b.Sandi_LBU='320' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f54=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g54=$row2['Nominal'];
//PBLKS102000003  Simpanan berjangka          330
$query_add= " and c.PBLKS_Level_3='PBLKS102000003' and b.Sandi_LBU='330' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f55=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g55=$row2['Nominal'];
//PBLKS102000004  Dana investasi revenue sharing ¹ 
$query_add= " and c.PBLKS_Level_3='PBLKS102000004'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f56=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g56=$row2['Nominal'];          
//PBLKS102000005  Pinjaman dari Bank Indonesia            340
$query_add= " and c.PBLKS_Level_3='PBLKS102000005' and b.Sandi_LBU='340' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f57=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g57=$row2['Nominal'];
//PBLKS102000006  Pinjaman dari bank lain         350
$query_add= " and c.PBLKS_Level_3='PBLKS102000006' and b.Sandi_LBU='350' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f58=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g58=$row2['Nominal'];
//PBLKS102000007  Liabilitas spot dan derivatif           351
$query_add= " and c.PBLKS_Level_3='PBLKS102000007' and b.Sandi_LBU='351' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f59=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g59=$row2['Nominal'];
//PBLKS102000008  Utang atas surat berharga yang dijual dengan janji dibeli kembali (repo)            352
$query_add= " and c.PBLKS_Level_3='PBLKS102000008' and b.Sandi_LBU='352' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f60=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g60=$row2['Nominal'];

//PBLKS102000009  Utang akseptasi         353
$query_add= " and c.PBLKS_Level_3='PBLKS102000009' and b.Sandi_LBU='353' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f61=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g61=$row2['Nominal'];
//PBLKS102000010  Surat berharga yang diterbitkan         355
$query_add= " and c.PBLKS_Level_3='PBLKS102000010' and b.Sandi_LBU='355' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f62=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g62=$row2['Nominal'];
//PBLKS102000011  Pinjaman yang diterima          360
$query_add= " and c.PBLKS_Level_3='PBLKS102000011' and b.Sandi_LBU='360' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f63=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g63=$row2['Nominal'];
//PBLKS102000012  Setoran jaminan   
$query_add= " and c.PBLKS_Level_3='PBLKS102000012' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f64=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g64=$row2['Nominal'];      
//PBLKS102000013  Liabilitas antar kantor ²  
$query_add= " and c.PBLKS_Level_3='PBLKS102000013'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f65=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g65=$row2['Nominal'];       
//PBLKS102000014  Liabilitas antar kantor  i.  Melakukan kegiatan operasional di Indonesiaa          393
$query_add= " and c.PBLKS_Level_3='PBLKS102000014' and b.Sandi_LBU='393' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f66=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g66=$row2['Nominal'];

//echo $query.$query_add;
//die();
//PBLKS102000015  Liabilitas antar kantor ii. Melakukan kegiatan operasional di luar  Indonesia        394
$query_add= " and c.PBLKS_Level_3='PBLKS102000015' and b.Sandi_LBU='394' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f67=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g67=$row2['Nominal'];
//PBLKS102000016  Liabilitas pajak tangguhan          396
$query_add= " and c.PBLKS_Level_3='PBLKS102000016' and b.Sandi_LBU='396' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f68=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g68=$row2['Nominal'];
//PBLKS102000017  Liabilitas lainnya         400
$query_add= " and c.PBLKS_Level_3='PBLKS102000017' and b.Sandi_LBU='400' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f69=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g69=$row2['Nominal'];
//PBLKS103000001  Dana investasi profit sharing ¹            401
$query_add= " and c.PBLKS_Level_3='PBLKS103000001' and b.Sandi_LBU='401' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f70=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g70=$row2['Nominal'];        


//PBLKS103000004        Modal dasar   421
$query_add= " and c.PBLKS_Level_3='PBLKS103000004' and b.Sandi_LBU='421' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f75=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g75=$row2['Nominal'];        

//PBLKS103000005        Modal yang belum disetor -/-      422
$query_add= " and c.PBLKS_Level_3='PBLKS103000005' and b.Sandi_LBU='422' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f76=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g76=$row2['Nominal'];        

//PBLKS103000006        Saham yang dibeli kembali (treasury stock) -/-    423
$query_add= " and c.PBLKS_Level_3='PBLKS103000006' and b.Sandi_LBU='423' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f77=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g77=$row2['Nominal'];        

//PBLKS103000007  Tambahan modal disetor
//PBLKS103000008        Agio   431
$query_add= " and c.PBLKS_Level_3='PBLKS103000008' and b.Sandi_LBU='431' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f79=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g79=$row2['Nominal'];        

//PBLKS103000009        Disagio -/-    432
$query_add= " and c.PBLKS_Level_3='PBLKS103000009' and b.Sandi_LBU='432' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f80=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g80=$row2['Nominal'];        

//PBLKS103000010        Modal sumbangan   433
$query_add= " and c.PBLKS_Level_3='PBLKS103000010' and b.Sandi_LBU='433' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f81=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g81=$row2['Nominal'];        

//PBLKS103000011        Dana setoran modal
$query_add= " and c.PBLKS_Level_3='PBLKS103000011'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f82=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g82=$row2['Nominal'];        

//PBLKS103000012        Lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS103000012'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f83=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g83=$row2['Nominal'];        

//PBLKS103000013  Pendapatan (kerugian) komprehensif lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS103000013'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f84=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g84=$row2['Nominal'];        

//PBLKS103000014        Penyesuaian akibat penjabaran laporan keuangan  436-437
$query_add= " and c.PBLKS_Level_3='PBLKS103000014' and b.Sandi_LBU in ('436','437')";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f85=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g85=$row2['Nominal'];        
/*
//PBLKS103000015        dalam mata uang asing
$query_add= " and c.PBLKS_Level_3='PBLKS103000015'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f85=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g85=$row2['Nominal'];        
*/
//PBLKS103000016        Keuntungan (kerugian) dari perubahan nilai aset
$query_add= " and c.PBLKS_Level_3='PBLKS103000016'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f86=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g86=$row2['Nominal'];        
/*
//PBLKS103000017        keuangan dalam kelompok tersedia untuk dijual
$query_add= " and c.PBLKS_Level_3='PBLKS103000001'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f70=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g70=$row2['Nominal'];        
*/
//PBLKS103000018        Bagian efektif lindung nilai arus kas
$query_add= " and c.PBLKS_Level_3='PBLKS103000018'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f87=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g87=$row2['Nominal'];        


//PBLKS103000019        Keuntungan revaluasi aset tetap  
$query_add= " and c.PBLKS_Level_3='PBLKS103000019'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f88=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g88=$row2['Nominal'];
//PBLKS103000020        Bagian pendapatan komprehensif lain dari entitas    
$query_add= " and c.PBLKS_Level_3='PBLKS103000020'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f89=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g89=$row2['Nominal'];
//PBLKS103000022        Keuntungan (kerugian) aktuarial program imbalan  
$query_add= " and c.PBLKS_Level_3='PBLKS103000022'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f90=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g90=$row2['Nominal'];
//PBLKS103000024        Pajak penghasilan terkait dengan penghasilan  
$query_add= " and c.PBLKS_Level_3='PBLKS103000024'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f91=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g91=$row2['Nominal'];
//PBLKS103000026        Lainnya     (440-445)
$query_add= " and c.PBLKS_Level_3='PBLKS103000026' and b.Sandi_LBU in ('440','441','442','443','444','445') ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f92=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g92=$row2['Nominal'];
//PBLKS103000027  Selisih kuasi reorganisasi ³
$query_add= " and c.PBLKS_Level_3='PBLKS103000027'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f93=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g93=$row2['Nominal'];
//PBLKS103000028  Selisih restrukturisasi entitas sepengendali  457
$query_add= " and c.PBLKS_Level_3='PBLKS103000028'  and b.Sandi_LBU in ('457') ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f94=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g94=$row2['Nominal'];
//PBLKS103000029  Ekuitas lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS103000029'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f95=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g95=$row2['Nominal'];
//PBLKS103000030  Cadangan
$query_add= " and c.PBLKS_Level_3='PBLKS103000030'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f96=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g96=$row2['Nominal'];
//PBLKS103000031        Cadangan umum     (451)
$query_add= " and c.PBLKS_Level_3='PBLKS103000031' and b.Sandi_LBU in ('451') ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f97=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g97=$row2['Nominal'];
//PBLKS103000032        Cadangan tujuan   (452)
$query_add= " and c.PBLKS_Level_3='PBLKS103000032'  and b.Sandi_LBU in ('452')";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f98=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g98=$row2['Nominal'];
//PBLKS103000033  Laba (Rugi)
$query_add= " and c.PBLKS_Level_3='PBLKS103000033'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f99=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g99=$row2['Nominal'];
//PBLKS103000034        Tahun-tahun lalu  (461,462)
$query_add= " and c.PBLKS_Level_3='PBLKS103000034' and b.Sandi_LBU in ('461','462') ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f100=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g100=$row2['Nominal'];
//PBLKS103000035        Tahun berjalan    (465,466)
$query_add= " and c.PBLKS_Level_3='PBLKS103000035' and b.Sandi_LBU in ('465','466') ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f101=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g101=$row2['Nominal'];
//PBLKS103000036  Kepentingan non pengendali 6)
$query_add= " and c.PBLKS_Level_3='PBLKS103000036'  ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$f105=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$g105=$row2['Nominal'];


################################################################################






// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

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

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder3 = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);


$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A11:G109')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(25);


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:G1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:G2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:G3');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B7:D8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F7:G7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A7:A8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E7:E8');


for ($i=11; $i <=109 ; $i++) { 
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B$i:D$i");
}






$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A10')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per $var_tgl dan $var_prev_tgl ");

$objPHPExcel->getActiveSheet()->setCellValue('A7', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'POS - POS');
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Sandi LBU');
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'BANK');
$objPHPExcel->getActiveSheet()->setCellValue('F8', "$var_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "$var_prev_tgl");

$objPHPExcel->getActiveSheet()->setCellValue('A10', 'ASET');

$objPHPExcel->getActiveSheet()->setCellValue('A11', '1.');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Kas');
$objPHPExcel->getActiveSheet()->setCellValue('E11', 100);
$objPHPExcel->getActiveSheet()->setCellValue('F11', $f11);
$objPHPExcel->getActiveSheet()->setCellValue('G11', $g11);
$objPHPExcel->getActiveSheet()->setCellValue('A12', '2.');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Penempatan pada Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('E12', 120);
$objPHPExcel->getActiveSheet()->setCellValue('F12', $f12);
$objPHPExcel->getActiveSheet()->setCellValue('G12', $g12);
$objPHPExcel->getActiveSheet()->setCellValue('A13', '3.');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Penempatan pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('E13', 130);
$objPHPExcel->getActiveSheet()->setCellValue('F13', $f13);
$objPHPExcel->getActiveSheet()->setCellValue('G13', $g13);
$objPHPExcel->getActiveSheet()->setCellValue('A14', '4.');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Tagihan spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('E14', 135);
$objPHPExcel->getActiveSheet()->setCellValue('F14', $f14);
$objPHPExcel->getActiveSheet()->setCellValue('G14', $g14);
$objPHPExcel->getActiveSheet()->setCellValue('A15', '5.');
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('E15', "");
$objPHPExcel->getActiveSheet()->setCellValue('F15', $f15);
$objPHPExcel->getActiveSheet()->setCellValue('G15', $g15);

$objPHPExcel->getActiveSheet()->setCellValue('A16', '');
$objPHPExcel->getActiveSheet()->setCellValue('B16', 'a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('F16', $f16);
$objPHPExcel->getActiveSheet()->setCellValue('G16', $g16);

$objPHPExcel->getActiveSheet()->setCellValue('E16', "");
$objPHPExcel->getActiveSheet()->setCellValue('A17', '');
$objPHPExcel->getActiveSheet()->setCellValue('B17', 'b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('E17', 143);
$objPHPExcel->getActiveSheet()->setCellValue('F17', $f17);
$objPHPExcel->getActiveSheet()->setCellValue('G17', $g17);
$objPHPExcel->getActiveSheet()->setCellValue('A18', '');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('E18', 144);
$objPHPExcel->getActiveSheet()->setCellValue('F18', $f18);
$objPHPExcel->getActiveSheet()->setCellValue('G18', $g18);
$objPHPExcel->getActiveSheet()->setCellValue('A19', '');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'd. Pinjaman yang diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('E19', 145);
$objPHPExcel->getActiveSheet()->setCellValue('F19', $f19);
$objPHPExcel->getActiveSheet()->setCellValue('G19', $g19);

$objPHPExcel->getActiveSheet()->setCellValue('A20', '6.');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Surat berharga yang dijual dengan janji dibeli kembali (repo)');
$objPHPExcel->getActiveSheet()->setCellValue('E20', 160);
$objPHPExcel->getActiveSheet()->setCellValue('F20', $f20);
$objPHPExcel->getActiveSheet()->setCellValue('G20', $g20);
$objPHPExcel->getActiveSheet()->setCellValue('A21', '7.');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Tagihan atas surat berharga yang dibeli dengan janji dijual kembali (reverse repo)');
$objPHPExcel->getActiveSheet()->setCellValue('E21', 164);
$objPHPExcel->getActiveSheet()->setCellValue('F21', $f21);
$objPHPExcel->getActiveSheet()->setCellValue('G21', $g21);
$objPHPExcel->getActiveSheet()->setCellValue('A22', '8.');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Tagihan akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('E22', 166);
$objPHPExcel->getActiveSheet()->setCellValue('F22', $f22);
$objPHPExcel->getActiveSheet()->setCellValue('G22', $g22);
$objPHPExcel->getActiveSheet()->setCellValue('A23', '9.');
$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Kredit ');
$objPHPExcel->getActiveSheet()->setCellValue('E23', "");
$objPHPExcel->getActiveSheet()->setCellValue('F23', $f23);
$objPHPExcel->getActiveSheet()->setCellValue('G23', $g23);
$objPHPExcel->getActiveSheet()->setCellValue('A24', '');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('E24', "");
$objPHPExcel->getActiveSheet()->setCellValue('F24', $f24);
$objPHPExcel->getActiveSheet()->setCellValue('G24', $g24);
$objPHPExcel->getActiveSheet()->setCellValue('A25', '');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('E25', 172);
$objPHPExcel->getActiveSheet()->setCellValue('F25', $f25);
$objPHPExcel->getActiveSheet()->setCellValue('G25', $g25);
$objPHPExcel->getActiveSheet()->setCellValue('A26', '');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('E26', 173);
$objPHPExcel->getActiveSheet()->setCellValue('F26', $f26);
$objPHPExcel->getActiveSheet()->setCellValue('G26', $g26);
$objPHPExcel->getActiveSheet()->setCellValue('A27', '');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'd. Pinjaman yang diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('E27', 175);
$objPHPExcel->getActiveSheet()->setCellValue('F27', $f27);
$objPHPExcel->getActiveSheet()->setCellValue('G27', $g27);

$objPHPExcel->getActiveSheet()->setCellValue('A28', '10.');
$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Pembiayaan syariah ¹⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E28', 174);
$objPHPExcel->getActiveSheet()->setCellValue('F28', $f28);
$objPHPExcel->getActiveSheet()->setCellValue('G28', $g28);
$objPHPExcel->getActiveSheet()->setCellValue('A29', '11.');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Penyertaan ');
$objPHPExcel->getActiveSheet()->setCellValue('E29', 200);
$objPHPExcel->getActiveSheet()->setCellValue('F29', $f29);
$objPHPExcel->getActiveSheet()->setCellValue('G29', $g29);
$objPHPExcel->getActiveSheet()->setCellValue('A30', '12.');
$objPHPExcel->getActiveSheet()->setCellValue('B30', 'Cadangan kerugian penurunan nilai aset keuangan -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E30', "");
$objPHPExcel->getActiveSheet()->setCellValue('F30', $f30);
$objPHPExcel->getActiveSheet()->setCellValue('G30', $g30);

$objPHPExcel->getActiveSheet()->setCellValue('A31', '');
$objPHPExcel->getActiveSheet()->setCellValue('B31', 'a. Surat Berharga');
$objPHPExcel->getActiveSheet()->setCellValue('E31', 201);
$objPHPExcel->getActiveSheet()->setCellValue('F31', $f31);
$objPHPExcel->getActiveSheet()->setCellValue('G31', $g31);
$objPHPExcel->getActiveSheet()->setCellValue('A32', '');
$objPHPExcel->getActiveSheet()->setCellValue('B32', 'b. Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('E32', 202);
$objPHPExcel->getActiveSheet()->setCellValue('F32', $f32);
$objPHPExcel->getActiveSheet()->setCellValue('G32', $g32);
$objPHPExcel->getActiveSheet()->setCellValue('A33', '');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'c. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('E33', 206);
$objPHPExcel->getActiveSheet()->setCellValue('F33', $f33);
$objPHPExcel->getActiveSheet()->setCellValue('G33', $g33);

$objPHPExcel->getActiveSheet()->setCellValue('A34', '13.');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Aset Tidak Berwujud');
$objPHPExcel->getActiveSheet()->setCellValue('E34', 212);
$objPHPExcel->getActiveSheet()->setCellValue('F34', $f34);
$objPHPExcel->getActiveSheet()->setCellValue('G43', $g34);
$objPHPExcel->getActiveSheet()->setCellValue('A35', '');
$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Akumulasi amortisasi aset tidak berwujud -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E35', 213);
$objPHPExcel->getActiveSheet()->setCellValue('F35', $f35);
$objPHPExcel->getActiveSheet()->setCellValue('G53', $g35);

$objPHPExcel->getActiveSheet()->setCellValue('A36', '14.');
$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Aset tetap dan inventaris');
$objPHPExcel->getActiveSheet()->setCellValue('E36', 214);
$objPHPExcel->getActiveSheet()->setCellValue('F36', $f36);
$objPHPExcel->getActiveSheet()->setCellValue('G36', $g36);
$objPHPExcel->getActiveSheet()->setCellValue('A37', '');
$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Akumulasi penyusutan aset tetap dan inventaris -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E37', 215);
$objPHPExcel->getActiveSheet()->setCellValue('F37', $f37);
$objPHPExcel->getActiveSheet()->setCellValue('G37', $g37);
$objPHPExcel->getActiveSheet()->setCellValue('A38', '15.');
$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Aset non produktif');
$objPHPExcel->getActiveSheet()->setCellValue('E38', "");
$objPHPExcel->getActiveSheet()->setCellValue('F38', $f38);
$objPHPExcel->getActiveSheet()->setCellValue('G38', $g38);

$objPHPExcel->getActiveSheet()->setCellValue('A39', '');
$objPHPExcel->getActiveSheet()->setCellValue('B39', 'a. Properti terbengkalai');
$objPHPExcel->getActiveSheet()->setCellValue('E39', 217);
$objPHPExcel->getActiveSheet()->setCellValue('F39', $f39);
$objPHPExcel->getActiveSheet()->setCellValue('G39', $g39);
$objPHPExcel->getActiveSheet()->setCellValue('A40', '');
$objPHPExcel->getActiveSheet()->setCellValue('B40', 'b. Aset yang diambil alih');
$objPHPExcel->getActiveSheet()->setCellValue('E40', 218);
$objPHPExcel->getActiveSheet()->setCellValue('F40', $f40);
$objPHPExcel->getActiveSheet()->setCellValue('G40', $g40);
$objPHPExcel->getActiveSheet()->setCellValue('A41', '');
$objPHPExcel->getActiveSheet()->setCellValue('B41', 'c. Rekening tunda');
$objPHPExcel->getActiveSheet()->setCellValue('E41', 219);
$objPHPExcel->getActiveSheet()->setCellValue('F41', $f41);
$objPHPExcel->getActiveSheet()->setCellValue('G41', $g41);
$objPHPExcel->getActiveSheet()->setCellValue('A42', '');
$objPHPExcel->getActiveSheet()->setCellValue('B42', 'd. Aset antarkantor ²⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E42', "");
$objPHPExcel->getActiveSheet()->setCellValue('F42', $f42);
$objPHPExcel->getActiveSheet()->setCellValue('G42', $g42);


$objPHPExcel->getActiveSheet()->setCellValue('B43', "i. Melakukan kegiatan operasional di Indonesia ");
$objPHPExcel->getActiveSheet()->setCellValue('E43', 223);
$objPHPExcel->getActiveSheet()->setCellValue('F43', $f43);
$objPHPExcel->getActiveSheet()->setCellValue('G43', $g43);
$objPHPExcel->getActiveSheet()->setCellValue('B44', "ii. Melakukan kegiatan operasional di luar Indonesia ");
$objPHPExcel->getActiveSheet()->setCellValue('E44', 224);

$objPHPExcel->getActiveSheet()->setCellValue('A45', '16.');
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'Cadangan kerugian penurunan nilai aset non keuangan -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E45', "Form 21");
$objPHPExcel->getActiveSheet()->setCellValue('F45', $f45);
$objPHPExcel->getActiveSheet()->setCellValue('G45', $g45);
$objPHPExcel->getActiveSheet()->setCellValue('A46', '17.');
$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Sewa pembiayaan ¹⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E46', 227);
$objPHPExcel->getActiveSheet()->setCellValue('F46', $f46);
$objPHPExcel->getActiveSheet()->setCellValue('G46', $g46);
$objPHPExcel->getActiveSheet()->setCellValue('A47', '18.');
$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Aset pajak tangguhan ');
$objPHPExcel->getActiveSheet()->setCellValue('E47', 228);
$objPHPExcel->getActiveSheet()->setCellValue('F47', $f47);
$objPHPExcel->getActiveSheet()->setCellValue('G47', $g47);
$objPHPExcel->getActiveSheet()->setCellValue('A48', '19.');
$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Aset lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('E48', 230);
$objPHPExcel->getActiveSheet()->setCellValue('F48', $f48);
$objPHPExcel->getActiveSheet()->setCellValue('G48', $g48);

$objPHPExcel->getActiveSheet()->setCellValue('B50', "TOTAL ASET");
$objPHPExcel->getActiveSheet()->setCellValue('E50', 290);


$objPHPExcel->getActiveSheet()->setCellValue('A51', "LIABILITAS DAN EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "LIABILITAS");

$objPHPExcel->getActiveSheet()->setCellValue('A53', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "Giro");
$objPHPExcel->getActiveSheet()->setCellValue('E53', 300);
$objPHPExcel->getActiveSheet()->setCellValue('F53', $f53);
$objPHPExcel->getActiveSheet()->setCellValue('G53', $g53);
$objPHPExcel->getActiveSheet()->setCellValue('A54', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "Tabungan");
$objPHPExcel->getActiveSheet()->setCellValue('E54', 320);
$objPHPExcel->getActiveSheet()->setCellValue('F54', $f54);
$objPHPExcel->getActiveSheet()->setCellValue('G54', $g54);
$objPHPExcel->getActiveSheet()->setCellValue('A55', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "Simpanan berjangka");
$objPHPExcel->getActiveSheet()->setCellValue('E55', 330);
$objPHPExcel->getActiveSheet()->setCellValue('F55', $f55);
$objPHPExcel->getActiveSheet()->setCellValue('G55', $g55);
$objPHPExcel->getActiveSheet()->setCellValue('A56', "4.");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "Dana investasi revenue sharing ¹⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E56', "");
$objPHPExcel->getActiveSheet()->setCellValue('F56', $f56);
$objPHPExcel->getActiveSheet()->setCellValue('G56', $g56);
$objPHPExcel->getActiveSheet()->setCellValue('A57', "5.");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "Pinjaman dari Bank Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E57', 340);
$objPHPExcel->getActiveSheet()->setCellValue('F57', $f57);
$objPHPExcel->getActiveSheet()->setCellValue('G57', $g57);
$objPHPExcel->getActiveSheet()->setCellValue('A58', "6.");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "Pinjaman dari bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('E58', 350);
$objPHPExcel->getActiveSheet()->setCellValue('F58', $f58);
$objPHPExcel->getActiveSheet()->setCellValue('G58', $g58);
$objPHPExcel->getActiveSheet()->setCellValue('A59', "7.");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "Liabilitas spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('E59', 351);
$objPHPExcel->getActiveSheet()->setCellValue('F59', $f59);
$objPHPExcel->getActiveSheet()->setCellValue('G59', $g59);
$objPHPExcel->getActiveSheet()->setCellValue('A60', "8.");
$objPHPExcel->getActiveSheet()->setCellValue('B60', "Utang atas surat berharga yang dijual dengan janji dibeli kembali (repo)");
$objPHPExcel->getActiveSheet()->setCellValue('E60', 352);
$objPHPExcel->getActiveSheet()->setCellValue('F60', $f60);
$objPHPExcel->getActiveSheet()->setCellValue('G60', $g60);
$objPHPExcel->getActiveSheet()->setCellValue('A61', "9.");
$objPHPExcel->getActiveSheet()->setCellValue('B61', "Utang akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('E61', 353);
$objPHPExcel->getActiveSheet()->setCellValue('F61', $f61);
$objPHPExcel->getActiveSheet()->setCellValue('G61', $g61);
$objPHPExcel->getActiveSheet()->setCellValue('A62', "10.");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "Surat berharga yang diterbitkan");
$objPHPExcel->getActiveSheet()->setCellValue('E62', 355);
$objPHPExcel->getActiveSheet()->setCellValue('F62', $f62);
$objPHPExcel->getActiveSheet()->setCellValue('G62', $g62);
$objPHPExcel->getActiveSheet()->setCellValue('A63', "11.");
$objPHPExcel->getActiveSheet()->setCellValue('B63', "Pinjaman yang diterima");
$objPHPExcel->getActiveSheet()->setCellValue('E63', 360);
$objPHPExcel->getActiveSheet()->setCellValue('F63', $f63);
$objPHPExcel->getActiveSheet()->setCellValue('G63', $g63);
$objPHPExcel->getActiveSheet()->setCellValue('A64', "12.");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "Setoran jaminan");
$objPHPExcel->getActiveSheet()->setCellValue('E64', "");
$objPHPExcel->getActiveSheet()->setCellValue('F64', $f64);
$objPHPExcel->getActiveSheet()->setCellValue('G64', $g64);
$objPHPExcel->getActiveSheet()->setCellValue('A65', "13.");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "Liabilitas antar kantor ²⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E65', "");
$objPHPExcel->getActiveSheet()->setCellValue('F65', $f65);
$objPHPExcel->getActiveSheet()->setCellValue('G65', $g65);
$objPHPExcel->getActiveSheet()->setCellValue('B66', "a. Melakukan kegiatan operasional di Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E66', 393);
$objPHPExcel->getActiveSheet()->setCellValue('F66', $f66);
$objPHPExcel->getActiveSheet()->setCellValue('G66', $g66);
$objPHPExcel->getActiveSheet()->setCellValue('B67', "b. Melakukan kegiatan operasional di luar Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E67', 394);
$objPHPExcel->getActiveSheet()->setCellValue('F67', $f67);
$objPHPExcel->getActiveSheet()->setCellValue('G67', $g67);
$objPHPExcel->getActiveSheet()->setCellValue('A68', "14.");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "Liabilitas pajak tangguhan");
$objPHPExcel->getActiveSheet()->setCellValue('E68', 396);
$objPHPExcel->getActiveSheet()->setCellValue('F68', $f68);
$objPHPExcel->getActiveSheet()->setCellValue('G68', $g68);
$objPHPExcel->getActiveSheet()->setCellValue('A69', "15.");
$objPHPExcel->getActiveSheet()->setCellValue('B69', "Liabilitas lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E69', 400);
$objPHPExcel->getActiveSheet()->setCellValue('F69', $f69);
$objPHPExcel->getActiveSheet()->setCellValue('G69', $g69);
$objPHPExcel->getActiveSheet()->setCellValue('A70', "16.");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "Dana investasi profit sharing ¹⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E70', 401);
$objPHPExcel->getActiveSheet()->setCellValue('F70', $f70);
$objPHPExcel->getActiveSheet()->setCellValue('G70', $g70);

$objPHPExcel->getActiveSheet()->setCellValue('B71', "TOTAL LIABILITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E71', 401);


$objPHPExcel->getActiveSheet()->setCellValue('B73', "EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('A74', "17");
$objPHPExcel->getActiveSheet()->setCellValue('B74', "Modal disetor");
$objPHPExcel->getActiveSheet()->setCellValue('E74', "");
$objPHPExcel->getActiveSheet()->setCellValue('F74', $f74);
$objPHPExcel->getActiveSheet()->setCellValue('G74', $g74);
$objPHPExcel->getActiveSheet()->setCellValue('B75', "a. Modal Dasar");
$objPHPExcel->getActiveSheet()->setCellValue('E75', "421");
$objPHPExcel->getActiveSheet()->setCellValue('F75', $f75);
$objPHPExcel->getActiveSheet()->setCellValue('G75', $g75);
$objPHPExcel->getActiveSheet()->setCellValue('B76', "b. Modal yang belum disetor -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E76', "422");
$objPHPExcel->getActiveSheet()->setCellValue('F76', $f76);
$objPHPExcel->getActiveSheet()->setCellValue('G76', $g76);
$objPHPExcel->getActiveSheet()->setCellValue('B77', "c. Saham yang dibeli kembali (treasury stock) -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E77', "423");
$objPHPExcel->getActiveSheet()->setCellValue('F77', $f77);
$objPHPExcel->getActiveSheet()->setCellValue('G77', $g77);
$objPHPExcel->getActiveSheet()->setCellValue('A78', "18");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Tambahan Modal disetor");
$objPHPExcel->getActiveSheet()->setCellValue('E78', "");
$objPHPExcel->getActiveSheet()->setCellValue('F78', $f78);
$objPHPExcel->getActiveSheet()->setCellValue('G78', $g78);
$objPHPExcel->getActiveSheet()->setCellValue('B79', "a. Agio");
$objPHPExcel->getActiveSheet()->setCellValue('E79', "431");
$objPHPExcel->getActiveSheet()->setCellValue('F79', $f79);
$objPHPExcel->getActiveSheet()->setCellValue('G79', $g79);
$objPHPExcel->getActiveSheet()->setCellValue('B80', "b. Disagio -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E80', "432");
$objPHPExcel->getActiveSheet()->setCellValue('F80', $f80);
$objPHPExcel->getActiveSheet()->setCellValue('G80', $g80);
$objPHPExcel->getActiveSheet()->setCellValue('B81', "c. Modal sumbangan");
$objPHPExcel->getActiveSheet()->setCellValue('E81', "433");
$objPHPExcel->getActiveSheet()->setCellValue('F81', $f81);
$objPHPExcel->getActiveSheet()->setCellValue('G81', $g81);
$objPHPExcel->getActiveSheet()->setCellValue('B82', "d. Dana setoran modal");
$objPHPExcel->getActiveSheet()->setCellValue('E82', "");
$objPHPExcel->getActiveSheet()->setCellValue('F82', $f82);
$objPHPExcel->getActiveSheet()->setCellValue('G82', $g82);
$objPHPExcel->getActiveSheet()->setCellValue('B83', "e. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E83', "");
$objPHPExcel->getActiveSheet()->setCellValue('F83', $f83);
$objPHPExcel->getActiveSheet()->setCellValue('G83', $g83);
$objPHPExcel->getActiveSheet()->setCellValue('A84', "19");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "Pendapatan (kerugian) komprehensif lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E84', "");
$objPHPExcel->getActiveSheet()->setCellValue('F84', $f84);
$objPHPExcel->getActiveSheet()->setCellValue('G84', $g84);
$objPHPExcel->getActiveSheet()->setCellValue('B85', "a. Penyesuaian akibat penjabaran laporan keuangan dfalam mata uang asing");
$objPHPExcel->getActiveSheet()->setCellValue('F85', $f85);
$objPHPExcel->getActiveSheet()->setCellValue('G85', $g85);
$objPHPExcel->getActiveSheet()->setCellValue('E85', "436-437");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual");
$objPHPExcel->getActiveSheet()->setCellValue('E86', "");
$objPHPExcel->getActiveSheet()->setCellValue('F86', $f86);
$objPHPExcel->getActiveSheet()->setCellValue('G86', $g86);
$objPHPExcel->getActiveSheet()->setCellValue('B87', "c. Bagian efektif lindung nilai arus kas");
$objPHPExcel->getActiveSheet()->setCellValue('E87', "");
$objPHPExcel->getActiveSheet()->setCellValue('F87', $f87);
$objPHPExcel->getActiveSheet()->setCellValue('G87', $g87);
$objPHPExcel->getActiveSheet()->setCellValue('B88', "d. Keuntungan revaluasi aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('E88', "");
$objPHPExcel->getActiveSheet()->setCellValue('F88', $f88);
$objPHPExcel->getActiveSheet()->setCellValue('G88', $g88);
$objPHPExcel->getActiveSheet()->setCellValue('B89', "e. Bagian pendapatan komprehensif lain dari entitas asosi");
$objPHPExcel->getActiveSheet()->setCellValue('E89', "");
$objPHPExcel->getActiveSheet()->setCellValue('F89', $f89);
$objPHPExcel->getActiveSheet()->setCellValue('G89', $g89);
$objPHPExcel->getActiveSheet()->setCellValue('B90', "f. Keuntungan (kerugian) aktuarial program imbalan pasti");
$objPHPExcel->getActiveSheet()->setCellValue('E90', "");
$objPHPExcel->getActiveSheet()->setCellValue('F90', $f90);
$objPHPExcel->getActiveSheet()->setCellValue('G90', $g90);
$objPHPExcel->getActiveSheet()->setCellValue('B91', "g. Pajak penghasilan terkait dengan penghasilan komprehensif lain");
$objPHPExcel->getActiveSheet()->setCellValue('E91', "");
$objPHPExcel->getActiveSheet()->setCellValue('F91', $f91);
$objPHPExcel->getActiveSheet()->setCellValue('G91', $g91);
$objPHPExcel->getActiveSheet()->setCellValue('B92', "h. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E92', "440-445");
$objPHPExcel->getActiveSheet()->setCellValue('F92', $f92);
$objPHPExcel->getActiveSheet()->setCellValue('G92', $g92);
$objPHPExcel->getActiveSheet()->setCellValue('A93', "20");
$objPHPExcel->getActiveSheet()->setCellValue('B93', "Selisih kuasi reorganisasi ³⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E93', "");
$objPHPExcel->getActiveSheet()->setCellValue('F93', $f93);
$objPHPExcel->getActiveSheet()->setCellValue('G93', $g93);
$objPHPExcel->getActiveSheet()->setCellValue('A94', "21");
$objPHPExcel->getActiveSheet()->setCellValue('B94', "Selisih restrukturisasi entitas sepengendali");
$objPHPExcel->getActiveSheet()->setCellValue('E94', "457");
$objPHPExcel->getActiveSheet()->setCellValue('F94', $f94);
$objPHPExcel->getActiveSheet()->setCellValue('G94', $g94);
$objPHPExcel->getActiveSheet()->setCellValue('A95', "22");
$objPHPExcel->getActiveSheet()->setCellValue('B95', "Ekuitas lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E95', "");
$objPHPExcel->getActiveSheet()->setCellValue('F95', $f95);
$objPHPExcel->getActiveSheet()->setCellValue('G95', $g95);
$objPHPExcel->getActiveSheet()->setCellValue('A96', "23");
$objPHPExcel->getActiveSheet()->setCellValue('B96', "Cadangan");
$objPHPExcel->getActiveSheet()->setCellValue('E96', "");
$objPHPExcel->getActiveSheet()->setCellValue('F96', $f96);
$objPHPExcel->getActiveSheet()->setCellValue('G96', $g96);
$objPHPExcel->getActiveSheet()->setCellValue('B97', "a. Cadangan Umum");
$objPHPExcel->getActiveSheet()->setCellValue('E97', "451");
$objPHPExcel->getActiveSheet()->setCellValue('F97', $f97);
$objPHPExcel->getActiveSheet()->setCellValue('G97', $g97);
$objPHPExcel->getActiveSheet()->setCellValue('B98', "b. Cadangan Tujuan");
$objPHPExcel->getActiveSheet()->setCellValue('E98', "452");
$objPHPExcel->getActiveSheet()->setCellValue('F98', $f98);
$objPHPExcel->getActiveSheet()->setCellValue('G98', $g98);
$objPHPExcel->getActiveSheet()->setCellValue('A99', "24");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "laba (Rugi)");
$objPHPExcel->getActiveSheet()->setCellValue('E99', "");
$objPHPExcel->getActiveSheet()->setCellValue('F99', $f99);
$objPHPExcel->getActiveSheet()->setCellValue('G99', $g99);
$objPHPExcel->getActiveSheet()->setCellValue('B100', "a. Tahun-tahun Lalu");
$objPHPExcel->getActiveSheet()->setCellValue('E100', "461-462");
$objPHPExcel->getActiveSheet()->setCellValue('F100', $f100);
$objPHPExcel->getActiveSheet()->setCellValue('G100', $g100);
$objPHPExcel->getActiveSheet()->setCellValue('B101', "b. Tahun Berjalan");
$objPHPExcel->getActiveSheet()->setCellValue('E101', "465-466");
$objPHPExcel->getActiveSheet()->setCellValue('F101', $f101);
$objPHPExcel->getActiveSheet()->setCellValue('G101', $g101);
$objPHPExcel->getActiveSheet()->setCellValue('B102', "TOTAL EKUITAS YANG DAPAT DIATRIBUSIKAN ");
$objPHPExcel->getActiveSheet()->setCellValue('E102', "");
$objPHPExcel->getActiveSheet()->setCellValue('B103', "KEPADA PEMILIK");

$objPHPExcel->getActiveSheet()->setCellValue('A105', "25");
$objPHPExcel->getActiveSheet()->setCellValue('B105', "Kepentingan non pengendali 6)");
$objPHPExcel->getActiveSheet()->setCellValue('E105', "");
$objPHPExcel->getActiveSheet()->setCellValue('F105', $f105);
$objPHPExcel->getActiveSheet()->setCellValue('G105', $g105);

$objPHPExcel->getActiveSheet()->setCellValue('B107', "TOTAL EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E107', "");


$objPHPExcel->getActiveSheet()->setCellValue('B109', "TOTAL LIABILITAS DAN EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E109', 490);

// SHEET 1 (BS)
$objPHPExcel->getActiveSheet()->getStyle('F11:G109')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');


$objPHPExcel->getActiveSheet()->setTitle('BS');




//SHEET 2 (PL)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(90);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C7:D7');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D103')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A1:A3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:D8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A9:D9')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A10:D10')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A18:D18')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A11:D11')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A14:D14')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A19:D19')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A36:D36')->applyFromArray($styleArrayFontBold);


$query =" select sum (a.Nominal) as Nominal from DM_Journal a
left join Referensi_GL_02 b on a.KodeGL=b.GLNO
left join referensi_publikasi c on b.PBLKS_Level_3=c.PBLKS_Level_3 
where a.DataDate ='$var_tgl' ";
$query2 =" select sum (a.Nominal) as Nominal from DM_Journal a
left join Referensi_GL_02 b on a.KodeGL=b.GLNO
left join referensi_publikasi c on b.PBLKS_Level_3=c.PBLKS_Level_3 
where a.DataDate ='$var_prev_tgl' ";



//PBLKS201000001  Pendapatan Bunga  
//PBLKS201000002        Rupiah 

$query_add= " and c.PBLKS_Level_3='PBLKS201000002' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c12=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d12=$row2['Nominal'];
//PBLKS201000003        Valuta Asing
$query_add= " and c.PBLKS_Level_3='PBLKS201000003' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c13=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d13=$row2['Nominal'];
//PBLKS202000001  Beban Bunga  
//PBLKS202000002        Rupiah
$query_add= " and c.PBLKS_Level_3='PBLKS202000002' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c15=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d15=$row2['Nominal'];
//PBLKS202000003        Valuta Asing
$query_add= " and c.PBLKS_Level_3='PBLKS202000003' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c16=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d16=$row2['Nominal'];
//PBLKS201000004  Pendapatan (Beban) Bunga Bersih
$query_add= " and c.PBLKS_Level_3='PBLKS201000004' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c17=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d17=$row2['Nominal'];
//PBLKS201000005  Pendapatan dan Beban Operasional selain Bunga
//PBLKS201000006  Pendapatan Operasional Selain Bunga
//PBLKS201000007  Peningkatan nilai wajar aset keuangan
$query_add= " and c.PBLKS_Level_3='PBLKS201000007' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c20=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d20=$row2['Nominal'];
//PBLKS201000008        Surat berharga
$query_add= " and c.PBLKS_Level_3='PBLKS201000008' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c21=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d21=$row2['Nominal'];
//PBLKS201000009        Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS201000009' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c22=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d22=$row2['Nominal'];
//PBLKS201000010        Spot dan derivatif
$query_add= " and c.PBLKS_Level_3='PBLKS201000010' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c23=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d23=$row2['Nominal'];
//PBLKS201000011        Aset keuangan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS201000011' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c24=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d24=$row2['Nominal'];

//PBLKS201000012  Penurunan nilai wajar liabilitas keuangan
$query_add= " and c.PBLKS_Level_3='PBLKS201000012' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c25=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d25=$row2['Nominal'];
//PBLKS201000013  Keuntungan penjualan aset keuangan
$query_add= " and c.PBLKS_Level_3='PBLKS201000013' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c26=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d26=$row2['Nominal'];
//PBLKS201000014        Surat berharga
$query_add= " and c.PBLKS_Level_3='PBLKS201000014' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c27=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d27=$row2['Nominal'];
//PBLKS201000015        Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS201000015' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c28=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d28=$row2['Nominal'];
//PBLKS201000016        Aset keuangan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS201000016' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c29=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d29=$row2['Nominal'];
//PBLKS201000017  Keuntungan transaksi spot dan derivatif (realised)
$query_add= " and c.PBLKS_Level_3='PBLKS201000017' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c30=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d30=$row2['Nominal'];
//PBLKS201000018  Deviden   
$query_add= " and c.PBLKS_Level_3='PBLKS201000018' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c31=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d31=$row2['Nominal'];
//PBLKS201000019  Keuntungan dari Penyertaan dengan equity Method
$query_add= " and c.PBLKS_Level_3='PBLKS201000019' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c32=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d32=$row2['Nominal'];
//PBLKS201000020  Komisi/provisi/fee dan administrasi 
$query_add= " and c.PBLKS_Level_3='PBLKS201000020' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c33=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d33=$row2['Nominal']; 
//PBLKS201000021  Pemulihan atas cadangan kerugian penurunan nilai
$query_add= " and c.PBLKS_Level_3='PBLKS201000021' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c34=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d34=$row2['Nominal'];
//PBLKS201000022  Pendapatan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS201000022' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c35=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d35=$row2['Nominal'];

//PBLKS202000004  Beban Operasional Selain Bunga
//PBLKS202000005  Penurunan nilai wajar aset keuangan (mark to market)
$query_add= " and c.PBLKS_Level_3='PBLKS202000005' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c37=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d37=$row2['Nominal'];
//PBLKS202000006        Surat berharga
$query_add= " and c.PBLKS_Level_3='PBLKS202000006' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c38=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d38=$row2['Nominal'];
//PBLKS202000007        Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS202000007' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c39=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d39=$row2['Nominal'];
//PBLKS202000008        Spot dan derivatif
$query_add= " and c.PBLKS_Level_3='PBLKS202000008' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c40=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d40=$row2['Nominal'];
//PBLKS202000009        Aset keuangan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS202000009' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c41=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d41=$row2['Nominal'];
//PBLKS202000010  Peningkatan nilai wajar liabilitas keuangan (mart to market)
$query_add= " and c.PBLKS_Level_3='PBLKS202000010' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c42=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d42=$row2['Nominal'];
//PBLKS202000011  Kerugian penjualan aset keuangan
$query_add= " and c.PBLKS_Level_3='PBLKS202000011' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c43=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d43=$row2['Nominal'];
//PBLKS202000012        Surat berharga
$query_add= " and c.PBLKS_Level_3='PBLKS202000012' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c44=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d44=$row2['Nominal'];
//PBLKS202000013        Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS202000013' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c45=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d45=$row2['Nominal'];
//PBLKS202000014        Aset keuangan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS202000014' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c46=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d46=$row2['Nominal'];
//PBLKS202000015  Kerugian transaksi spot dan derivatif (realised)
$query_add= " and c.PBLKS_Level_3='PBLKS202000015' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c47=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d47=$row2['Nominal'];
//PBLKS202000016  Kerugian penurunan nilai aset keuangan (impairment)
$query_add= " and c.PBLKS_Level_3='PBLKS202000016' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c48=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d48=$row2['Nominal'];
//PBLKS202000017        Surat berharga
$query_add= " and c.PBLKS_Level_3='PBLKS202000017' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c49=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d49=$row2['Nominal'];
//PBLKS202000018        Kredit
$query_add= " and c.PBLKS_Level_3='PBLKS202000018' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c50=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d50=$row2['Nominal'];
//PBLKS202000019        Pembiayaan syariah
$query_add= " and c.PBLKS_Level_3='PBLKS202000019' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c51=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d51=$row2['Nominal'];
//PBLKS202000020        Aset keuangan lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS202000020' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c52=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d52=$row2['Nominal'];
//PBLKS202000021  Kerugian terkait risiko operasional *)
$query_add= " and c.PBLKS_Level_3='PBLKS202000021' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c53=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d53=$row2['Nominal'];
//PBLKS202000022  Kerugian dari penyertaan dengan equity method
$query_add= " and c.PBLKS_Level_3='PBLKS202000022' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c54=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d54=$row2['Nominal'];
//PBLKS202000023  Komisi/provisi/fee dan administrasi  
$query_add= " and c.PBLKS_Level_3='PBLKS202000023' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c55=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d55=$row2['Nominal'];
//PBLKS202000024  Kerugian penurunan nilai aset lainnya (non keuangan)
$query_add= " and c.PBLKS_Level_3='PBLKS202000024' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c56=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d56=$row2['Nominal'];
//PBLKS202000025  Beban tenaga kerja
$query_add= " and c.PBLKS_Level_3='PBLKS202000025' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c57=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d57=$row2['Nominal'];
//PBLKS202000026  Beban promosi
$query_add= " and c.PBLKS_Level_3='PBLKS202000026' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c58=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d58=$row2['Nominal'];
//PBLKS202000027  Beban lainnya
$query_add= " and c.PBLKS_Level_3='PBLKS202000027' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c59=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d59=$row2['Nominal'];
//PBLKS202000028  Pendapatan (Beban) Operasional Selain Bunga Bersih
$query_add= " and c.PBLKS_Level_3='PBLKS202000028' ";
$result=odbc_exec($connection2, $query.$query_add);
$row=odbc_fetch_array($result);
$c60=$row['Nominal'];
$result2=odbc_exec($connection2, $query2.$query_add);
$row2=odbc_fetch_array($result2);
$d60=$row2['Nominal'];



/*

PBLKS203000002  Keuntungan (kerugian) penjualan aset tetap dan inventaris
PBLKS203000003  Keuntungan (kerugian) penjabaran transaksi valuta asing
PBLKS203000004  Pendapatan (beban) non operasional lainnya

PBLKS204000000  Pajak penghasilan
PBLKS204000001       Taksiran pajak tahun berjalan
PBLKS204000002       Pendapatan (beban) pajak tangguhan
*/





$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN LABA RUGI KOMPREHENSIF");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per $var_tgl dan $prev_tgl ");


$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "POS - POS");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "$var_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "$prev_tgl");

$objPHPExcel->getActiveSheet()->setCellValue('B9', "PENDAPATAN DAN BEBAN OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('A10', "A");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "Pendapatan dan Beban Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A11', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Pendapatan Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Beban Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Pendapatan (Beban) Bunga Bersih");

$objPHPExcel->getActiveSheet()->setCellValue('A18', "B.");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Pendapatan dan Beban Operasional selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A19', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Pendapatan Operasional Selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "a. Peningkatan nilai wajar");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "iii. Spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "iv . Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "b. Penurunan nilai wajar liabilitas keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "c. Keuntungan penjualan aset keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "iii. Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "d. Keuntungan transaksi spot dan derivatif (realised)");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "e. Deviden ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "f. Keuntungan dari Penyertaan dengan equity Method");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "g. Komisi/provisi/fee dan administrasi");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "h. Pemulihan atas cadangan kerugian penurunan nilai");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "i. Pendapatan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A36', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "Beban Operasional Selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "a. Penurunan nilai wajar aset keuangan (mark to market) ");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "iii. Spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B41', "iv . Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "b. Peningkatan nilai wajar liabilitas keuangan (mart to market)");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "c. Kerugian penjualan aset keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "iii. Aset keuangan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B47', "d. Kerugian transaksi spot dan derivatif (realised)");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "e. Kerugian penurunan nilai aset keuangan (impairment)");
$objPHPExcel->getActiveSheet()->setCellValue('B49', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "iii. Pembiayaan syariah");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "iv . Aset keuangan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B53', "f. Kerugian terkait risiko operasional *)");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "g. Kerugian dari penyertaan dengan equity method");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "h. Komisi/provisi/fee dan administrasi");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "i. Kerugian penurunan nilai aset lainnya (non keuangan)");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "j. Beban tenaga kerja");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "k. Beban promosi");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "l. Beban lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B60', "Pendapatan (Beban) Operasional Selain Bunga Bersih");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "LABA (RUGI) OPERASIONAL");

$objPHPExcel->getActiveSheet()->setCellValue('A63', "PENDAPATAN DAN BEBAN NON OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('A64', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "Keuntungan (kerugian) penjualan aset tetap dan inventaris");
$objPHPExcel->getActiveSheet()->setCellValue('A65', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "Keuntungan (kerugian) penjabaran transaksi valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('A66', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B66', "Pendapatan (beban) non operasional lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "LABA (RUGI) OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "LABA (RUGI) TAHUN BERJALAN");
$objPHPExcel->getActiveSheet()->setCellValue('B71', "Pajak Penghasilan");
$objPHPExcel->getActiveSheet()->setCellValue('B72', "a. Taksiran Pajak Tahun Berjalan");
$objPHPExcel->getActiveSheet()->setCellValue('B73', "b. Pendapatan (beban) pajak tangguhan");
$objPHPExcel->getActiveSheet()->setCellValue('B75', "LABA (RUGI) BERSIH");

$objPHPExcel->getActiveSheet()->setCellValue('A77', "PENGHASILAN KOMPREHENSIF LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('A78', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Pos-pos yang tidak akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "a. Keuntungan revaluasi aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('B80', "b. Keuntungan (kerugian) aktuarial program imbalan pasti");
$objPHPExcel->getActiveSheet()->setCellValue('B81', "c. Bagian pendapatan komprehensif lain dari entitas asosiasi");
$objPHPExcel->getActiveSheet()->setCellValue('B82', "d. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B83', "e. Pajak penghasilan terkait pos-pos yang tidak akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('A84', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "Pos-pos yang akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('B85', "a. Penyesuaian akibat penjabaran laporan keuangan dalam mata uang asing");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "c. Bagian efektif dari lindung nilai arus kas");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "d. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "e. Pajak penghasilan terkait pos-pos yang akan direklasifikasi ke laba rugi");

$objPHPExcel->getActiveSheet()->setCellValue('B90', "PENGHASILAN KOMPREHENSIF LAIN TAHUN BERJALAN - NET PAJAK PENGHASILAN TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('B91', "TOTAL LABA (RUGI) KOMPREHENSIF TAHUN BERJALAN");
$objPHPExcel->getActiveSheet()->setCellValue('B92', "Laba yang dapat diatribusikan kepada :");
$objPHPExcel->getActiveSheet()->setCellValue('B93', "PEMILIK");
$objPHPExcel->getActiveSheet()->setCellValue('B94', "KEPENTINGAN NNON PENGENDALI");
$objPHPExcel->getActiveSheet()->setCellValue('B95', "TOTAL LABA TAHUN BERJALAN ");

$objPHPExcel->getActiveSheet()->setCellValue('B96', "Total Penghasilan Komprehensif Lain yang dapat diatribusikan kepada :");
$objPHPExcel->getActiveSheet()->setCellValue('B97', "PEMILIK");
$objPHPExcel->getActiveSheet()->setCellValue('B98', "KEPENTINGAN NNON PENGENDALI");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "TOTAL PENGHASILAN KOMPREHENSIF LAIN TAHUN BERJALAN ");

$objPHPExcel->getActiveSheet()->setCellValue('B101', "TRANSFER LABA (RUGI) KE KANTOR PUSAT ");
$objPHPExcel->getActiveSheet()->setCellValue('B102', "DIVIDEN ");
$objPHPExcel->getActiveSheet()->setCellValue('B103', "LABA (RUGI) BERSIH PER SAHAM");




$objPHPExcel->getActiveSheet()->setCellValue('C12', $c12);
$objPHPExcel->getActiveSheet()->setCellValue('D12', $d12);
$objPHPExcel->getActiveSheet()->setCellValue('C13', $c13);
$objPHPExcel->getActiveSheet()->setCellValue('D13', $d13);

$objPHPExcel->getActiveSheet()->setCellValue('C15', $c15);
$objPHPExcel->getActiveSheet()->setCellValue('D15', $d15);
$objPHPExcel->getActiveSheet()->setCellValue('C16', $c16);
$objPHPExcel->getActiveSheet()->setCellValue('D16', $d16);


$objPHPExcel->getActiveSheet()->setCellValue('C20', $c15);
$objPHPExcel->getActiveSheet()->setCellValue('D20', $d15);
$objPHPExcel->getActiveSheet()->setCellValue('C21', $c21);
$objPHPExcel->getActiveSheet()->setCellValue('D21', $d21);
$objPHPExcel->getActiveSheet()->setCellValue('C22', $c22);
$objPHPExcel->getActiveSheet()->setCellValue('D22', $d22);
$objPHPExcel->getActiveSheet()->setCellValue('C23', $c23);
$objPHPExcel->getActiveSheet()->setCellValue('D23', $d23);
$objPHPExcel->getActiveSheet()->setCellValue('C24', $c24);
$objPHPExcel->getActiveSheet()->setCellValue('D24', $d24);
$objPHPExcel->getActiveSheet()->setCellValue('C25', $c25);
$objPHPExcel->getActiveSheet()->setCellValue('D25', $d25);
$objPHPExcel->getActiveSheet()->setCellValue('C26', $c26);
$objPHPExcel->getActiveSheet()->setCellValue('D26', $d26);
$objPHPExcel->getActiveSheet()->setCellValue('C27', $c27);
$objPHPExcel->getActiveSheet()->setCellValue('D27', $d27);
$objPHPExcel->getActiveSheet()->setCellValue('C28', $c28);
$objPHPExcel->getActiveSheet()->setCellValue('D28', $d28);
$objPHPExcel->getActiveSheet()->setCellValue('C29', $c29);
$objPHPExcel->getActiveSheet()->setCellValue('D29', $d29);
$objPHPExcel->getActiveSheet()->setCellValue('C30', $c30);
$objPHPExcel->getActiveSheet()->setCellValue('D30', $d30);
$objPHPExcel->getActiveSheet()->setCellValue('C31', $c31);
$objPHPExcel->getActiveSheet()->setCellValue('D31', $d31);
$objPHPExcel->getActiveSheet()->setCellValue('C32', $c32);
$objPHPExcel->getActiveSheet()->setCellValue('D32', $d32);
$objPHPExcel->getActiveSheet()->setCellValue('C33', $c33);
$objPHPExcel->getActiveSheet()->setCellValue('D33', $d33);
$objPHPExcel->getActiveSheet()->setCellValue('C34', $c34);
$objPHPExcel->getActiveSheet()->setCellValue('D34', $d34);
$objPHPExcel->getActiveSheet()->setCellValue('C35', $c35);
$objPHPExcel->getActiveSheet()->setCellValue('D35', $d35);

$objPHPExcel->getActiveSheet()->setCellValue('C37', $c37);
$objPHPExcel->getActiveSheet()->setCellValue('D37', $d37);
$objPHPExcel->getActiveSheet()->setCellValue('C38', $c38);
$objPHPExcel->getActiveSheet()->setCellValue('D38', $d38);
$objPHPExcel->getActiveSheet()->setCellValue('C39', $c39);
$objPHPExcel->getActiveSheet()->setCellValue('D39', $d39);
$objPHPExcel->getActiveSheet()->setCellValue('C40', $c40);
$objPHPExcel->getActiveSheet()->setCellValue('D40', $d40);
$objPHPExcel->getActiveSheet()->setCellValue('C41', $c41);
$objPHPExcel->getActiveSheet()->setCellValue('D41', $d41);
$objPHPExcel->getActiveSheet()->setCellValue('C42', $c42);
$objPHPExcel->getActiveSheet()->setCellValue('D42', $d42);
$objPHPExcel->getActiveSheet()->setCellValue('C43', $c43);
$objPHPExcel->getActiveSheet()->setCellValue('D43', $d43);
$objPHPExcel->getActiveSheet()->setCellValue('C44', $c44);
$objPHPExcel->getActiveSheet()->setCellValue('D44', $d44);
$objPHPExcel->getActiveSheet()->setCellValue('C45', $c45);
$objPHPExcel->getActiveSheet()->setCellValue('D45', $d45);
$objPHPExcel->getActiveSheet()->setCellValue('C46', $c46);
$objPHPExcel->getActiveSheet()->setCellValue('D46', $d46);
$objPHPExcel->getActiveSheet()->setCellValue('C47', $c47);
$objPHPExcel->getActiveSheet()->setCellValue('D47', $d47);
$objPHPExcel->getActiveSheet()->setCellValue('C48', $c48);
$objPHPExcel->getActiveSheet()->setCellValue('D48', $d48);
$objPHPExcel->getActiveSheet()->setCellValue('C49', $c49);
$objPHPExcel->getActiveSheet()->setCellValue('D49', $d49);
$objPHPExcel->getActiveSheet()->setCellValue('C50', $c50);
$objPHPExcel->getActiveSheet()->setCellValue('D50', $d50);
$objPHPExcel->getActiveSheet()->setCellValue('C51', $c51);
$objPHPExcel->getActiveSheet()->setCellValue('D51', $d51);
$objPHPExcel->getActiveSheet()->setCellValue('C52', $c52);
$objPHPExcel->getActiveSheet()->setCellValue('D52', $d52);
$objPHPExcel->getActiveSheet()->setCellValue('C53', $c53);
$objPHPExcel->getActiveSheet()->setCellValue('D53', $d53);
$objPHPExcel->getActiveSheet()->setCellValue('C54', $c54);
$objPHPExcel->getActiveSheet()->setCellValue('D54', $d54);
$objPHPExcel->getActiveSheet()->setCellValue('C55', $c55);
$objPHPExcel->getActiveSheet()->setCellValue('D55', $d55);
$objPHPExcel->getActiveSheet()->setCellValue('C56', $c56);
$objPHPExcel->getActiveSheet()->setCellValue('D56', $d56);
$objPHPExcel->getActiveSheet()->setCellValue('C57', $c57);
$objPHPExcel->getActiveSheet()->setCellValue('D57', $d57);
$objPHPExcel->getActiveSheet()->setCellValue('C58', $c58);
$objPHPExcel->getActiveSheet()->setCellValue('D58', $d58);
$objPHPExcel->getActiveSheet()->setCellValue('C59', $c59);
$objPHPExcel->getActiveSheet()->setCellValue('D59', $d59);
$objPHPExcel->getActiveSheet()->setCellValue('C60', $c60);
$objPHPExcel->getActiveSheet()->setCellValue('D60', $d60);



$objPHPExcel->getActiveSheet()->getStyle('C10:D103')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');



$objPHPExcel->getActiveSheet()->setTitle('PL');

###################################################################################3
//SHEET 3 (Rek Adm)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);


$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(90);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C7:D7');

$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A50:B50');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D54')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN KOMITMEN & KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015 dan 31 Desember 2014 ");



$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "POS - POS");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('A9', "I");
$objPHPExcel->getActiveSheet()->setCellValue('A9', "TAGIHAN KOMITMEN");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "1. Fasilitas pinjaman yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "2. Posisi pembelian spot dan derivatif yang masih berjalan  ");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "3. Lainnya ");
    
$objPHPExcel->getActiveSheet()->setCellValue('A15', "II");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "KEWAJIBAN KOMITMEN");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "1. Fasilitas kredit kepada nasabah yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "a. BUMN ");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "i  . Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "- Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "ii . Uncommitted ");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "- Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "b. Lainnya ");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "i  . Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "ii . Uncommitted");

$objPHPExcel->getActiveSheet()->setCellValue('B27', "2. Fasilitas kredit kepada bank lain yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "a. Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "i  . Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "ii . Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "b . Uncommitted ");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "i  . Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "ii . Valuta Asing ");

$objPHPExcel->getActiveSheet()->setCellValue('B35', "3. Irrevocable L/C yang masih berjalan ");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "a . L/C luar negeri ");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "b . L/C dalam negeri");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "4 . Posisi penjualan spot dan derivatif yang masih berjalan  ");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "5 . Lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A41', "III");
$objPHPExcel->getActiveSheet()->setCellValue('A41', "TAGIHAN KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "1. Garansi yang diterima  ");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "2. Pendapatan bunga dalam penyelesaian");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "a. Bunga Kredit yang diberikan ");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "b. Bunga Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "3. Lainnya ");


$objPHPExcel->getActiveSheet()->setCellValue('A50', "IV KEWAJIBAN KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "1. Garansi yang diterima  ");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "2. Lainnya ");



$objPHPExcel->getActiveSheet()->setTitle('Rek Adm');





//SHEET 4 (Rasio)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3);

$objPHPExcel->getActiveSheet()->setTitle('Rasio');





$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(103);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('C7:D7');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D32')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A1', "Rasio");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015 dan 2014 ");


$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Rasio");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('A9', "Rasio Kinerja");
$objPHPExcel->getActiveSheet()->setCellValue('A11', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Kewajiban Penyediaan Modal Minimum (KPMM) ");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Aset produktif bermasalah dan aset non produktif bermasalah terhadap total aset produktif dan aset non produktif");
$objPHPExcel->getActiveSheet()->setCellValue('A13', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "Aset produktif bermasalah terhadap total aset produktif ");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Cadangan kerugian penurunan nilai (CKPN) aset keuangan  terhadap aset produktif ");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "NPL gross ");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "NPL net ");
$objPHPExcel->getActiveSheet()->setCellValue('A17', "7");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Return on Asset (ROA) ");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "8");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Return on Equity (ROE)
 ");
$objPHPExcel->getActiveSheet()->setCellValue('A19', "9");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Net Interest Margin (NIM) ");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "10");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "Biaya Operasional terhadap Pendapatan Operasional (BOPO) ");
$objPHPExcel->getActiveSheet()->setCellValue('A21', "11");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "Loan to Deposit Ratio (LDR) ");

$objPHPExcel->getActiveSheet()->setCellValue('A22', "Kepatuhan (Compliance)");
$objPHPExcel->getActiveSheet()->setCellValue('A23', "1. ");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "a.   Persentase pelanggaran BMPK");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "i .  Pihak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "ii.  Pihak tidak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "b.   Persentase pelampauan BMPK");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "i .  Pihak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "ii.  Pihak tidak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('A29', "2. ");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "Giro Wajib Minimum (GWM) ");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "a.   GWM Utama Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "b.   GWM Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('A32', "3. ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "Posisi Devisa Neto (PDN) secara keseluruhan ");



####################################################################################
//SHEET 5 (Derivatif)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(4);



$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A1:G1');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A2:G2');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A3:G3');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A7:A9');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B7:B9');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('C7:G7');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('C8:C9');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('D8:E8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('F8:G8');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B11:G11');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B26:G26');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B39:G39');



$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:A40')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A7:A9')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B7:G40')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A11:G11')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A26:G26')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A39:G39')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A40:G40')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN TRANSAKSI SPOT DAN DERIVATIF");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015");

$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Transaksi");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "Nilai Notional");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "Tujuan");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "Tagihan dan Liabilitas Derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('D9', "Trading");
$objPHPExcel->getActiveSheet()->setCellValue('E9', "Hedging");
$objPHPExcel->getActiveSheet()->setCellValue('F9', "Tagihan");
$objPHPExcel->getActiveSheet()->setCellValue('G9', "Liabilitas");

$objPHPExcel->getActiveSheet()->setCellValue('A11', "A.");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Terkait dengan Nilai Tukar");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Spot");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Forward");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "Option");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "a. Jual");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "b. Beli");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "Future");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "Swap");
$objPHPExcel->getActiveSheet()->setCellValue('A24', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "Lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A26', "B.");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "Terkait dengan Suku Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A27', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "Forward");
$objPHPExcel->getActiveSheet()->setCellValue('A29', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "Option");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "a. Jual");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "b. Beli");
$objPHPExcel->getActiveSheet()->setCellValue('A33', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "Future");
$objPHPExcel->getActiveSheet()->setCellValue('A35', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "Swap");
$objPHPExcel->getActiveSheet()->setCellValue('A37', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('A39', "C.");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "JUMLAH");

$objPHPExcel->getActiveSheet()->setTitle('Derivatif');


//SHEET 6 (KA dan CKPN)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(5);
// 

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(75);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(15);

$objPHPExcel->setActiveSheetIndex(5)->mergeCells('A1:N1');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('A2:N2');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('A3:N3');

$objPHPExcel->setActiveSheetIndex(5)->mergeCells('A7:A9');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('B7:B9');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('C7:N7');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('C8:H8');
$objPHPExcel->setActiveSheetIndex(5)->mergeCells('I8:N8');

$objPHPExcel->getActiveSheet()->getStyle('A1:N3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:N9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:N9')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:N9')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A1:A3')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A7:N100')->applyFromArray($styleArrayBorder1);



$objPHPExcel->getActiveSheet()->setCellValue('A1', "KUALITAS ASET PRODUKTIF DAN INFORMASI LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "30 September 2015 dan 2014");

$objPHPExcel->getActiveSheet()->setCellValue('A7', "NO");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "POS-POS");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sept-2016");
$objPHPExcel->getActiveSheet()->setCellValue('I8', "30-Sept-2016");

$objPHPExcel->getActiveSheet()->setCellValue('C9', "L");
$objPHPExcel->getActiveSheet()->setCellValue('D9', "DPK");
$objPHPExcel->getActiveSheet()->setCellValue('E9', "KL");
$objPHPExcel->getActiveSheet()->setCellValue('F9', "D");
$objPHPExcel->getActiveSheet()->setCellValue('G9', "M");
$objPHPExcel->getActiveSheet()->setCellValue('H9', "Jumlah");
$objPHPExcel->getActiveSheet()->setCellValue('I9', "L");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "DPK");
$objPHPExcel->getActiveSheet()->setCellValue('K9', "KL");
$objPHPExcel->getActiveSheet()->setCellValue('L9', "D");
$objPHPExcel->getActiveSheet()->setCellValue('M9', "M");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "Jumlah");

$objPHPExcel->getActiveSheet()->setCellValue('A11', "I.");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "PIHAK TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Penempatan pada Bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "Tagihan spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Surat berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A21', "4.");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "Surat Berharga yang dijual dengan janji dibeli kembali (Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A24', "5.");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "Tagihan atas surat berharga yang dibeli dengan janji dijual kembali ");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A27', "6.");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "Tagihan Akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A30', "7.");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "a. Debitur Usaha Mikro, Kecil dan Menengah (UMKM)");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "b. Bukan debitur UMKM ");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "c. Kredit yang direstrukturisasi");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "d. Kredit properti ");
$objPHPExcel->getActiveSheet()->setCellValue('A41', "8.");
$objPHPExcel->getActiveSheet()->setCellValue('B41', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('A42', "9.");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('A43', "10.");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('A44', "11.");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A47', "12.");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "Kredit");


$objPHPExcel->getActiveSheet()->setCellValue('A48', "II.");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "PIHAK TIDAK TERKAIT ");
$objPHPExcel->getActiveSheet()->setCellValue('A49', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B49', "Penempatan pada bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A52', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "Tagihan spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A55', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "Surat berharga ");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A58', "4.");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "Surat Berharga yang dijual dengan janji dibeli kembali (Repo)");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B60', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A61', "5.");
$objPHPExcel->getActiveSheet()->setCellValue('B61', "Tagihan atas surat berharga yang dibeli dengan janji dijual kembali ");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B63', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A64', "6.");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "Tagihan Akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('A65', "7.");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B67', "a. Debitur Usaha Mikro, Kecil dan Menengah (UMKM)");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B69', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "b. Bukan debitur UMKM ");
$objPHPExcel->getActiveSheet()->setCellValue('B71', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B72', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B73', "c. Kredit yang direstrukturisasi");
$objPHPExcel->getActiveSheet()->setCellValue('B74', "i.  Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B75', "ii. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B76', "d. Kredit properti ");
$objPHPExcel->getActiveSheet()->setCellValue('A77', "8.");
$objPHPExcel->getActiveSheet()->setCellValue('B77', "Penyertaan");
$objPHPExcel->getActiveSheet()->setCellValue('A78', "9.");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Penyertaan modal sementara ");
$objPHPExcel->getActiveSheet()->setCellValue('A79', "10.");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "Tagihan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B80', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B81', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A82', "11.");
$objPHPExcel->getActiveSheet()->setCellValue('B82', "Komitmen dan Kotinjensi");
$objPHPExcel->getActiveSheet()->setCellValue('B83', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('A85', "12.");
$objPHPExcel->getActiveSheet()->setCellValue('B85', "Aset yang diambil alih");

$objPHPExcel->getActiveSheet()->setCellValue('A86', "III.");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "INFORMASI LAIN ");
$objPHPExcel->getActiveSheet()->setCellValue('A87', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "Total aset bank yang dijaminkan : ");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "a. Pada Bank Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "b. Pada Pihak Lain");
$objPHPExcel->getActiveSheet()->setCellValue('A90', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B90', "Total CKPN aset keuangan atas aset produktif");
$objPHPExcel->getActiveSheet()->setCellValue('A91', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B91', "Total PPA yang wajib dibentuk atas aset produktif");
$objPHPExcel->getActiveSheet()->setCellValue('A92', "4.");
$objPHPExcel->getActiveSheet()->setCellValue('B92', "Persentase kredit kepada UMKM  terhadap total kredit");
$objPHPExcel->getActiveSheet()->setCellValue('A93', "5.");
$objPHPExcel->getActiveSheet()->setCellValue('B93', "Persentase kredit kepada Usaha Mikro Kecil (UMK) terhadap total kredit ");
$objPHPExcel->getActiveSheet()->setCellValue('A94', "6.");
$objPHPExcel->getActiveSheet()->setCellValue('B94', "Persentase jumlah debitur UMKM terhadap total debitur");
$objPHPExcel->getActiveSheet()->setCellValue('A95', "7.");
$objPHPExcel->getActiveSheet()->setCellValue('B95', "Persentase jumlah debitur Usaha Mikro Kecil (UMK) terhadap total debitur ");
$objPHPExcel->getActiveSheet()->setCellValue('A96', "8.");
$objPHPExcel->getActiveSheet()->setCellValue('B96', "a. Penerusan kredit ");
$objPHPExcel->getActiveSheet()->setCellValue('B97', "b. Penyaluran dana Mudharabah Muqayyadah ");
$objPHPExcel->getActiveSheet()->setCellValue('B98', "c. Aset produktif yang dihapus buku");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "d. Aset produktif dihapusbuku yg dipulihkan/berhasil ditagih");
$objPHPExcel->getActiveSheet()->setCellValue('B100', "e. Aset produktif yang dihapus tagih");
     

$objPHPExcel->getActiveSheet()->setTitle('KA dan CKPN');


//SHEET 7 (KPMM)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(6);
// 

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(6)->mergeCells('A1:M1');
$objPHPExcel->setActiveSheetIndex(6)->mergeCells('A2:M2');
$objPHPExcel->setActiveSheetIndex(6)->mergeCells('A3:M3');

$objPHPExcel->setActiveSheetIndex(6)->mergeCells('A7:K8');
$objPHPExcel->setActiveSheetIndex(6)->mergeCells('L7:L8');
$objPHPExcel->setActiveSheetIndex(6)->mergeCells('M7:M8');

$objPHPExcel->setActiveSheetIndex(6)->mergeCells('A50:K50');


for ($i=52; $i <= 58 ; $i++) { 
    $objPHPExcel->setActiveSheetIndex(6)->mergeCells("A$i:E$i");
    $objPHPExcel->setActiveSheetIndex(6)->mergeCells("F$i:G$i");
    $objPHPExcel->setActiveSheetIndex(6)->mergeCells("H$i:I$i");
    $objPHPExcel->setActiveSheetIndex(6)->mergeCells("J$i:K$i");
}


$objPHPExcel->getActiveSheet()->getStyle('A1:M3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:M7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:M8')->applyFromArray($styleArrayBorder1);
//$objPHPExcel->getActiveSheet()->getStyle('A7:A40')->applyFromArray($styleArrayBorder2);
//$objPHPExcel->getActiveSheet()->getStyle('A7:A9')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A52:M58')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B9:K49')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('L9:M49')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('A9:B49')->applyFromArray($styleArrayBorder2);

$objPHPExcel->getActiveSheet()->getStyle('A9:M9')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A41:M41')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A42:M46')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A49:M49')->applyFromArray($styleArrayBorder2);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN PERHITUNGAN KEWAJIBAN PENYEDIAAN MODAL MINIMUM");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015");



$objPHPExcel->getActiveSheet()->setCellValue('L7', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('M7', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('A9', "I");
$objPHPExcel->getActiveSheet()->setCellValue('B9', "Modal Inti (Tier 1)");

$objPHPExcel->getActiveSheet()->setCellValue('B10', "1");
$objPHPExcel->getActiveSheet()->setCellValue('C10', "Modal Inti Utama (CET 1)");
$objPHPExcel->getActiveSheet()->setCellValue('C11', "1.1  Modal disetor (Setelah dikurangi Saham Treasury)");
$objPHPExcel->getActiveSheet()->setCellValue('C12', "1.2  Cadangan Tambahan Modal 1)");
$objPHPExcel->getActiveSheet()->setCellValue('C13', "1.2.1   Agio / Disagio");
$objPHPExcel->getActiveSheet()->setCellValue('C14', "1.2.2   Modal sumbangan");
$objPHPExcel->getActiveSheet()->setCellValue('C15', "1.2.3   Cadangan umum");
$objPHPExcel->getActiveSheet()->setCellValue('C16', "1.2.4   Laba/Rugi tahun-tahun lalu yang dapat diperhitungkan");
$objPHPExcel->getActiveSheet()->setCellValue('C17', "1.2.5   Laba/Rugi tahun berjalan yang dapat diperhitungkan");
$objPHPExcel->getActiveSheet()->setCellValue('C18', "1.2.6   Selisih lebih karena penjabaran laporan keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('C19', "1.2.7   Dana setoran modal ");
$objPHPExcel->getActiveSheet()->setCellValue('C20', "1.2.8   Waran yang diterbitkan");
$objPHPExcel->getActiveSheet()->setCellValue('C21', "1.2.9   Opsi saham yang diterbitkan dalam rangka program kompensasi berbasis saham");
$objPHPExcel->getActiveSheet()->setCellValue('C22', "1.2.10  Pendapatan komprehensif lain");
$objPHPExcel->getActiveSheet()->setCellValue('C23', "1.2.11  Saldo surplus revaluasi aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('C24', "1.2.12  Selisih kurang antara PPA dan cadangan kerugian penurunan nilai atas aset produktif");
$objPHPExcel->getActiveSheet()->setCellValue('C25', "1.2.13  Penyisihan Penghapusan Aset (PPA) atas aset non produktif yang wajib dihitung");
$objPHPExcel->getActiveSheet()->setCellValue('C26', "1.2.14  Selisih kurang jumlah penyesuaian nilai wajar dari instrumen keuangan dalam trading book");
$objPHPExcel->getActiveSheet()->setCellValue('C27', "1.3  Kepentingan Non Pengendali yang dapat diperhitungkan)");
$objPHPExcel->getActiveSheet()->setCellValue('C28', "1.4  Faktor Pengurang Modal Inti Utama 1)");
$objPHPExcel->getActiveSheet()->setCellValue('C29', "1.4.1   Perhitungan pajak tangguhan");
$objPHPExcel->getActiveSheet()->setCellValue('C30', "1.4.2   Goodwill");
$objPHPExcel->getActiveSheet()->setCellValue('C31', "1.4.3   Aset tidak berwujud lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('C32', "1.4.4   Penyertaan yang diperhitungkan sebagai faktor pengurang");
$objPHPExcel->getActiveSheet()->setCellValue('C33', "1.4.5   Kekurangan modal pada perusahaan anak asuransi");
$objPHPExcel->getActiveSheet()->setCellValue('C34', "1.4.6   Eksposur sekuritisasi");
$objPHPExcel->getActiveSheet()->setCellValue('C35', "1.4.7   Faktor Pengurang modal inti lainnya ");
$objPHPExcel->getActiveSheet()->setCellValue('C36', "1.4.8   Investasi pada instrumen AT1 dan Tier 2 pada bank lain 2");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "2");
$objPHPExcel->getActiveSheet()->setCellValue('C37', "Modal Inti Tambahan (AT-1)  1)");
$objPHPExcel->getActiveSheet()->setCellValue('C38', "2.1  Instrumen yang memenuhi persyaratan AT-1");
$objPHPExcel->getActiveSheet()->setCellValue('C39', "2.2  Agio / Disagio");
$objPHPExcel->getActiveSheet()->setCellValue('C40', "2.3  Faktor Pengurang: Investasi pada instrumen AT1 dan Tier 2 pada bank lain 2)");                        

$objPHPExcel->getActiveSheet()->setCellValue('A41', "II");  
$objPHPExcel->getActiveSheet()->setCellValue('B41', "Modal Pelengkap (Tier 2)");  
$objPHPExcel->getActiveSheet()->setCellValue('B42', "1   Instrumen modal dalam bentuk saham atau lainnya yang memenuhi persyaratan ");  
$objPHPExcel->getActiveSheet()->setCellValue('B43', "2   Agio / disagio yang berasal dari penerbitan instrumen modal pelengkap");  
$objPHPExcel->getActiveSheet()->setCellValue('B44', "3   Cadangan umum aset produktif PPA yang wajib dibentuk (maks 1,25% ATMR Risiko Kredit)");  
$objPHPExcel->getActiveSheet()->setCellValue('B45', "4   Cadangan tujuan        ");  
$objPHPExcel->getActiveSheet()->setCellValue('B46', "5   Faktor Pengurang Modal Pelengkap 1)  ");  
$objPHPExcel->getActiveSheet()->setCellValue('C47', "5.2  Investasi pada instrumen Tier 2 pada bank lain 2)");  
              
$objPHPExcel->getActiveSheet()->setCellValue('A49', "Total Modal");             
                
$objPHPExcel->getActiveSheet()->setCellValue('A52', "");  
$objPHPExcel->getActiveSheet()->setCellValue('F52', "30 September 2015");  
$objPHPExcel->getActiveSheet()->setCellValue('H52', "30 September 2014");            
$objPHPExcel->getActiveSheet()->setCellValue('J52', "KETERANGAN"); 

$objPHPExcel->getActiveSheet()->setCellValue('A53', "ASET TERTIMBANG MENURUT RISIKO"); 
$objPHPExcel->getActiveSheet()->setCellValue('A54', "ATMR RISIKO KREDIT 3)"); 
$objPHPExcel->getActiveSheet()->setCellValue('A55', "ATMR RISIKO PASAR "); 
$objPHPExcel->getActiveSheet()->setCellValue('A56', "ATMR RISIKO OPERASIONAL "); 
$objPHPExcel->getActiveSheet()->setCellValue('A57', "TOTAL ATMR "); 
$objPHPExcel->getActiveSheet()->setCellValue('A58', "RASIO KPMM SESUAI PROFIL RISIKO"); 

$objPHPExcel->getActiveSheet()->setCellValue('J53', "RASIO KPMM"); 
$objPHPExcel->getActiveSheet()->setCellValue('J54', "Rasio CET1"); 
$objPHPExcel->getActiveSheet()->setCellValue('J55', "Rasio Tier 1"); 
$objPHPExcel->getActiveSheet()->setCellValue('J56', "Rasio Tier 2"); 
$objPHPExcel->getActiveSheet()->setCellValue('J57', "Rasio total"); 
$objPHPExcel->getActiveSheet()->setCellValue('J58', "CET 1 UNTUK  BUFFER "); 

  




          
        
         


$objPHPExcel->getActiveSheet()->setTitle('KPMM');

//SHEET 8 (Pengurus dan Pmlk)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(7);
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('A3:D3');

$objPHPExcel->setActiveSheetIndex(7)->mergeCells('A5:B5');
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('C5:D5');
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('C23:D23');
$objPHPExcel->setActiveSheetIndex(7)->mergeCells('C24:D24');

for ($i=7; $i <=19 ; $i++) { 
   $objPHPExcel->setActiveSheetIndex(7)->mergeCells("C$i:D$i");
}
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(70);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(50);

$objPHPExcel->getActiveSheet()->getStyle('A1:M3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A5:D5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C23:D29')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:D5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('C24:D24')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A13')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A5:B30')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('C5:D30')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A5:D5')->applyFromArray($styleArrayBorder2);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "SUSUNAN PENGURUS");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015");

$objPHPExcel->getActiveSheet()->setCellValue('A5', "PENGURUS BANK");
$objPHPExcel->getActiveSheet()->setCellValue('C5', "PEMEGANG SAHAM ");



$objPHPExcel->getActiveSheet()->setCellValue('A7', "DEWAN KOMISARIS");
$objPHPExcel->getActiveSheet()->setCellValue('A8', '- Presiden Komisaris (independen) ');
$objPHPExcel->getActiveSheet()->setCellValue('A9', "- Komisaris");
$objPHPExcel->getActiveSheet()->setCellValue('A10', '- Komisaris Independen');

$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Bambang Ratmanto ');
$objPHPExcel->getActiveSheet()->setCellValue('B9', "Purnadi Harjono");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "Eko B. Supriyanto");



$objPHPExcel->getActiveSheet()->setCellValue('A13', "DIREKSI");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "-  Presiden Direktur");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "- Direktur ");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "- Direktur ");
$objPHPExcel->getActiveSheet()->setCellValue('A17', "- Direktur yang Membawahkan Fungsi Kepatuhan");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "- Direktur Independen");

$objPHPExcel->getActiveSheet()->setCellValue('B13', "");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Benny Purnomo");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "Benny Helman");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "Nerfita Primasari");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Tjit Siat Fun");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Widiatama Bunarto");



$objPHPExcel->getActiveSheet()->setCellValue('C7', "Pemegang Saham Pengendali (PSP) :");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "PT MNC Kapital Indonesia Tbk, dengan kepemilikan saham sebesar 39.88%");
$objPHPExcel->getActiveSheet()->setCellValue('C11', "Pemegang Saham Bukan PSP melalui pasar modal  (≥   5%) : ");
$objPHPExcel->getActiveSheet()->setCellValue('C12', "RBC Singapore - CLIENTS A/C sebesar 12.70% ");
$objPHPExcel->getActiveSheet()->setCellValue('C13', "Citibank Singapore S/A BK Julius Baer & CO LTD CLIENT A/C sebesar 5.38% ");
$objPHPExcel->getActiveSheet()->setCellValue('C14', "       
");
$objPHPExcel->getActiveSheet()->setCellValue('C15', "Pemegang Saham Bukan PSP melalui pasar modal  (<   5%) : Masyarakat sebesar 41,97% ");
$objPHPExcel->getActiveSheet()->setCellValue('C18', "Pemegang Saham Bukan PSP tidak melalui pasar modal  : ");
$objPHPExcel->getActiveSheet()->setCellValue('C19', "AJB Bumiputera 1912 sebesar 0.07% ");

       
       
        
        
       
     
    
      
        
        
      
      






$objPHPExcel->getActiveSheet()->setCellValue('C23', "Jakarta, 20 Oktober 2015");
$objPHPExcel->getActiveSheet()->setCellValue('C24', "PT Bank MNC Internasional Tbk");

$objPHPExcel->getActiveSheet()->setCellValue('C28', "Benny Purnomo");
$objPHPExcel->getActiveSheet()->setCellValue('C29', "Presiden Direktur");
$objPHPExcel->getActiveSheet()->setCellValue('D28', "Benny Helman");
$objPHPExcel->getActiveSheet()->setCellValue('D29', "Direktur");



$objPHPExcel->getActiveSheet()->setTitle('Pengurus dan Pmlk');







//SHEET 9 (Arus Kas)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(8);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(100);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);

$objPHPExcel->setActiveSheetIndex(8)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(8)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(8)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(8)->mergeCells('A4:D4');

$objPHPExcel->getActiveSheet()->getStyle('A1:D4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:D4')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B8:D60')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN ARUS KAS");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Untuk tahun yang berakhir pada tanggal 30 Sept 2015 dan 2014");
$objPHPExcel->getActiveSheet()->setCellValue('A4', "(Disajikan Dalam Jutaan Rupiah, Kecuali Dinyatakan Lain)");

$objPHPExcel->getActiveSheet()->setCellValue('C7', "30 Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D7', "30 Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('B8', "ARUS KAS DARI AKTIVITAS OPERASI");
$objPHPExcel->getActiveSheet()->setCellValue('B9', "Penerimaan bunga, provisi dan komisi kredit yang diberikan ");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "Pembayaran bunga, hadiah, provisi dan komisi dana ");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Penerimaan pendapatan operasional lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Pembayaran gaji dan tunjangan karyawan");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "Pembayaran beban operasional lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Penerimaan pendapatan beban non-operasional ");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "Pembayaran beban non-operasional ");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "Arus kas operasi sebelum perubahan dalam aset dan kewajiban ");
  
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Penurunan (kenaikan) dalam aset operasi");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Penempatan pada bank indonesia dan bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "Efek-efek yang diperdagangkan");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "Kredit yang diberikan");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "Tagihan Derevatif");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "Tagihan Akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "Agunan yang diambil alih");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "Aset lain-lain");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "Kenaikan (penurunan) Kewajiban Operasi");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "Liabilitas segera");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "Simpanan");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "Simpanan dari bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "Efek yang dijual dengan janji dibeli kembali");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "Liabilitas derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "Liabilitas Akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "Liabilitas lain-lain");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "Kas Bersih yang diperoleh dari (Digunakan untuk) Aktivitas Operasional");


$objPHPExcel->getActiveSheet()->setCellValue('B36', "ARUS KAS DARI AKTIVITAS INVESTASI");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "Pencairan (Perolehan) dari investasi keuangan ");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "Pencairan dari investasi keuangan ");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "Hasil penjualan aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "Perolehan Aset Tetap dan perangkat lunak");
$objPHPExcel->getActiveSheet()->setCellValue('B41', "Kas Bersih yang diperoleh dari (Digunakan untuk) Aktivitas Investasi   ");

$objPHPExcel->getActiveSheet()->setCellValue('B43', "ARUS KAS DARI AKTIVITAS PENDANAAN ");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "Penambahan Modal Saham  ");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "Penerbitan Obligasi Wajib Konversi");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "Biaya Emisi Ekuitas  ");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "Penambahan Dana Cadangan Modal  ");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "Pembayaran Deviden ");       
$objPHPExcel->getActiveSheet()->setCellValue('B49', "Pembayaran Pinjaman yang Diterima ");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "Kas Bersih yang diperoleh dari (Digunakan untuk) Aktivitas Pendanaan");                           
                                                                                                      
$objPHPExcel->getActiveSheet()->setCellValue('B52', "Kenaikan (Penurunan) Bersih Kas dan Setara Kas");  
$objPHPExcel->getActiveSheet()->setCellValue('B53', "Saldo Kas dan Setara Kas Pada Awal Tahun");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "Saldo Kas dan Setara Kas pada Akhir Tahun"); 

$objPHPExcel->getActiveSheet()->setCellValue('B56', "Kas dan Setara Kas Terdiri Dari :");       
$objPHPExcel->getActiveSheet()->setCellValue('B57', "Kas  ");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "Penempatan pada Bank Indonesia"); 
$objPHPExcel->getActiveSheet()->setCellValue('B59', "Penempatan pada bank lain ");       
$objPHPExcel->getActiveSheet()->setCellValue('B60', "Jumlah Kas dan Setara Kas");
                                                           
$objPHPExcel->getActiveSheet()->setTitle('Arus Kas');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_KEU_PUBLIKASI_".$label_tgl."_".$file_eksport.".xls");

// LOAD FROM EXCEL FILE

$objPHPExcel = PHPExcel_IOFactory::load("download/Report_KEU_PUBLIKASI_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);


?>


<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> LAPORAN KEUANGAN PUBLIKASI 
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
                                        BS </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        PL </a>
                                    </li>
                                    <li >
                                        <a href="#tab_15_3" data-toggle="tab">
                                        REK ADMIN </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_4" data-toggle="tab">
                                        RASIO </a>
                                    </li>
                                    <li >
                                        <a href="#tab_15_5" data-toggle="tab">
                                        Derivatif </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_6" data-toggle="tab">
                                        KA dan CKPN </a>
                                    </li>
                                    <li >
                                        <a href="#tab_15_7" data-toggle="tab">
                                        KPMM </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_8" data-toggle="tab">
                                        Pengurus dan Pemilik </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_9" data-toggle="tab">
                                        Arus Kas </a>
                                    </li>
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_KEU_PUBLIKASI_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>
    <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b>  Laporan Posisi Keuangan </b>
                                    </div>                                  
                                        
                                        <p>
                                        
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="left" rowspan="2"><b>No</b></td>
                                                <td width="40%" align="center" rowspan="2"><b>POS - POS </b></td>
                                                <td width="10%" align="center" rowspan="2"><b>Sandi LBU </b></td>
                                                <td width="40%" align="center" colspan="2"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="20%" align="center"><b><?php echo $var_tgl; ?> </b></td>
                                                <td width="20%" align="center"><b><?php echo $var_prev_tgl;?> </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  style="font-size:12px" align="left"  colspan="5">ASET </td>
                                               
                                                </tr>

                                                <?php
                                                //$number=11;
                                                 for ($i=11; $i <=109 ; $i++) { 
                                                     
                                                    if ($i=='51') {


                                                    ?>
                                                <tr>
                                                <td  style="font-size:12px" align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                               
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
<?php
} else {
?>

                                                <tr>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

<?php    
}   
}


                                                ?>
                                                
                                              

                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>
                                  
                                     <div class="tab-pane" id="tab_15_2">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> PL </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(1);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="left"><b>No</b></td>
                                                <td width="50%" align="center"><b>POS - POS </b></td>
                                                <td width="40%" align="center" colspan="2"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="10%" align="left"><b></b></td>
                                                <td width="50%" align="center"><b> </b></td>
                                                <td width="20%" align="center" c><b>Tanggal1 </b></td>
                                                <td width="20%" align="center" c><b>Tanggal2 </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="center" ></td>
                                                <td  align="left" >PENDAPATAN DAN BEBAN OPERASIONAL </td>
                                                <td  align="center" ></td>
                                                <td  align="center"></td>
                                                </tr>
<?php
for ($i=10; $i <=103 ; $i++) { 
if ($i=='63' || $i=='77') {

 
?>
                                                  <tr>
                                                <td  align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                               
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>

<?php
} else {
?>



                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
}
}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>

                                    <div class="tab-pane" id="tab_15_3">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> REK ADMIN </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(2);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="left"><b>No</b></td>
                                                <td width="50%" align="center"><b>POS - POS </b></td>
                                                <td width="40%" align="center" colspan="2"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="10%" align="left"><b></b></td>
                                                <td width="50%" align="center"><b> </b></td>
                                                <td width="20%" align="center" c><b>Tanggal1 </b></td>
                                                <td width="20%" align="center" c><b>Tanggal2 </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="center" ></td>
                                                <td  align="left" >TAGIHAN KOMITMEN</td>
                                                <td  align="center" ></td>
                                                <td  align="center"></td>
                                                </tr>
<?php
for ($i=10; $i <=54 ; $i++) { 
 
if ($i=='15' || $i=='41' || $i=='50' )
{

?>
                                            

                                            <tr>
                                                <td  align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
} else  {
?>



                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
}
}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>

                                    <div class="tab-pane" id="tab_15_4">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> RASIO </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(3);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="left"><b>No</b></td>
                                                <td width="50%" align="center"><b>POS - POS </b></td>
                                                <td width="40%" align="center" colspan="2"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="10%" align="left"><b></b></td>
                                                <td width="50%" align="center"><b> </b></td>
                                                <td width="20%" align="center" c><b>Tanggal1 </b></td>
                                                <td width="20%" align="center" c><b>Tanggal2 </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="center" ></td>
                                                <td  align="left" >Rasio Kinerja </td>
                                                <td  align="center" ></td>
                                                <td  align="center"></td>
                                                </tr>
<?php
for ($i=11; $i <=32 ; $i++) { 
 if ($i=='22') {


?>
                                                <tr>
                                                <td  align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
} else {
?>


                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
}
}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>
                                    <div class="tab-pane" id="tab_15_5">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> DERIVATIF </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(4);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="10%" align="Right" rowspan="3"><b>No</b></td>
                                                <td width="30%" align="center" rowspan="3"><b>Transaksi </b></td>
                                                <td width="40%" align="center" colspan="5"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                               
                                                <td width="50%" align="center" rowspan="2"><b> Nilai Notional</b></td>
                                                <td width="20%" align="center" colspan="2"><b>Tujuan </b></td>
                                                <td width="20%" align="center" colspan="2"><b>Tagihan dan Liabilitas Derivatif</b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="10%" align="center"><b>Trading</b></td>
                                                <td width="10%" align="center"><b> Hedging </b></td>
                                                <td width="10%" align="center" ><b>TTagihan </b></td>
                                                <td width="10%" align="center" ><b>Liabilitas </b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
                                                
<?php
for ($i=11; $i <=40 ; $i++) { 
 
?>
                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php

}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>


                                    <div class="tab-pane" id="tab_15_6">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> KA dan CKPN</b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(5);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="left"><b>No</b></td>
                                                <td width="30%" align="center"><b>POS - POS </b></td>
                                                <td width="65%" align="center" colspan="12"><b>BANK </b></td>
                                                </tr>
                                                <tr class="active">
                                                <td width="5%" align="left"><b></b></td>
                                                <td width="30%" align="center"><b> </b></td>
                                                <td width="22%" align="center" colspan="6"><b>30 Sept 15 </b></td>
                                                <td width="23%" align="center" colspan="6"><b>30 Sept 14 </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="right" ></td>
                                                <td  align="right" ></td>
                                                <td  align="center">L</td>
                                                <td  align="center">DPK</td>
                                                <td  align="center">KL</td>
                                                <td  align="center">D</td>
                                                <td  align="center">M</td>
                                                <td  align="center">Jumlah</td>
                                                <td  align="center">L</td>
                                                <td  align="center">DPK</td>
                                                <td  align="center">KL</td>
                                                <td  align="center">D</td>
                                                <td  align="center">M</td>
                                                <td  align="center">Jumlah</td>
                                                </tr>
<?php
for ($i=11; $i <=100 ; $i++) { 
 
?>
                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("I$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("K$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("N$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php

}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>

                                    <div class="tab-pane" id="tab_15_7">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> KPMM </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(6);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="60%" align="left" colspan="3"><b></b></td>
                                                <td width="20%" align="center"><b>30-Sept 15</b></td>
                                                <td width="20%" align="center"><b>30-Sept 14</b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="left" width="5%">I</td>
                                                <td  align="left" width="55%" colspan="2" >Modal Inti (Tier 1) </td>
                                                <td  align="center"></td>
                                                <td  align="center"></td>
                                                </tr>
<?php
for ($i=10; $i <=47 ; $i++) { 
 if ($i=='41' || $i=='42' || $i=='43' || $i=='44' || $i=='45' || $i=='46' ){
?>
<tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php
}
else {
?>

                                                <tr>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>

<?php
}  //end i=41
}

?>
<tr>
                                                <td  align="center" colspan="3"><b><?php echo $objPHPExcel->getActiveSheet()->getCell("A49")->getValue(); ?></b></td>
                                                
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M49")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
                                                
                                                
                                                </tbody>
                                            </table>
                                            <br><br>
<table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="20%" align="left" ><b></b></td>
                                                <td width="15%" align="right"><b>30-Sept 15</b></td>
                                                <td width="15%" align="right"><b>30-Sept 14</b></td>
                                                <td width="15%" align="left"><b>KETERANGAN</b></td>
                                                <td width="15%" align="right"><b></b></td>
                                                <td width="15%" align="right"><b></b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
<?php
for ($i=53; $i <=58 ; $i++) { 
 
?>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("H$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("J$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("L$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("M$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                            
                                                </tr>
<?php

}

?>
                                                
                                                
                                                </tbody>
                                            </table>


                                        </div>
                                        
                                        </p>
                                    </div>

                                    <div class="tab-pane" id="tab_15_8">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> Pengurus dan Pemilik </b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(7);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                              
                                                <tr class="active">
                                                
                                                <td width="50%" align="center" colspan="2"><b>PRNGURUS BANK</b></td>
                                                <td width="50%" align="center" colspan="2"><b>PEMEGANG SAHAM </b></td>
                                                
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                
<?php

for ( $i=7; $i <=24 ; $i++) { 
 
 if ($i=='23' || $i=='24'){
?>

                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="center" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getValue(); ?></td>
                                           
                                            
                                                </tr>
<?php
} else  {
?>


                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("A$i")->getValue(); ?></td>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="left" colspan="2"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getValue(); ?></td>
                                           
                                            
                                                </tr>
<?php
}
}

?>
                                                
                                                <tr>
                                                <td  align="left"></td>
                                                <td  align="left"></td>
                                                <td  align="center" ><?php echo $objPHPExcel->getActiveSheet()->getCell("C28")->getValue(); ?></td>
                                                <td  align="center" ><?php echo $objPHPExcel->getActiveSheet()->getCell("D28")->getValue(); ?></td>
                                           
                                            
                                                </tr>
                                                <tr>
                                                <td  align="left"></td>
                                                <td  align="left"></td>
                                                <td  align="center" ><?php echo $objPHPExcel->getActiveSheet()->getCell("C29")->getValue(); ?></td>
                                                <td  align="center" ><?php echo $objPHPExcel->getActiveSheet()->getCell("D29")->getValue(); ?></td>
                                           
                                            
                                                </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>


                                    <div class="tab-pane" id="tab_15_9">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> ARUS KAS</b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(8);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                              
                                                <tr class="active">
                                                
                                                <td width="50%" align="center"><b> </b></td>
                                                <td width="20%" align="center" c><b>Tanggal1 </b></td>
                                                <td width="20%" align="center" c><b>Tanggal2 </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                
<?php

for ( $i=8; $i <=60 ; $i++) { 
 
?>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                           
                                            
                                                </tr>
<?php

}

?>
                                                
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>


                                </div>
                            </div>
                            
                        </div>
                </div>

