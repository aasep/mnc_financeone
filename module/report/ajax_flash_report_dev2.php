<?php
//require_once 'config/config.php';
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';


date_default_timezone_set("Asia/Bangkok");
$file_eksport=date('Y_m_d_H_i_s');
logActivity("generate flash-report",date('Y_m_d_H_i_s'));

error_reporting(0);


###################################
/*
// 2016-12-19
##  CREATE FILE TGL SEBELUMNYA
include "flash_report_day_min_1.php";
//$objPHPExcel = new PHPExcel();
##  output file Flash_Report_".$label_tgl."_".$file_eksport."_min1.xls
 //$objPHPExcel = PHPExcel_IOFactory::load("download/Flash_Report_30-May-16_2016_06_29_10_04_30_min1.xls");
 $objPHPExcel = PHPExcel_IOFactory::load("download/Flash_Report_".$label_tgl."_".$file_eksport."_min1.xls");
 $objWorksheet = $objPHPExcel->getActiveSheet('2');
 $m61_is_report_day_min1=$objPHPExcel->getActiveSheet()->getCell('M61')->getFormattedValue('#,##0,,;(#,##0,,);"-"');

// echo "nilai sebelumnya : ".$m61_is_report_day_min1;
 //die();
 */
###################################

$tanggal=$_POST['tanggal']; 

if (!isset($tanggal) || $tanggal=""){
$tanggal="2017-02-28";    
}


//error_reporting(0);
//$tanggal="2015-06-30";
$day=date('d',strtotime($tanggal));
$day_min1=date('j',strtotime($tanggal))-1;

$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));
$year_budget=date('Y',strtotime($tanggal));
$mon_budget=date('m',strtotime($tanggal));

$prev_date=date('t-M-y', strtotime(date('Y-m',strtotime($tanggal))." -1 month")); // tanggal terakhir pada bulan sebelumnya 

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih

$label_tgl_min1=date('d-M-y', strtotime(date('Y-m-d',strtotime($tanggal))." -1 day")); // tanggal terpilih minus (-) 1
$label_tgl_next1=date('d-M-y', strtotime(date('Y-m-d',strtotime($tanggal))." 1 day"));// tanggal next

$label_tgl_year_min1=date('d-M-y', strtotime(date('Y-m-d',strtotime($label_tgl_next1))." -1 year"));// 

$curr_tgl=date('Y-m-d',strtotime($tanggal));
$curr_tgl_min1=date('Y-m-d',strtotime($label_tgl_min1));
$curr_mon_min1=date('Y-m-d',strtotime($prev_date));

$curr_mon_min2=date('Y-m-t', strtotime(date('Y-m',strtotime($curr_mon_min1))." -1 month"));
//$curr_day_budget=
//$curr_mon_budget=
//query budget

$var_tabel=date('Ymd',strtotime($tanggal));
$var_tabel_min1=date('Ymd',strtotime($label_tgl_min1));




$query_budget=" select budget from Budget_bS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget' ";


$query_pajak=" select Nominal_Pajak from Master_Pajak where  Month(DataDate)='$mon_budget' and Year(DataDate)='$year_budget' ";
 $result_pajak=odbc_exec($connection2, $query_pajak);
        $row_pajak=odbc_fetch_array($result_pajak);
        $found_pajak=odbc_num_rows($result_pajak);
        $curr_pajak=$row_pajak['Nominal_Pajak'];

        $waktu_pajak=$mon_budget."-".$year_budget; 

        if ($found_pajak ==0 || !isset($found_pajak)){ 
                $query_pajak=" select top 1 *  from Master_Pajak order by DataDate desc ";
                $result_pajak=odbc_exec($connection2, $query_pajak);
                $row_pajak=odbc_fetch_array($result_pajak);
                $found_pajak=odbc_num_rows($result_pajak);
                $curr_pajak=$row_pajak['Nominal_Pajak'];

        $waktu_pajak=date('m-Y',strtotime($row_pajak['DataDate']));

  // echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Anda Harus Menginput Pajak di Bulan ".$mon." (".$year.")  terlebih dahulu...!</strong> </div>";
        
  


    }

//echo "<div class='alert alert-info alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Data menggunakan Pajak Bulan $waktu_pajak </strong> </div>";
  //   die();

//=======================================
$budget=0;
//==============hardcode

$var_curr_tgl="  a.DataDate='".$curr_tgl."' and a. Flag_M='Y' ";//var tgl terpilih
$var_curr_tgl_min1="  a.DataDate='".$curr_tgl_min1."' and a. Flag_M='Y' ";//var tgl terpilih minus 1
$var_curr_mon_min1="  a.DataDate='".$curr_mon_min1."' and a. Flag_M='Y' ";//var tgl terakhir bulan sebelumnya

$var_curr_mon_min2="  a.DataDate='".$curr_mon_min2."' and a. Flag_M='Y' ";//var tgl terakhir bulan sebelumnya


#------------Query to Master Modal
$q_budget=" select Nominal_Modal as modal_master from Master_Modal where Month(DataDate)='$mon_budget' and Year(DataDate)='$year_budget' ";
$res_budget=odbc_exec($connection2, $q_budget);
$row_budget=odbc_fetch_array($res_budget);
$found_budget=odbc_num_rows($res_budget);
$budget_nilai_fix=$row_budget['modal_master'];


if ($found_budget ==0 || !isset($found_budget)){
	echo "<div class='alert alert-danger alert-dismissable'><b>Data on month $mon_budget - $year_budget is not available. Please select the other date.</b> </div>";
    die();
    }


########### Perhitungan Profit Loss Current Year ##############################

$query_c55=" select Tanggal,Jumlah/1000000 as Jumlah_Nominal from ( 
select DataDate as Tanggal,sum (Nominal) as Jumlah from DM_Journal WHERE DataDate='$curr_tgl' 
AND LEFT(KODEGL,1) in ('4','5') 
group by datadate
)as tabel1 ";


        $res_c55=odbc_exec($connection2, $query_c55); 
        $row_c55=odbc_fetch_array($res_c55); 
        $c55=abs($row_c55['Jumlah_Nominal']); 
		//echo $query_c55;
		//die();

$query_d55=" select Tanggal,Jumlah/1000000 as Jumlah_Nominal from ( 
select DataDate as Tanggal,sum (Nominal) as Jumlah from DM_Journal WHERE DataDate='$curr_tgl_min1' 
AND LEFT(KODEGL,1) in ('4','5') 
group by datadate
)as tabel1 ";
        $res_d55=odbc_exec($connection2, $query_d55); 
        $row_d55=odbc_fetch_array($res_d55); 
        $d55=abs($row_d55['Jumlah_Nominal']); 



//echo $query_c55."<br>";
//echo $query_d55."<br>";
//die();





/*

$query_NPL="SELECT Flag,sum(nilai1)as NPL
FROM(
SELECT a.datadate,DATEADD(day,-1,'2016-02-24')AS Yesterday ,'2016-02-24' AS Today,sum(a.nominal) as nilai1,b.Kolektibilitas,b.status,
CASE 
WHEN b.Kolektibilitas IN ('3','4','5') AND a.DataDate = '2016-02-24' THEN '1.New NPL'
WHEN b.Kolektibilitas IN('1','2') AND a.DataDate = DATEADD(day,-1,'2016-02-24') THEN '1.New NPL'
WHEN b.Kolektibilitas IN ('3','4','5') AND a.DataDate = '2016-02-24' AND b.status IN('2','8') THEN '1.New NPL' 
WHEN b.Kolektibilitas IN ('3','4','5') AND a.DataDate = DATEADD(day,-1,'2016-02-24')THEN '2.NPL To PL'
WHEN b.Kolektibilitas IN('1','2') AND a.DataDate = '2016-02-24' THEN '2.NPL To PL' 
WHEN b.Kolektibilitas IN ('3','4','5') AND a.DataDate = DATEADD(day,-1,'2016-02-24') AND b.status IN('2','8') THEN '3.NPL Lunas' 
WHEN b.Kolektibilitas IN ('3','4','5')AND a.DataDate = DATEADD(day,-1,'2016-02-24')THEN '4.Penambah OS NPL' 
WHEN b.Kolektibilitas IN('1','2') AND a.DataDate = '2016-02-24' THEN '4.Penambah OS NPL' 
WHEN BakiDebet - BakiDebet > 0 THEN '4.Penambah OS NPL'
WHEN b.Kolektibilitas IN ('3','4','5') AND b.Kolektibilitas IN('3','4','5') THEN '5.Pengurang OS NPL' 
WHEN BakiDebet - BakiDebet > 0 THEN '5.Pengurang OS NPL'
WHEN Managed_Product_Code LIKE '%Kartu Kredit Visa%' THEN '6.NPL Movement CC'
--WHEN BakiDebet = BakiDebet AND Tabel1.Flag <> '1.New NPL' THEN '7.ExcRateDiff'
else 'Tidak Ada Status'
END as Flag
FROM DM_JOURNAL a
JOIN DM_ASETKREDIT b ON b.Managed_GL_Code = a.KodeGL
WHERE a.DataDate='2016-02-24' AND a.JenisMataUang <> 'IDR'
GROUP BY a.DataDate,b.BakiDebet,b.Kolektibilitas,b.status,b.managed_product_code
) AS Tabel1
group by flag order by flag asc
";
$result_npl=odbc_exec($connection2, $query_NPL);
$row_npl=odbc_fetch_array($result_npl);

*/


#######################ACCUMULATION########################################

//$tgl='2016-12-02';
/*
        $tgl_acc=date('Y-n-j',strtotime($tanggal));

        $bln_acc=date('n',strtotime($tanggal));
        
        if ($bln_acc > 1){

                for( $i=1;$i<$bln_acc;$i++){
    
                    $var_tgl_acc=" a.Datadate='".('Y-m-t', strtotime(date("Y-$i",strtotime($tgl_acc))."'"));

                
                    $query_acc.=" SELECT SUM(Nilai) AS jml_nominal FROM( ";
                    $query_acc.=" SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a  ";
                    $query_acc.=" JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                    $query_acc.=" JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 ;"
                    $query_acc.=" WHERE  $var_tgl_acc  $parameter2 ";
                    $query_acc.=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
                    $result_acc=odbc_exec($connection2, $query_acc);
                    $row_acc=odbc_fetch_array($result_acc);
                    $jml_acc=$row_acc['jml_nominal'];
                    if (!isset($jml_acc) || $jml_acc=="" || $jml_acc==NULL || $jml_acc=='0')
                    {
                        $jml_acc=0;
                    }
                    $tot_acc=$tot_acc+$jml_acc;


                }
        } else {

                for( $i=1;$i<13;$i++){

                    $var_tgl_acc=('Y-m-t', strtotime(date("Y-$i",strtotime($tgl_acc)). " -1 year"));
                }

            }

*/
###################### END ACCUMULATION #######################################



//and a.FLASH_Level_3_Description='Cash'

/*
$query_currentDate="select ROUND (sum (total),2) as jml_nominal from (
select DISTINCT(c.KodeGL),sum(c.Nominal) as total
from DM_Journal c
JOIN Referensi_GL_02_New b ON c.KodeGL = b.GLNO 
JOIN Referensi_Flash_Report a ON a.FLASH_LEVEL_3 = b.FLASH_LEVEL_3
WHERE  ";


$var_flash_add=" group by c.KodeGL ) as tabel1 ";
*/
/*
$query_currentDate=" SELECT SUM(total) as jml_nominal from (
select c.KodeGL,c.Nominal as total,c.KodeProduct,b.FLASH_LEVEL_3
from DM_Journal c
JOIN Referensi_GL_02_New b ON b.GLNO=c.KodeGL
JOIN Referensi_Flash_Report a ON a.FLASH_LEVEL_3 = b.FLASH_LEVEL_3 
where  ";

//c.DataDate ='2016-02-24'and b.FLASH_LEVEL_3 ='FLASH101000001'

$var_flash_add=" GROUP BY c.KodeGL,c.Nominal,c.KodeProduct,b.FLASH_LEVEL_3
) AS Tabel1 ";
*/

/* tgl 21 maret

$query_currentDate="select sum (Nominal_Baru) as jml_nominal FROM (
SELECT parameter,Nominal,isnull(debet,'0') as Adj_Debet,isnull(kredit,'0')Adj_Kredit,Nominal+isnull(debet,'0')-isnull(kredit,'0') as Nominal_Baru 
FROM (
SELECT b.FLASH_LEVEL_3 as parameter, sum(c.Nominal) AS Nominal,d.NominalDebet as debet,d.NominalKredit as kredit
FROM DM_Journal c
JOIN Referensi_GL_02_New b ON b.GLNO = c.KodeGL
JOIN Referensi_Flash_Report a ON a.FLASH_Level_3 = b.FLASH_LEVEL_3
left JOIN Adjustment_Ref d ON d.NOGL=c.KodeGL and d.BulanData = MONTH(c.DataDate) and d.TahunData = YEAR(c.DataDate)
WHERE  ";
//a.DataDate='2016-02-24' and b.FLASH_Level_3='FLASH101000014' 
$var_flash_add="  group by d.NominalDebet,d.NominalKredit,b.FLASH_LEVEL_3
)As Hitung1
group by Nominal,debet,kredit,parameter
) as tabel2
group by parameter";

*/
/*
$query_currentDate=" SELECT SUM(Nilai)/1000000 AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
LEFT JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
LEFT JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE  ";
//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
*/

$query_currentDate=" SELECT SUM(Nilai)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";

// a.DataDate='2016-05-26' AND b.FLASH_LEVEL_3 ='FLASH101000007' 
$var_flash_add=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1 ";



 //=============   case "Cash":

        $var_flash=" and b.FLASH_Level_3='FLASH101000001' ";
//echo  $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die ();


        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $cash=$row2['jml_nominal'];


        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $cash3=$row3['jml_nominal'];

        $cash5=$cash-$cash3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $cash4=$row4['jml_nominal'];
        $cash6=$cash-$cash4;



        $var_budget=" and FLASH_Level_3='FLASH101000001' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_cash=$row5['budget'];


        $cash7=$cash-$budget_cash;



//Placement with BI
//FLASH201000005
        $var_flash=" and b.FLASH_Level_3='FLASH201000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $placement_wbi=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $placement_wbi3=$row3['jml_nominal'];

        $placement_wbi5=$placement_wbi-$placement_wbi3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $placement_wbi4=$row4['jml_nominal'];
        $placement_wbi6=$placement_wbi-$placement_wbi4;


        $var_budget=" and FLASH_Level_3='FLASH201000005' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_placement_wbi=$row5['budget'];


        $placement_wbi7=$placement_wbi-$budget_placement_wbi;

 


        $var_flash=" and b.FLASH_Level_3='FLASH101000002' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account_bi=$row2['jml_nominal'];
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $current_account_bi3=$row3['jml_nominal'];

        $current_account_bi5=$current_account_bi-$current_account_bi3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $current_account_bi4=$row4['jml_nominal'];
        $current_account_bi6=$current_account_bi-$current_account_bi4;

        $var_budget=" and FLASH_Level_3='FLASH101000002' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_current_account_bi=$row5['budget'];

        $current_account_bi7=$current_account_bi-$budget_current_account_bi;

  //      break;
  //  case "Certificate of bank Indonesia (SBI & BI call money)":
        $var_flash=" and b.FLASH_Level_3='FLASH101000003' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $certificate_bi=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $certificate_bi3=$row3['jml_nominal'];

        $certificate_bi5=$certificate_bi-$certificate_bi3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $certificate_bi4=$row4['jml_nominal'];
        //FLASH101000003
        $var_budget=" and FLASH_Level_3='FLASH101000003' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_certificate_bi=$row5['budget'];

        //$var_budget=" and FLASH_Level_3='FLASH101000002' ";
        //$result5=odbc_exec($connection2, $query_budget.$var_budget);
        //$row5=odbc_fetch_array($result5);
        //$budget_current_account_bi=$row5['budget'];

        $certificate_bi7=$certificate_bi-$budget_certificate_bi;


  //      break;
  //  case "Interbank Placement":
        $var_flash=" and b.FLASH_Level_3='FLASH101000004' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $interbank_placement=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $interbank_placement3=$row3['jml_nominal'];

        $interbank_placement5=$interbank_placement-$interbank_placement3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $interbank_placement4=$row4['jml_nominal'];
        $interbank_placement6=$interbank_placement-$interbank_placement4;

        $var_budget=" and FLASH_Level_3='FLASH101000004' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_interbank_placement=$row5['budget'];

        $interbank_placement7=$interbank_placement-$budget_interbank_placement;
 //       break;
 //   case "Securities":
        $var_flash=" and b.FLASH_Level_3='FLASH101000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $scurities=$row2['jml_nominal'];
        //echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
        //die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $scurities3=$row3['jml_nominal'];

        $scurities5=$scurities-$scurities3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $scurities4=$row4['jml_nominal'];

        $scurities6=$scurities-$scurities4;

        //FLASH101000005
        $var_budget=" and FLASH_Level_3='FLASH101000005' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_scurities=$row5['budget'];


        $scurities7=$scurities-$budget_scurities;
   //     break;
  //  case "Allowance For Securities":
        $var_flash=" and b.FLASH_Level_3='FLASH101000006' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $allowence_fs=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $allowence_fs3=$row3['jml_nominal'];

        $allowence_fs5=$allowence_fs-$allowence_fs3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $allowence_fs4=$row4['jml_nominal'];
        $allowence_fs6=$allowence_fs-$allowence_fs4;

        $allowence_fs7=$allowence_fs-$budget;




#############################################################################################


//$curr_tgl=date('Y-m-d',strtotime($tanggal));
//$curr_tgl_min1=date('Y-m-d',strtotime($label_tgl_min1));
//$curr_mon_min1=date('Y-m-d',strtotime($prev_date));

$var_tabel=date('Ymd',strtotime($curr_tgl));
$table_asetkredit="DM_AsetKredit_$var_tabel";

$var_tabel_min1=date('Ymd',strtotime($curr_tgl_min1));
$table_asetkredit_min1="DM_AsetKredit_$var_tabel_min1";

$var_tabel_mon_min1=date('Ymd',strtotime($curr_mon_min1));
$table_asetkredit_mon_min1="DM_AsetKredit_$var_tabel_mon_min1";

#-- QUERY LOAN-- 
$q_loan=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit ";
$q_loan_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_min1 "; 
$q_loan_mon_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_mon_min1 "; 

//echo $q_loan."<br>".$q_loan_min1."<br>".$q_loan_mon_min1;
//die();

        $result2=odbc_exec($connection2, $q_loan);
        $row2=odbc_fetch_array($result2);
        $loans=$row2['jml_nominal'];
		//echo $q_loan;
        //die();
      
        $result3=odbc_exec($connection2, $q_loan_min1);
        $row3=odbc_fetch_array($result3);
        $loans3=$row3['jml_nominal'];

        $loans5=$loans-$loans3;

        $result4=odbc_exec($connection2, $q_loan_mon_min1);
        $row4=odbc_fetch_array($result4);
        $loans4=$row4['jml_nominal'];
        $loans6=$loans-$loans4;

        $var_budget=" and FLASH_Level_3='FLASH101000007' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_loans=$row5['budget'];

        $loans7=$loans-$budget_loans;

#--PERFORMING LOAN--
//select sum (jumlahkreditperiodelaporan) from $table_asetkredit where kolektibilitas in ('1','2')
$q_loan=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit where kolektibilitas in ('1','2') ";
$q_loan_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_min1 where kolektibilitas in ('1','2') "; 
$q_loan_mon_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_mon_min1 where kolektibilitas in ('1','2') "; 
       
        $result2=odbc_exec($connection2, $q_loan);
        $row2=odbc_fetch_array($result2);
        $performing_loan=$row2['jml_nominal'];
		//echo $q_loan;
        //die();

        $result3=odbc_exec($connection2, $q_loan_min1);
        $row3=odbc_fetch_array($result3);
        $performing_loan3=$row3['jml_nominal'];

        $performing_loan5=$performing_loan-$performing_loan3;

        $result4=odbc_exec($connection2, $q_loan_mon_min1);
        $row4=odbc_fetch_array($result4);
        $performing_loan4=$row4['jml_nominal'];
        $performing_loan6=$performing_loan-$performing_loan4;

        $performing_loan7=$performing_loan-$budget;






#--NON PERFORMING LOAN--
//select sum (jumlahkreditperiodelaporan) from $table_asetkredit where kolektibilitas in ('3','4','5')
$q_loan=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit where kolektibilitas in ('3','4','5') ";
$q_loan_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_min1 where kolektibilitas in ('3','4','5') "; 
$q_loan_mon_min1=" select sum (jumlahkreditperiodelaporan)/1000000 as jml_nominal from $table_asetkredit_mon_min1 where kolektibilitas in ('3','4','5') "; 
        
        $result2=odbc_exec($connection2, $q_loan);
        $row2=odbc_fetch_array($result2);
        $non_performing_loan=$row2['jml_nominal'];
		//echo $q_loan;
        //die();

        $result3=odbc_exec($connection2, $q_loan_min1);
        $row3=odbc_fetch_array($result3);
        $non_performing_loan3=$row3['jml_nominal'];

        $non_performing_loan5=$non_performing_loan-$non_performing_loan3;

        $result4=odbc_exec($connection2, $q_loan_mon_min1);
        $row4=odbc_fetch_array($result4);
        $non_performing_loan4=$row4['jml_nominal'];
        $non_performing_loan6=$non_performing_loan-$non_performing_loan4;


        $non_performing_loan7=$non_performing_loan-$budget;
  






   // case "Allowance For Loan":
        $var_flash=" and b.FLASH_Level_3='FLASH101000010' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $allowence_for_loan=$row2['jml_nominal'];
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $allowence_for_loan3=$row3['jml_nominal']; 

        $allowence_for_loan5=$allowence_for_loan-$allowence_for_loan3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $allowence_for_loan4=$row4['jml_nominal'];
        $allowence_for_loan6=$allowence_for_loan-$allowence_for_loan4;

        //FLASH101000010
        $var_budget=" and FLASH_Level_3='FLASH101000010' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_allowence_for_loan=$row5['budget'];

        $allowence_for_loan7=$allowence_for_loan-$budget_allowence_for_loan;
   //     break;
   // case "Acceptance receivables":
        $var_flash=" and b.FLASH_Level_3='FLASH101000011' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $acceptance_recevables=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $acceptance_recevables3=$row3['jml_nominal'];

        $acceptance_recevables5=$acceptance_recevables-$acceptance_recevables3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $acceptance_recevables4=$row4['jml_nominal'];
        $acceptance_recevables6=$acceptance_recevables-$acceptance_recevables4;

        
        $var_budget=" and FLASH_Level_3='FLASH101000011' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_acceptance_recevables=$row5['budget'];

        $acceptance_recevables7=$acceptance_recevables-$budget_acceptance_recevables;
   //     break; //==================================================================================================
   // case "Derivative receivables":
        $var_flash=" and b.FLASH_Level_3='FLASH101000012' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $deferred_receivables=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $deferred_receivables3=$row3['jml_nominal'];

        $deferred_receivables5=$deferred_receivables-$deferred_receivables3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $deferred_receivables4=$row4['jml_nominal'];
        $deferred_receivables6=$deferred_receivables-$deferred_receivables4;

        //FLASH101000012
        $var_budget=" and FLASH_Level_3='FLASH101000012' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_deferred_receivables=$row5['budget'];

 		$deferred_receivables7=$deferred_receivables-$budget_deferred_receivables;

   //     break;
   // case "Fixed assets (Property, Plant Equipment)":
        $var_flash=" and b.FLASH_Level_3='FLASH101000013' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $fixed_assets=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $fixed_assets3=$row3['jml_nominal']; 

        $fixed_assets5=$fixed_assets-$fixed_assets3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $fixed_assets4=$row4['jml_nominal'];
        $fixed_assets6=$fixed_assets-$fixed_assets4;
        
        $var_budget=" and FLASH_Level_3='FLASH101000013' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_fixed_assets=$row5['budget'];


        $fixed_assets7=$fixed_assets-$budget_fixed_assets;

   //     break;
   // case "Deferred taxes":
        $var_flash=" and b.FLASH_Level_3='FLASH101000014' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $deferred_taxes=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $deferred_taxes3=$row3['jml_nominal']; 

        $deferred_taxes5=$deferred_taxes-$deferred_taxes3;


        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $deferred_taxes4=$row4['jml_nominal'];
         $deferred_taxes6=$deferred_taxes-$deferred_taxes4;

         
        $var_budget=" and FLASH_Level_3='FLASH101000014' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_deferred_taxes=$row5['budget'];

        //echo  $query_budget.$var_budget;
        //die();

        $deferred_taxes7=$deferred_taxes-$budget_deferred_taxes;
		
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add."<br>".$query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add;
		//die();
  //      break;
  //  case "Others - Assets":
        $var_flash=" and b.FLASH_Level_3='FLASH101000015' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_assets_=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $others_assets3_=$row3['jml_nominal'];

        $others_assets5_=$others_assets_-$others_assets3_;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $others_assets4_=$row4['jml_nominal']; 
        $others_assets6_=$others_assets_-$others_assets4_;

        $others_assets7=$others_assets_-$budget;

        $var_budget=" and FLASH_Level_3='FLASH101000015' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_others_assets_=$row5['budget'];




  //      break;
  //  case "Foreclosed properties":
        $var_flash=" and b.FLASH_Level_3='FLASH101000016' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $foreclose_properties=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $foreclose_properties3=$row3['jml_nominal']; 

        $foreclose_properties5=$foreclose_properties-$foreclose_properties3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $foreclose_properties4=$row4['jml_nominal'];
        $foreclose_properties6=$foreclose_properties-$foreclose_properties4;

        
        $var_budget=" and FLASH_Level_3='FLASH101000016' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_foreclose_properties=$row5['budget'];

        $foreclose_properties7=$foreclose_properties-$budget_foreclose_properties;

    //    break;
   // case "Allowance For Foreclosed properties":
        $var_flash=" and b.FLASH_Level_3='FLASH101000017' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $allowence_for_fp=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $allowence_for_fp3=$row3['jml_nominal']; 

        $allowence_for_fp5=$allowence_for_fp-$allowence_for_fp3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $allowence_for_fp4=$row4['jml_nominal'];
        $allowence_for_fp6=$allowence_for_fp-$allowence_for_fp4;

        $allowence_for_fp7=$allowence_for_fp-$budget;
   //     break;
   // case "Account receivable":
        $var_flash=" and b.FLASH_Level_3='FLASH101000018' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $account_receivable=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $account_receivable3=$row3['jml_nominal']; 

        $account_receivable5=$account_receivable-$account_receivable3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $account_receivable4=$row4['jml_nominal'];
        $account_receivable6=$account_receivable-$account_receivable4;

        $account_receivable7=$account_receivable-$budget;

   //     break;
   // case "Others - Other assets":
        $var_flash=" and b.FLASH_Level_3='FLASH101000019' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_assets=$row2['jml_nominal'];
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $others_assets3=$row3['jml_nominal']; 


        $others_assets5=$others_assets-$others_assets3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $others_assets4=$row4['jml_nominal'];
        $others_assets6=$others_assets-$others_assets4;

        $var_budget=" and FLASH_Level_3='FLASH101000019' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_others_assets=$row5['budget'];

        $others_assets7=$others_assets-$budget_others_assets;
  //      break;
  //  case "Allowances For Suspence Account":
        $var_flash=" and b.FLASH_Level_3='FLASH101000020' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $allowence_fsa=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $allowence_fsa3=$row3['jml_nominal']; 

        $allowence_fsa5=$allowence_fsa-$allowence_fsa3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $allowence_fsa4=$row4['jml_nominal'];
        $allowence_fsa6=$allowence_fsa-$allowence_fsa4;




        $allowence_fsa7=$allowence_fsa-$budget;
   //     break;
   
  
   
  //  case "Current Account": labilities1
        $var_flash=" and b.FLASH_Level_3='FLASH102000001' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $current_account3=$row3['jml_nominal'];

        $current_account5=$current_account-$current_account3;


        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $current_account4=$row4['jml_nominal'];
        $current_account6=$current_account-$current_account4;


        $var_budget=" and FLASH_Level_3='FLASH102000001' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_current_account=$row5['budget'];
        $current_account7=$current_account-$budget_current_account;
		
		
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();
   //     break; FLASH102000002
   // case "Saving Deposits":
        $var_flash=" and b.FLASH_Level_3='FLASH102000002' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $saving_deposits=$row2['jml_nominal']; 
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $saving_deposits3=$row3['jml_nominal'];

        $saving_deposits5=$saving_deposits-$saving_deposits3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $saving_deposits4=$row4['jml_nominal'];
        $saving_deposits6=$saving_deposits-$saving_deposits4;

        
        $var_budget=" and FLASH_Level_3='FLASH102000002' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_saving_deposits=$row5['budget'];

		$saving_deposits7=$saving_deposits-$budget_saving_deposits;

   //     break;
  //  case "Time deposits":
		//FLASH102000003
        $var_flash=" and b.FLASH_Level_3='FLASH102000003' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $time_deposits=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $time_deposits3=$row3['jml_nominal']; 

        $time_deposits5=$time_deposits-$time_deposits3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $time_deposits4=$row4['jml_nominal'];
        $time_deposits6=$time_deposits-$time_deposits4;

        $var_budget=" and FLASH_Level_3='FLASH102000003' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_time_deposits=$row5['budget'];

        $time_deposits7=$time_deposits-$budget_time_deposits;
  //      break;
  //  case "Interbank":
        $var_flash=" and b.FLASH_Level_3='FLASH102000004' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $interbank=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $interbank3=$row3['jml_nominal']; 

        $interbank5=$interbank-$interbank3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $interbank4=$row4['jml_nominal'];
        $interbank6=$interbank-$interbank4;

        $var_budget=" and FLASH_Level_3='FLASH102000004' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_interbank=$row5['budget'];

        $interbank7=$interbank-$budget_interbank;
//        break;
//    case "Call Money":
        $var_flash=" and b.FLASH_Level_3='FLASH102000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $call_money=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $call_money3=$row3['jml_nominal'];

        $call_money5=$call_money-$call_money3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $call_money4=$row4['jml_nominal'];
        $call_money6=$call_money-$call_money4;


        $call_money7=$call_money-$budget;

 //       break;
 //   case "Bank deposits":
        $var_flash=" and b.FLASH_Level_3='FLASH102000006' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $bank_deposit=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $bank_deposit3=$row3['jml_nominal']; 

        $bank_deposit5=$bank_deposit-$bank_deposit3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $bank_deposit4=$row4['jml_nominal'];
        $bank_deposit6=$bank_deposit-$bank_deposit4;

        $bank_deposit7=$bank_deposit-$budget;

//        break;
//    case "Interbank Current Account":
        $var_flash=" and b.FLASH_Level_3='FLASH102000011' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account_interbank=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $current_account_interbank3=$row3['jml_nominal']; 

        $current_account_interbank5=$current_account_interbank-$current_account_interbank3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $current_account_interbank4=$row4['jml_nominal'];
        $current_account_interbank6=$current_account_interbank-$current_account_interbank4;

 $var_budget=" and FLASH_Level_3='FLASH102000011' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_current_account_interbank7=$row5['budget'];


        //case "Interbank Current Account":
        $var_flash=" and b.FLASH_Level_3='FLASH102000007' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account_interbank2=$row2['jml_nominal'];
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $current_account_interbank32=$row3['jml_nominal']; 

        $current_account_interbank52=$current_account_interbank-$current_account_interbank3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $current_account_interbank42=$row4['jml_nominal'];
        $current_account_interbank62=$current_account_interbank2-$current_account_interbank42;

        $var_budget=" and FLASH_Level_3='FLASH102000007' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_current_account_interbank72=$row5['budget'];



       // $current_account7=$current_account-$budget_current_account;

        $current_account_interbank7=$current_account_interbank-$budget_current_account_interbank7;
  //      break;
//	case "Saving accounts":

        $var_flash=" and b.FLASH_Level_3='FLASH102000008' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $saving_account=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $saving_account3=$row3['jml_nominal']; 

        $saving_account5=$saving_account-$saving_account3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $saving_account4=$row4['jml_nominal'];
        $saving_account6=$saving_account-$saving_account4;

        $saving_account7=$saving_account-$budget;

        $var_budget=" and FLASH_Level_3='FLASH102000008' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_saving_account=$row5['budget'];

        $derivative_payable7=$derivative_payable-$budget_derivative_payable;
 //       break;
 //   case "Derivative payable":
        $var_flash=" and b.FLASH_Level_3='FLASH102000009' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $derivative_payable=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $derivative_payable3=$row3['jml_nominal']; 

        $derivative_payable5=$derivative_payable-$derivative_payable3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $derivative_payable4=$row4['jml_nominal'];
        $derivative_payable6=$derivative_payable-$derivative_payable4;
        
        $var_budget=" and FLASH_Level_3='FLASH102000009' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_derivative_payable=$row5['budget'];

        $derivative_payable7=$derivative_payable-$budget_derivative_payable;
 //       break;
//    case "Acceptance payable": FLASH102000010
        $var_flash=" and b.FLASH_Level_3='FLASH102000010' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $acceptance_payable=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $acceptance_payable3=$row3['jml_nominal']; 

        $acceptance_payable5=$acceptance_payable-$acceptance_payable3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $acceptance_payable4=$row4['jml_nominal'];
        $acceptance_payable6=$acceptance_payable-$acceptance_payable4;

        $var_budget=" and FLASH_Level_3='FLASH102000010' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_acceptance_payable=$row5['budget'];

        $acceptance_payable7=$acceptance_payable-$budget_acceptance_payable;
//        break;
 //   case "KLBI Payable":
        $var_flash=" and b.FLASH_Level_3='FLASH102000012' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $klbi_payable=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $klbi_payable3=$row3['jml_nominal']; 

        $klbi_payable5=$klbi_payable-$klbi_payable3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $klbi_payable4=$row4['jml_nominal'];
        $klbi_payable6=$klbi_payable-$klbi_payable4;

        $var_budget=" and FLASH_Level_3='FLASH102000012' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_klbi_payable=$row5['budget'];
        $klbi_payable7=$klbi_payable-$budget;
//        break;
//    case "Mandatory Convertible Bonds":
        $var_flash=" and b.FLASH_Level_3='FLASH102000013' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $mandatory_convertible_bonds=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $mandatory_convertible_bonds3=$row3['jml_nominal'];

        $mandatory_convertible_bonds5=$mandatory_convertible_bonds-$mandatory_convertible_bonds3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $mandatory_convertible_bond4=$row4['jml_nominal'];
        $mandatory_convertible_bonds6=$mandatory_convertible_bonds-$mandatory_convertible_bonds4;

        $var_budget=" and FLASH_Level_3='FLASH102000013' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_mandatory_convertible_bonds=$row5['budget'];

        $mandatory_convertible_bonds7=$mandatory_convertible_bonds-$budget;
 //       break;
 //   case "Securities sold with agreement to repurchase":
        $var_flash=" and b.FLASH_Level_3='FLASH102000014' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $scurities_sold_watr=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $scurities_sold_watr3=$row3['jml_nominal'];

        $scurities_sold_watr5=$scurities_sold_watr-$scurities_sold_watr3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $scurities_sold_watr4=$row4['jml_nominal']; 
        $scurities_sold_watr6=$scurities_sold_watr-$scurities_sold_watr4;
        
        $var_budget=" and FLASH_Level_3='FLASH102000014' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_scurities_sold_watr=$row5['budget'];

        $scurities_sold_watr7=$scurities_sold_watr-$budget;
//        break;
//    case "Others Liabilities":
        $var_flash=" and b.FLASH_Level_3='FLASH102000015' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_liabilities=$row2['jml_nominal'];
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();
		

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $others_liabilities3=$row3['jml_nominal']; 

        $others_liabilities5=$others_liabilities-$others_liabilities3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $others_liabilities4=$row4['jml_nominal']; 
        $others_liabilities6=$others_liabilities-$others_liabilities4;
        
        $var_budget=" and FLASH_Level_3='FLASH102000015' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_others_liabilities=$row5['budget'];


        $others_liabilities7=$others_liabilities-$budget_others_liabilities;
 //       break;
 //   case "Paid in capital":
        $var_flash=" and b.FLASH_Level_3='FLASH103000001' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $paid_in_capital=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $paid_in_capital3=$row3['jml_nominal']; 


        $paid_in_capital5=$paid_in_capital-$paid_in_capital3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $paid_in_capital4=$row4['jml_nominal'];
        $paid_in_capital6=$paid_in_capital-$paid_in_capital4;
        
        $var_budget=" and FLASH_Level_3='FLASH103000001' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_paid_in_capital=$row5['budget'];

        $paid_in_capital7=$paid_in_capital-$budget_paid_in_capital;
//        break;
 //   case "Agio ( disagio)":
        $var_flash=" and b.FLASH_Level_3='FLASH103000002' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $agio_disagio=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $agio_disagio3=$row3['jml_nominal']; 


        $agio_disagio5=$agio_disagio-$agio_disagio3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $agio_disagio4=$row4['jml_nominal'];
        $agio_disagio6=$agio_disagio-$agio_disagio4;
        
        $var_budget=" and FLASH_Level_3='FLASH103000002' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_agio_disagio=$row5['budget'];
        $agio_disagio7=$agio_disagio-$budget_agio_disagio;
 //       break;
 //   case "General reserve":
        $var_flash=" and b.FLASH_Level_3='FLASH103000003' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $general_reserve=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $general_reserve3=$row3['jml_nominal']; 


        $general_reserve5=$general_reserve-$general_reserve3;

      

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $general_reserve4=$row4['jml_nominal'];
        $general_reserve6=$general_reserve-$general_reserve4;

        $var_budget=" and FLASH_Level_3='FLASH103000003' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_general_reserve=$row5['budget'];

        $general_reserve7=$general_reserve-$budget_general_reserve;
  //      break;
  //  case "Available for sale securities - net":
        $var_flash=" and b.FLASH_Level_3='FLASH103000004' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $available_fss_net=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $available_fss_net3=$row3['jml_nominal']; 

        

        $available_fss_net5=$available_fss_net-$available_fss_net3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $available_fss_net4=$row4['jml_nominal'];
        $available_fss_net6=$available_fss_net-$available_fss_net4;

        $var_budget=" and FLASH_Level_3='FLASH103000004' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_available_fss_net=$row5['budget'];

        $available_fss_net7=$available_fss_net-$budget_available_fss_net;
 //       break;
 //   case "Retained earnings":
        $var_flash=" and b.FLASH_Level_3='FLASH103000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $retained_earning=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $retained_earning3=$row3['jml_nominal']; 

        $retained_earning5=$retained_earning-$retained_earning3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $retained_earning4=$row4['jml_nominal'];
        $retained_earning6=$retained_earning-$retained_earning4;
        
        $var_budget=" and FLASH_Level_3='FLASH103000005' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_retained_earning=$row5['budget'];

        $retained_earning7=$retained_earning-$budget_retained_earning;
//        break;
//    case "Profit/loss current year":
        $var_flash=" and b.FLASH_Level_3='FLASH103000006' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $profit_los=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $profit_los3=$row3['jml_nominal']; 

        
        
        $profit_los5=$profit_los-$profit_los3;


        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $profit_los4=$row4['jml_nominal'];
        $profit_los6=$profit_los-$profit_los4;

        $var_budget=" and FLASH_Level_3='FLASH103000006' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_profit_los=$row5['budget'];

        $profit_los7=$profit_los-$budget_profit_los;
//        break;

//}
//echo "Process end.........";
//die();

######################### for is report #######################
$query_budget=" select Budget_MTD,Budget_YTD from Budget_PL where datepart(month,DataDate) ='$mon_budget' and datepart(year,datadate) = '$year_budget' ";
$query_budgetx=" select Budget_MTD,Budget_YTD from Budget_PL where datepart(month,DataDate) ='12' and datepart(year,datadate) = '$year_budget' ";

//Account of Expense    General Provision   FLASH202000015

$query_currentDate=" SELECT SUM(Nilai) *(-1)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";




        $var_flash=" and b.FLASH_Level_3='FLASH202000015' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $general_provision=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $general_provision3=$row3['jml_nominal']; 
        $general_provision5=$general_provision-$general_provision3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $general_provision4=$row4['jml_nominal'];
        $general_provision6=$general_provision-$general_provision4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        //$result4=odbc_exec($connection2, $query_currentDate2.$var_curr_mon_min1.$var_flash.$var_flash_add2);
        $row_m2=odbc_fetch_array($result_m2);
        $general_provision_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000015' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_general_provision=$row5['Budget_MTD'];
        $budget_general_provision2=$row5['Budget_YTD'];

        $resultx=odbc_exec($connection2, $query_budgetx.$var_budget);
        $rowx=odbc_fetch_array($resultx);
        $budget_general_provisionx=$rowx['Budget_YTD'];

        //$budget_general_provision=$row5['budget'];
        

        $general_provision7=$general_provision-$budget_general_provision;

        //$acc_general_provision=getAccumulationMonth($var_curr_tgl,$var_flash);
        //echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add."<br>".$query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add."<br>".$query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
       // die();

//Account of Expense    Specific Provision Charged  FLASH202000016
        $var_flash=" and b.FLASH_Level_3='FLASH202000016' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $specific_pc=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $specific_pc3=$row3['jml_nominal']; 
        $specific_pc5=$specific_pc-$specific_pc3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $specific_pc4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $specific_pc6=$specific_pc-$specific_pc4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $specific_pc_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000016' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_specific_pc=$row5['budget'];
        $budget_specific_pc=$row5['Budget_MTD'];
        $budget_specific_pc2=$row5['Budget_YTD'];

        $resultx=odbc_exec($connection2, $query_budgetx.$var_budget);
        $rowx=odbc_fetch_array($resultx);
        $budget_specific_pcx=$rowx['Budget_YTD'];

        $specific_pc7=$specific_pc-$budget_specific_pc;

//Account of Expense    Specific Provision Recovery     FLASH202000017
        $var_flash=" and b.FLASH_Level_3='FLASH202000017' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $specific_pr=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $specific_pr3=$row3['jml_nominal']; 
        $specific_pr5=$specific_pr-$specific_pr3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $specific_pr4=$row4['jml_nominal'];
        $specific_pr6=$specific_pr-$specific_pr4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $specific_pr_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000017' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_specific_pr=$row5['budget'];
        $budget_specific_pr=$row5['Budget_MTD'];
        $budget_specific_pr2=$row5['Budget_YTD'];

        $resultx=odbc_exec($connection2, $query_budgetx.$var_budget);
        $rowx=odbc_fetch_array($resultx);
        $budget_specific_prx=$rowx['Budget_YTD'];

        $specific_pr7=$specific_pr-$budget_specific_pr;

//Account of Expense    Write Offs Charged  FLASH202000018
        $var_flash=" and b.FLASH_Level_3='FLASH202000018' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $write_off_ch=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $write_off_ch3=$row3['jml_nominal']; 
        $write_off_ch5=$write_off_ch-$write_off_ch3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $write_off_ch4=$row4['jml_nominal'];
        $write_off_ch6=$write_off_ch-$write_off_ch4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $write_off_ch_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000018' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_write_off_ch=$row5['budget'];
        $budget_write_off_ch=$row5['Budget_MTD'];
        $budget_write_off_ch2=$row5['Budget_YTD'];

        $resultx=odbc_exec($connection2, $query_budgetx.$var_budget);
        $rowx=odbc_fetch_array($resultx);
        $budget_write_off_chx=$rowx['Budget_YTD'];

        $write_off_ch7=$write_off_ch-$budget_write_off_ch;

//Account of Expense    Write Offs Recovered    FLASH202000019
        $var_flash=" and b.FLASH_Level_3='FLASH202000019' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $write_off_rec=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $write_off_rec3=$row3['jml_nominal']; 
        $write_off_rec5=$write_off_rec-$write_off_rec3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $write_off_rec4=$row4['jml_nominal'];
        $write_off_rec6=$write_off_rec-$write_off_rec4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $write_off_rec_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000019' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_write_off_rec=$row5['budget'];
        $budget_write_off_rec=$row5['Budget_MTD'];
        $budget_write_off_rec2=$row5['Budget_YTD'];


        $write_off_rec7=$staff_cost-$budget_write_off_rec;

//Account of Expense    Foreclose Properties Provision  FLASH202000020
        $var_flash=" and b.FLASH_Level_3='FLASH202000020' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $foreclose_pp=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $foreclose_pp3=$row3['jml_nominal']; 
        $foreclose_pp5=$foreclose_pp-$foreclose_pp3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $foreclose_pp4=$row4['jml_nominal'];
        $foreclose_pp6=$foreclose_pp-$foreclose_pp4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $foreclose_pp_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000020' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_foreclose_pp=$row5['budget'];
        $budget_foreclose_pp=$row5['Budget_MTD'];
        $budget_foreclose_pp2=$row5['Budget_YTD'];

        $foreclose_pp7=$foreclose_pp-$budget_foreclose_pp;

//Account of Expense    Others Provision    FLASH202000021
        $var_flash=" and b.FLASH_Level_3='FLASH202000021' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $other_provision=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $other_provision3=$row3['jml_nominal']; 
        $other_provision5=$other_provision-$other_provision3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $other_provision4=$row4['jml_nominal'];
        $other_provision6=$other_provision-$other_provision4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $other_provision_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000021' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_other_provision=$row5['budget'];
        $budget_other_provision=$row5['Budget_MTD'];
        $budget_other_provision2=$row5['Budget_YTD'];

        $other_provision7=$other_provision-$budget_other_provision;

//Account of Expense    Total Provision FLASH202000022
        $var_flash=" and b.FLASH_Level_3='FLASH202000022' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $tot_provision=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $tot_provision3=$row3['jml_nominal']; 
        $tot_provision5=$tot_provision-$tot_provision3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $tot_provision4=$row4['jml_nominal'];
        $tot_provision6=$tot_provision-$tot_provision4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $tot_provision_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH202000022' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_tot_provision=$row5['budget'];
        $budget_tot_provision=$row5['Budget_MTD'];
        $budget_tot_provision2=$row5['Budget_YTD'];

        $tot_provision7=$tot_provision-$budget_tot_provision;

//Account of Expense    Corporate Tax   FLASH202000023
        $var_flash=" and b.FLASH_Level_3='FLASH202000023' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $corporate_tax=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $corporate_tax3=$row3['jml_nominal']; 
        $corporate_tax5=$corporate_tax-$corporate_tax3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $corporate_tax4=$row4['jml_nominal'];
        $corporate_tax6=$corporate_tax-$corporate_tax4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $corporate_tax_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000023' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_corporate_tax=$row5['budget'];
        $budget_corporate_tax=$row5['Budget_MTD'];
        $budget_corporate_tax2=$row5['Budget_YTD'];

        $corporate_tax7=$corporate_tax-$budget_corporate_tax;



//========================================== WITH OTHER INCOME (OI)====================================
        //Forex gain/(loss) on transactions FLASH201000008
$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";

// a.DataDate='2016-05-26' AND b.FLASH_LEVEL_3 ='FLASH101000007' 
$var_flash_add=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1 ";



/*  old 30-06-2016

$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
LEFT JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
LEFT JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE  ";
//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
*/
        $var_flash=" and b.FLASH_Level_3='FLASH201000008' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $forex_gain=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $forex_gain3=$row3['jml_nominal'];

        $forex_gain5=$forex_gain-$forex_gain3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $forex_gain4=$row4['jml_nominal'];
        $forex_gain6=$forex_gain-$forex_gain4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $forex_gain_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH201000008' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_forex_gain=$row5['budget'];
        $budget_forex_gain=$row5['Budget_MTD'];
        $budget_forex_gain2=$row5['Budget_YTD'];


        $forex_gain7=$forex_gain-$budget_forex_gain;


//Gain/(Loss) on sale of securities/bonds   FLASH201000009

        $var_flash=" and b.FLASH_Level_3='FLASH201000009' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $gain_loss=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $gain_loss3=$row3['jml_nominal'];

        $gain_loss5=$gain_loss-$gain_loss3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $gain_loss4=$row4['jml_nominal'];
        $gain_loss6=$gain_loss-$gain_loss4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $gain_loss_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH201000009' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_gain_loss=$row5['budget'];
        $budget_gain_loss=$row5['Budget_MTD'];
        $budget_gain_loss2=$row5['Budget_YTD'];


        $gain_loss7=$gain_loss-$budget_gain_loss;

// Remittance fee   FLASH201000010
        $var_flash=" and b.FLASH_Level_3='FLASH201000010' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $remittance_fee=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $remittance_fee3=$row3['jml_nominal'];

        $remittance_fee5=$remittance_fee-$remittance_fee3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $remittance_fee4=$row4['jml_nominal'];
        $remittance_fee6=$remittance_fee-$remittance_fee4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $remittance_fee_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000010' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_remittance_fee=$row5['budget'];
        $budget_remittance_fee=$row5['Budget_MTD'];
        $budget_remittance_fee2=$row5['Budget_YTD'];


        $remittance_fee7=$remittance_fee-$budget_remittance_fee;
// Trade Finance fee    FLASH201000011
        $var_flash=" and b.FLASH_Level_3='FLASH201000011' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $trade_finance_fee=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $trade_finance_fee3=$row3['jml_nominal'];

        $trade_finance_fee5=$trade_finance_fee-$trade_finance_fee3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $trade_finance_fee4=$row4['jml_nominal'];
        $trade_finance_fee6=$trade_finance_fee-$trade_finance_fee4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $trade_finance_fee_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000011' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_trade_finance_fee=$row5['budget'];
        $budget_trade_finance_fee=$row5['Budget_MTD'];
        $budget_trade_finance_fee2=$row5['Budget_YTD'];

        $trade_finance_fee7=$trade_finance_fee-$budget_trade_finance_fee;

// Processing fee   FLASH201000012
        $var_flash=" and b.FLASH_Level_3='FLASH201000012' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $processing_fee=$row2['jml_nominal'];
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $processing_fee3=$row3['jml_nominal'];

        $processing_fee5=$processing_fee-$processing_fee3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $processing_fee4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $processing_fee6=$processing_fee-$processing_fee4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $processing_fee_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000012' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_processing_fee=$row5['budget'];
        $budget_processing_fee=$row5['Budget_MTD'];
        $budget_processing_fee2=$row5['Budget_YTD'];

        $processing_fee7=$processing_fee-$budget_processing_fee;


// Credit Card fee  FLASH201000013
        $var_flash=" and b.FLASH_Level_3='FLASH201000013' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $credit_card_fee=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $credit_card_fee3=$row3['jml_nominal'];

        $credit_card_fee5=$credit_card_fee-$credit_card_fee3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $credit_card_fee4=$row4['jml_nominal'];
        $credit_card_fee6=$credit_card_fee-$credit_card_fee4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $credit_card_fee_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000013' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_credit_card_fee=$row5['budget'];
        $budget_credit_card_fee=$row5['Budget_MTD'];
        $budget_credit_card_fee2=$row5['Budget_YTD'];

        $credit_card_fee7=$credit_card_fee-$budget_credit_card_fee;
// Insurance Fee    FLASH201000014
         $var_flash=" and b.FLASH_Level_3='FLASH201000014' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $insurance_fee=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $insurance_fee3=$row3['jml_nominal'];

        $insurance_fee5=$insurance_fee-$insurance_fee3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $insurance_fee4=$row4['jml_nominal'];
        $insurance_fee6=$insurance_fee-$insurance_fee4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $insurance_fee_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH201000014' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_insurance_fee=$row5['budget'];
        $budget_insurance_fee=$row5['Budget_MTD'];
        $budget_insurance_fee2=$row5['Budget_YTD'];

        $insurance_fee7=$insurance_fee-$budget_insurance_fee;
// Service Charges  FLASH201000015
        $var_flash=" and b.FLASH_Level_3='FLASH201000015' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $service_charges=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $service_charges3=$row3['jml_nominal'];

        $service_charges5=$service_charges-$service_charges3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $service_charges4=$row4['jml_nominal'];
        $service_charges6=$service_charges-$service_charges4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $service_charges_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000015' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_service_chargese=$row5['budget'];
        $budget_service_charges=$row5['Budget_MTD'];
        $budget_service_charges2=$row5['Budget_YTD'];

        $service_charges7=$service_charges-$budget_service_charges;
// Other Commission & Fee   FLASH201000016
         $var_flash=" and b.FLASH_Level_3='FLASH201000016' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $other_cf=$row2['jml_nominal'];


        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $other_cf3=$row3['jml_nominal'];

        $other_cf5=$other_cf-$other_cf3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $other_cf4=$row4['jml_nominal'];
        $other_cf6=$other_cf-$other_cf4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $other_cf_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH201000016' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_other_cf=$row5['budget'];
        $budget_other_cf=$row5['Budget_MTD'];
        $budget_other_cf2=$row5['Budget_YTD'];

        $other_cf7=$other_cf-$budget_other_cf;



// Penalty  FLASH201000017
         $var_flash=" and b.FLASH_Level_3='FLASH201000017' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $penalty=$row2['jml_nominal'];
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die ();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $penalty3=$row3['jml_nominal'];

        $penalty5=$penalty-$penalty3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $penalty4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $penalty6=$penalty-$penalty4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $penalty_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000017' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_penalty=$row5['budget'];

        $budget_penalty=$row5['Budget_MTD'];
        $budget_penalty2=$row5['Budget_YTD'];



        $penalty7=$penalty-$budget_penalty;
// Write Offs Recovered FLASH201000018   revenue
         $var_flash=" and b.FLASH_Level_3='FLASH201000018' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $write_orr=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $write_orr3=$row3['jml_nominal'];

        $write_orr5=$write_orr-$write_orr3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $write_orr4=$row4['jml_nominal'];
        $write_orr6=$write_orr-$write_orr4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $write_orr_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000018' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_write_orr=$row5['budget'];
        $budget_write_orr=$row5['Budget_MTD'];
        $budget_write_orr2=$row5['Budget_YTD'];



        $write_orr7=$write_orr-$budget_write_orr;
// Other Income FLASH201000019
        $var_flash=" and b.FLASH_Level_3='FLASH201000019' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $other_income=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $other_income3=$row3['jml_nominal'];

        $other_income5=$other_income-$other_income3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $other_income4=$row4['jml_nominal'];
        $other_income6=$other_income-$other_income4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $other_income_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000019' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_other_income=$row5['budget'];
        $budget_other_income=$row5['Budget_MTD'];
        $budget_other_income2=$row5['Budget_YTD'];

        $other_income7=$other_income-$budget_other_income;
// Total Other Income   FLASH201000020
         $var_flash=" and b.FLASH_Level_3='FLASH201000020' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $total_other_income=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $total_other_income3=$row3['jml_nominal'];

        $total_other_income5=$total_other_income-$total_other_income3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $total_other_income4=$row4['jml_nominal'];
        //echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
        //die();
        $total_other_income6=$total_other_income-$total_other_income4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $total_other_income_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH201000020' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        //$budget_total_other_income=$row5['budget'];
        $budget_total_other_income=$row5['Budget_MTD'];
        $budget_total_other_income2=$row5['Budget_YTD'];

        $total_other_income7=$total_other_income-$budget_total_other_income;


    


//========================================================================================IS REPORT===================

//
$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";

// a.DataDate='2016-05-26' AND b.FLASH_LEVEL_3 ='FLASH101000007' 
$var_flash_add=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1 ";




/* old 30-06-2016

$query_currentDate=" SELECT SUM(Nilai)/1000000 AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
LEFT JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
LEFT JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE  ";
//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
*/

$query_budget=" select Budget_MTD,Budget_YTD from Budget_PL where datepart(month,DataDate) ='$mon_budget' and datepart(year,datadate) = '$year_budget' ";





        $var_flash=" and b.FLASH_LEVEL_3='FLASH201000002' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        //$result2=odbc_exec($connection2, $query_currentDate2.$var_curr_tgl.$var_flash.$var_flash_add2);
        $row2=odbc_fetch_array($result2);
        $is_loans=$row2['jml_nominal'];
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        //$result3=odbc_exec($connection2, $query_currentDate2.$var_curr_tgl_min1.$var_flash.$var_flash_add2);
        $row3=odbc_fetch_array($result3);
        $is_loans3=$row3['jml_nominal'];


        $is_loans5=$is_loans-$is_loans3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        //$result4=odbc_exec($connection2, $query_currentDate2.$var_curr_mon_min1.$var_flash.$var_flash_add2);
        $row4=odbc_fetch_array($result4);
        $is_loans4=$row4['jml_nominal'];
//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
//die();

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        //$result4=odbc_exec($connection2, $query_currentDate2.$var_curr_mon_min1.$var_flash.$var_flash_add2);
        $row_m2=odbc_fetch_array($result_m2);
        $is_loans_m2=$row_m2['jml_nominal'];

//echo $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add;
// die();     
        $is_loans6=$is_loans-$is_loans4;

        $var_budget=" and FLASH_Level_3='FLASH201000002' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);


        //$result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_loans=$row5['Budget_MTD'];
        $budget_is_loans2=$row5['Budget_YTD'];

        $loans7=$is_loans-$budget_is_loans;

$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";


// a.DataDate='2016-05-26' AND b.FLASH_LEVEL_3 ='FLASH101000007' 
$var_flash_add=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1 ";




/*
$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
LEFT JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
LEFT JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE  ";
//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
*/
//TREASURY BILL


        $var_flash=" and b.FLASH_LEVEL_3='FLASH201000003' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_treasury=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_treasury3=$row3['jml_nominal'];

        $is_treasury5=$is_treasury-$is_treasury3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_treasury4=$row4['jml_nominal'];
        $is_treasury6=$is_treasury-$is_treasury4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_treasury_m2=$row_m2['jml_nominal'];



        $var_budget=" and FLASH_Level_3='FLASH201000003' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);


        $row5=odbc_fetch_array($result5);
        $budget_is_treasury=$row5['Budget_MTD'];
        $budget_is_treasury2=$row5['Budget_YTD'];


        $is_treasury7=$is_treasury-$budget_is_treasury;





       //$acc_is_treasury=getAccumulationMonth($tanggal,$var_flash);
       //echo $acc_is_treasury;
      // die();
/*
       $tgl_acc=date('Y-n-j',strtotime($tanggal));
        $bln_acc=date('n',strtotime($tanggal));
        
        if ($bln_acc > 1){

                for( $i=1;$i<$bln_acc;$i++){
    
                    $var_tgl_acc=" a.Datadate='".date('Y-m-t', strtotime(date("Y-$i",strtotime($tgl_acc))." "))."' ";
                    //$var_tgl_acc=('Y-m-t', strtotime(date("Y-$i",strtotime($tgl_acc)).""));
                    $query_acc.=" SELECT SUM(Nilai) AS jml_nominal FROM( ";
                    $query_acc.=" SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a  ";
                    $query_acc.=" JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                    $query_acc.=" JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 ";
                    $query_acc.=" WHERE  $var_tgl_acc  $parameter2 ";
                    $query_acc.=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
                    $result_acc=odbc_exec($connection2, $query_acc);
                    $row_acc=odbc_fetch_array($result_acc);
                    $jml_acc=$row_acc['jml_nominal'];
                    if (!isset($jml_acc) || $jml_acc=="" || $jml_acc==NULL || $jml_acc=='0')
                    {
                        $jml_acc=0;
                    }
                    $tot_acc=$tot_acc+$jml_acc;

                    echo $query_acc."<br>";
                    echo $tot_acc."<br>";
                    die();
                }
        } 
*/

//  case "Interbank Placement":
        $var_flash=" and b.FLASH_LEVEL_3='FLASH201000004' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_interbank_placement=$row2['jml_nominal'];

       // echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
        //die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_interbank_placement3=$row3['jml_nominal'];

        $is_interbank_placement5=$is_interbank_placement-$is_interbank_placement3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_interbank_placement4=$row4['jml_nominal'];
        $is_interbank_placement6=$is_interbank_placement-$is_interbank_placement4;

        $var_budget=" and FLASH_Level_3='FLASH201000004' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_interbank_placement=$row5['Budget_MTD'];
        $budget_is_interbank_placement2=$row5['Budget_YTD'];

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_interbank_placement_m2=$row_m2['jml_nominal'];


        $is_interbank_placement7=$is_interbank_placement-$budget_is_interbank_placement;
        //$acc_is_interbank_placement=getAccumulationMonth($tanggal,$var_flash);
 //       break;

//Placement with BI
//FLASH201000005
        $var_flash=" and b.FLASH_LEVEL_3='FLASH201000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_placement_wbi=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_placement_wbi3=$row3['jml_nominal'];

        $is_placement_wbi5=$is_placement_wbi-$is_placement_wbi3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_placement_wbi4=$row4['jml_nominal'];
        $is_placement_wbi6=$is_placement_wbi-$is_placement_wbi4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_placement_wbi_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH201000005' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_placement_wbi=$row5['Budget_MTD'];
        $budget_is_placement_wbi2=$row5['Budget_YTD'];

        $is_placement_wbi7=$is_placement_wbi-$budget_is_placement_wbi;

        //$acc_is_placement_wbi=getAccumulationMonth($tanggal,$var_flash);

        //  case "Others - ": II - Others
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000006' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_ii_others=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_ii_others3=$row3['jml_nominal'];

        $is_ii_others5=$is_ii_others-$is_ii_others3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $is_ii_others4=$row4['jml_nominal']; 
        $is_ii_others6=$is_ii_others-$is_ii_others4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_ii_others_m2=$row_m2['jml_nominal'];    

        $var_budget=" and FLASH_Level_3='FLASH202000006' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_ii_others=$row5['Budget_MTD'];
        $budget_is_ii_others2=$row5['Budget_YTD'];


        $is_ii_others7=$is_ii_others-$budget;
        //$acc_is_ii_others=getAccumulationMonth($tanggal,$var_flash);
  //      break;



 //  case "Current Account": 
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000002' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_current_account=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_current_account3=$row3['jml_nominal'];

        $is_current_account5=$is_current_account-$is_current_account3;


        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_current_account4=$row4['jml_nominal'];
        $is_current_account6=$is_current_account-$is_current_account4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_current_account_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000002' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_current_account=$row5['Budget_MTD'];
        $budget_is_current_account2=$row5['Budget_YTD'];

        $is_current_account7=$is_current_account-$budget_is_current_account;
        
        //$acc_is_current_account=getAccumulationMonth($tanggal,$var_flash);
//  case "Saving accounts":
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000003' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_saving_account=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_saving_account3=$row3['jml_nominal']; 

        $is_saving_account5=$is_saving_account-$is_saving_account3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_saving_account4=$row4['jml_nominal'];
        $is_saving_account6=$is_saving_account-$is_saving_account4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_saving_account_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000003' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_saving_account=$row5['Budget_MTD'];
        $budget_is_saving_account2=$row5['Budget_YTD'];


        $is_saving_account7=$is_saving_account-$budget;
        //$acc_is_saving_account=getAccumulationMonth($tanggal,$var_flash);
 //       break;

 //  case "Time deposits":
        //FLASH102000003
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000004' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_time_deposits=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_time_deposits3=$row3['jml_nominal']; 

        $is_time_deposits5=$is_time_deposits-$is_time_deposits3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_time_deposits4=$row4['jml_nominal'];
        $is_time_deposits6=$is_time_deposits-$is_time_deposits4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_time_deposits_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH202000004' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_time_deposits=$row5['Budget_MTD'];
        $budget_is_time_deposits2=$row5['Budget_YTD'];

        $is_time_deposits7=$is_time_deposits-$budget_is_time_deposits;
        //$acc_is_time_deposits=getAccumulationMonth($tanggal,$var_flash);
  //      break;


//   case "Bank deposits":
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000005' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_bank_deposit=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_bank_deposit3=$row3['jml_nominal']; 

        $is_bank_deposit5=$is_bank_deposit-$is_bank_deposit3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_bank_deposit4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $is_bank_deposit6=$is_bank_deposit-$is_bank_deposit4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_bank_deposit_m2=$row_m2['jml_nominal'];

        $var_budget=" and FLASH_Level_3='FLASH202000005' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_bank_deposit=$row5['Budget_MTD'];
        $budget_is_bank_deposit2=$row5['Budget_YTD'];

        $is_bank_deposit7=$is_bank_deposit-$budget;
        //$acc_is_bank_deposit=getAccumulationMonth($tanggal,$var_flash);


 //Borrowings (MCB)
//FLASH202000007
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000007' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_borrowings=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_borrowings3=$row3['jml_nominal'];

        $is_borrowings5=$is_borrowings-$is_borrowings3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_borrowings4=$row4['jml_nominal'];
        $is_borrowings6=$is_borrowings-$is_borrowings4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_borrowings_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000007' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_borrowings=$row5['Budget_MTD'];
        $budget_is_borrowings2=$row5['Budget_YTD'];
       


        $is_borrowings7=$is_borrowings-$budget_is_borrowings;

        //$acc_is_borrowings=getAccumulationMonth($tanggal,$var_flash);

 //Guaranteed premium
//FLASH202000008

        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000008' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_guaranteed=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_guaranteed3=$row3['jml_nominal'];

        $is_guaranteed5=$is_guaranteed-$is_guaranteed3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_guaranteed4=$row4['jml_nominal'];
        $is_guaranteed6=$is_guaranteed-$is_guaranteed4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_guaranteed_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000008' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_guaranteed=$row5['Budget_MTD'];
        $budget_is_guaranteed2=$row5['Budget_YTD'];


        $is_guaranteed7=$is_guaranteed-$budget_is_guaranteed;
        $acc_is_guaranteed=getAccumulationMonth($tanggal,$var_flash);

 //  case "Others - ": 
        $var_flash=" and b.FLASH_LEVEL_3='FLASH202000009' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $is_ie_others_assets=$row2['jml_nominal'];
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add."<br>";
//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $is_ie_others_assets3=$row3['jml_nominal'];

        $is_ie_others_assets5=$is_ie_others_assets-$is_ie_others_assets3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $is_ie_others_assets4=$row4['jml_nominal']; 
//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
//die ();
        $is_ie_others_assets6=$is_ie_others_assets-$is_ie_others_assets4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $is_ie_others_assets_m2=$row_m2['jml_nominal'];
//echo $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add;

//die();

        $var_budget=" and FLASH_Level_3='FLASH202000009' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_is_ie_others_assets=$row5['Budget_MTD'];
        $budget_is_ie_others_assets2=$row5['Budget_YTD'];

        $is_ie_others_assets7=$is_ie_others_assets-$budget;
        //$acc_is_ie_others_assets=getAccumulationMonth($tanggal,$var_flash);
  //      break;


//=============OPEX=================================================
/*
$query_currentDate=" SELECT SUM(Nilai) AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_OPEX c ON c.FLASH_Level_OPEX = b.FLASH_LEVEL_1_OPEX
WHERE  ";
//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_1_OPEX )AS tabel1 ";
*/





$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM
(SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE ";

// a.DataDate='2016-05-26' AND b.FLASH_LEVEL_3 ='FLASH101000007' 
$var_flash_add=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1 ";


/*

$query_currentDate=" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM(
SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a 
LEFT JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
LEFT JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3
WHERE  ";

//a.DataDate='2016-02-24' AND b.FLASH_Level_3 ='FLASH101000001'
$var_flash_add=" GROUP BY a.kodegl ,b.FLASH_LEVEL_3 )AS tabel1 ";
*/

//echo "2769"."<br>";
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
       // die();
//Account of Expense    Staff Cost  FLASH202000010
        $var_flash=" and b.FLASH_Level_3='FLASH202000010' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $staff_cost=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $staff_cost3=$row3['jml_nominal']; 
        $staff_cost5=$staff_cost-$staff_cost3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $staff_cost4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $staff_cost6=$staff_cost-$staff_cost4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $staff_cost_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000010' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);

        $budget_staff_cost=$row5['Budget_MTD'];
        $budget_staff_cost2=$row5['Budget_YTD'];

        $staff_cost7=$staff_cost-$budget_staff_cost;
        //$acc_staff_cost=getAccumulationMonth($tanggal,$var_flash);

//echo "Staff Cost <br>";
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();

//Account of Expense General & Administrative Expenses   FLASH202000011
        $var_flash=" and b.FLASH_Level_3='FLASH202000011' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $general_ae=$row2['jml_nominal']; 
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $general_ae3=$row3['jml_nominal']; 
        $general_ae5=$general_ae-$general_ae3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $general_ae4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $general_ae6=$general_ae-$general_ae4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $general_ae_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000011' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);

        $budget_general_ae=$row5['Budget_MTD'];
        $budget_general_ae2=$row5['Budget_YTD'];

        $general_ae7=$general_ae-$budget_general_ae;
        //$acc_general_ae=getAccumulationMonth($tanggal,$var_flash);

//Account of Expense    Depreciation    FLASH202000012
        $var_flash=" and b.FLASH_Level_3='FLASH202000012' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $depreciation=$row2['jml_nominal']; 
		//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
		//die();

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $depreciation3=$row3['jml_nominal']; 
        $depreciation5=$depreciation-$depreciation3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $depreciation4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $depreciation6=$depreciation-$depreciation4;

        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $depreciation_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000012' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        
        $budget_depreciation=$row5['Budget_MTD'];
        $budget_depreciation2=$row5['Budget_YTD'];

        $depreciation7=$depreciation-$budget_depreciation;
        //$acc_depreciation=getAccumulationMonth($tanggal,$var_flash);

//Account of Expense    Other Operating Expense/Income  FLASH202000014
        $var_flash=" and b.FLASH_Level_3='FLASH202000014' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $other_oei=$row2['jml_nominal']; 
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $other_oei3=$row3['jml_nominal']; 
        $other_oei5=$other_oei-$other_oei3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $other_oei4=$row4['jml_nominal'];
		//echo $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add;
		//die();
        $other_oei6=$other_oei-$other_oei4;


        $result_m2=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min2.$var_flash.$var_flash_add);
        $row_m2=odbc_fetch_array($result_m2);
        $other_oei_m2=$row_m2['jml_nominal'];


        $var_budget=" and FLASH_Level_3='FLASH202000014' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        
        $budget_other_oei=$row5['Budget_MTD'];
        $budget_other_oei2=$row5['Budget_YTD'];

        $other_oei7=$other_oei-$budget_other_oei;


        //$acc_other_oei=getAccumulationMonth($tanggal,$var_flash);





//
//
//
//
//
//
//
////$i++;
//}

$total_assets_curr=$cash+$current_account_bi+$certificate_bi+$interbank_placement+$scurities+$allowence_fs+$loans+$performing_loan+$non_performing_loan+$allowence_for_loan+$acceptance_recevables+$deferred_receivables+$fixed_assets+$deferred_taxes+$others_assets+$foreclose_properties+$allowence_for_fp+$account_receivable+$others_assets+$allowence_fsa;
$total_assets_curr_min1=$cash3+$current_account_bi3+$certificate_bi3+$interbank_placement3+$scurities3+$allowence_fs3+$loans3+$performing_loan3+$non_performing_loan3+$allowence_for_loan3+$acceptance_recevables3+$deferred_receivables3+$fixed_assets3+$deferred_taxes3+$others_assets3+$foreclose_properties3+$allowence_for_fp3+$account_receivable3+$others_assets3+$allowence_fsa3;
$total_assets_curr_mon_min1=$cash4+$current_account_bi4+$certificate_bi4+$interbank_placement4+$scurities4+$allowence_fs4+$loans4+$performing_loan4+$non_performing_loan4+$allowence_for_loan4+$acceptance_recevables4+$deferred_receivables4+$fixed_assets4+$deferred_taxes4+$others_assets4+$foreclose_properties4+$allowence_for_fp4+$account_receivable4+$others_assets4+$allowence_fsa4;
$total_assets_var=$cash5+$current_account_bi5+$certificate_bi5+$interbank_placement5+$scurities5+$allowence_fs5+$loans5+$performing_loan5+$non_performing_loan5+$allowence_for_loan5+$acceptance_recevables5+$deferred_receivables5+$fixed_assets5+$deferred_taxes5+$others_assets5+$foreclose_properties5+$allowence_for_fp5+$account_receivable5+$others_assets5+$allowence_fsa5;
$total_assets_var_mtd="";

$total_deposit_curr=$current_account+$saving_account+$time_deposits;
$total_deposit_curr_min1=$current_account3+$saving_account3+$time_deposits3;
$total_deposit_curr_mon_min1=$current_account4+$saving_account4+$time_deposits4;
$total_deposit_var=$current_account5+$saving_account5+$time_deposits5;
$total_deposit_var_mtd="";

$total_ol_curr=$call_money+$bank_deposit+$current_account+$saving_account+$derivative_payable+$acceptance_payable+$klbi_payable+$mandatory_convertible_bonds+$scurities_sold_watr+$others_liabilities;
$total_ol_curr_min1=$call_money3+$bank_deposit3+$current_account3+$saving_account3+$derivative_payable3+$acceptance_payable3+$klbi_payable3+$mandatory_convertible_bonds3+$scurities_sold_watr3+$others_liabilities3;
$total_ol_curr_mon_min1=$call_money4+$bank_deposit4+$current_account4+$saving_account4+$derivative_payable4+$acceptance_payable4+$klbi_payable4+$mandatory_convertible_bonds4+$scurities_sold_watr4+$others_liabilities4;
$total_ol_var=$call_money5+$bank_deposit5+$current_account5+$saving_account5+$derivative_payable5+$acceptance_payable5+$klbi_payable5+$mandatory_convertible_bonds5+$scurities_sold_watr5+$others_liabilities5;
$total_ol_var_mtd="";

$total_equity_curr=$paid_in_capital+$agio_disagio+$general_reserve+$available_fss_net+$retained_earning+$profit_los;
$total_equity_curr_min1=$paid_in_capital3+$agio_disagio3+$general_reserve3+$available_fss_net3+$retained_earning3+$profit_los3;
$total_equity_var=$paid_in_capital5+$agio_disagio5+$general_reserve5+$available_fss_net5+$retained_earning5+$profit_los5;
$total_equity_curr_mon_min1=$paid_in_capital4+$agio_disagio4+$general_reserve4+$available_fss_net4+$retained_earning4+$profit_los4;



   

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
//$styleArraybackgroundRed = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignmentCenter = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignmentCenter2 = array('alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));
$styleArrayAlignmentRight= array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        ));


//BOLD
$objPHPExcel->getActiveSheet()->getStyle('B1:B3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B72:C72')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B60')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B63')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B71')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B56')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B57')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B49')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B37')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B22')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B28')->applyFromArray($styleArrayFontBold);

//NUMBER FORMAT==================
//$objPHPExcel->getActiveSheet()->getStyle('C8:J28')->getNumberFormat()->setFormatCode('#,##0.00,,;(#,##0.00,,)');
//$objPHPExcel->getActiveSheet()->getStyle('C34:J57')->getNumberFormat()->setFormatCode('#,##0.00,,;(#,##0.00,,)');

//$objPHPExcel->getActiveSheet()->getStyle('C8:H28')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C8:H28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I8:I28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
//$objPHPExcel->getActiveSheet()->getStyle('J8:J28')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J8:J28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
#,##0_);(#,##0)

$objPHPExcel->getActiveSheet()->getStyle('C34:H57')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I34:I57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J34:J57')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');






//Bakgroud
//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArraybackgroundRed);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B60:C60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B63:C63')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B72:C72')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');

//CENTER
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayAlignmentCenter2);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayAlignmentCenter2);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayAlignmentCenter);
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayAlignmentCenter);
//$objPHPExcel->getActiveSheet()->getStyle('B5:B5')->applyFromArray($styleArrayAlignmentCenter);
//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(20);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'PT BANK MNC INTERNASIONAL TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'BALANCE SHEET');
$objPHPExcel->getActiveSheet()->setCellValue('B3', $label_tgl);

//GLOBAL


//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:A1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K1:Z1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B73:C1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:J1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:Z1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B1:J1');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B2:J2');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B3:J3');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B4:J4');
//$objPHPExcel->getActiveSheet()->getStyle('B5:T6')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('A9A9A9');
$objPHPExcel->getActiveSheet()->getStyle('A1:Z4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A29:Z30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A58:Z59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:A100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A73:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('D60:Z71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('K5:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
//$objPHPExcel->getActiveSheet()->getStyle('U5:Z61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B5:B7');//Account of Assets
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C5:G6');//
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H5:J6');//
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A73:B73');//
$objPHPExcel->getActiveSheet()->getStyle('D55')->applyFromArray($styleArrayAlignmentRight);

$objPHPExcel->getActiveSheet()->setCellValue('C5', 'For The Month');
$objPHPExcel->getActiveSheet()->setCellValue('H5', $label_bln);

$objPHPExcel->getActiveSheet()->setCellValue('C7', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('D7', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('F7', $prev_date);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'Var MTD');
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'Actual');
$objPHPExcel->getActiveSheet()->setCellValue('I7', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'Var');




		
	
		
$objPHPExcel->getActiveSheet()->setCellValue('B5', 'Account of Assets');	

$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Cash');
$objPHPExcel->getActiveSheet()->setCellValue('C8', $cash);
$objPHPExcel->getActiveSheet()->setCellValue('D8', $cash3);
$objPHPExcel->getActiveSheet()->setCellValue('E8', $cash5);
$objPHPExcel->getActiveSheet()->setCellValue('F8', $cash4);
$objPHPExcel->getActiveSheet()->setCellValue('G8', $cash6);
$objPHPExcel->getActiveSheet()->setCellValue('H8', $cash);
$objPHPExcel->getActiveSheet()->setCellValue('I8',  $budget_cash);
$objPHPExcel->getActiveSheet()->setCellValue('J8', "=H8-I8");
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Current account - Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('C9', $current_account_bi);
$objPHPExcel->getActiveSheet()->setCellValue('D9', $current_account_bi3);
$objPHPExcel->getActiveSheet()->setCellValue('E9', $current_account_bi5);
$objPHPExcel->getActiveSheet()->setCellValue('F9', $current_account_bi4);
$objPHPExcel->getActiveSheet()->setCellValue('G9', $current_account_bi6);
$objPHPExcel->getActiveSheet()->setCellValue('H9', $current_account_bi);
$objPHPExcel->getActiveSheet()->setCellValue('I9', $budget_current_account_bi);
$objPHPExcel->getActiveSheet()->setCellValue('J9', $current_account_bi7);
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Certificate of bank Indonesia (SBI & BI call money)	');
$objPHPExcel->getActiveSheet()->setCellValue('C10', $certificate_bi);
$objPHPExcel->getActiveSheet()->setCellValue('D10', $certificate_bi3);
$objPHPExcel->getActiveSheet()->setCellValue('E10', $certificate_bi5);
$objPHPExcel->getActiveSheet()->setCellValue('F10', $certificate_bi4);
$objPHPExcel->getActiveSheet()->setCellValue('G10', "=(C10-F10)");
$objPHPExcel->getActiveSheet()->setCellValue('H10', $certificate_bi);
$objPHPExcel->getActiveSheet()->setCellValue('I10', $budget_certificate_bi);
$objPHPExcel->getActiveSheet()->setCellValue('J10', $certificate_bi7);
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Interbank Placement');
$objPHPExcel->getActiveSheet()->setCellValue('C11', $interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('D11', $interbank_placement3);
$objPHPExcel->getActiveSheet()->setCellValue('E11', $interbank_placement5);
$objPHPExcel->getActiveSheet()->setCellValue('F11', $interbank_placement4);
$objPHPExcel->getActiveSheet()->setCellValue('G11', $interbank_placement6);
$objPHPExcel->getActiveSheet()->setCellValue('H11', $interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('I11', $budget_interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('J11', $interbank_placement7);

$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Securities	');
$objPHPExcel->getActiveSheet()->setCellValue('C12', $scurities);
$objPHPExcel->getActiveSheet()->setCellValue('D12', $scurities3);
$objPHPExcel->getActiveSheet()->setCellValue('E12', $scurities5);
$objPHPExcel->getActiveSheet()->setCellValue('F12', $scurities4);
$objPHPExcel->getActiveSheet()->setCellValue('G12', $scurities6);
$objPHPExcel->getActiveSheet()->setCellValue('H12', $scurities);
$objPHPExcel->getActiveSheet()->setCellValue('I12', $budget_scurities);
$objPHPExcel->getActiveSheet()->setCellValue('J12', $scurities7);

$objPHPExcel->getActiveSheet()->setCellValue('B13', '-	Allowance For Securities');
$objPHPExcel->getActiveSheet()->setCellValue('C13', $allowence_fs);
$objPHPExcel->getActiveSheet()->setCellValue('D13', $allowence_fs3);
$objPHPExcel->getActiveSheet()->setCellValue('E13', $allowence_fs5);
$objPHPExcel->getActiveSheet()->setCellValue('F13', $allowence_fs4);
$objPHPExcel->getActiveSheet()->setCellValue('G13', $allowence_fs6);
$objPHPExcel->getActiveSheet()->setCellValue('H13', $allowence_fs);
$objPHPExcel->getActiveSheet()->setCellValue('J13', $allowence_fs7);

$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Loans');
$objPHPExcel->getActiveSheet()->setCellValue('C14', $loans);
$objPHPExcel->getActiveSheet()->setCellValue('D14', $loans3);
$objPHPExcel->getActiveSheet()->setCellValue('E14', $loans5);
$objPHPExcel->getActiveSheet()->setCellValue('F14', $loans4);
$objPHPExcel->getActiveSheet()->setCellValue('G14', $loans6);
$objPHPExcel->getActiveSheet()->setCellValue('H14', $loans);
$objPHPExcel->getActiveSheet()->setCellValue('I14', $budget_loans);
$objPHPExcel->getActiveSheet()->setCellValue('J14', "=H14-I14");


###########################################################################################


$objPHPExcel->getActiveSheet()->setCellValue('B15', '-	Performing Loan');
$objPHPExcel->getActiveSheet()->setCellValue('C15', $performing_loan);
$objPHPExcel->getActiveSheet()->setCellValue('D15', $performing_loan_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E15', "=C15-D15");
$objPHPExcel->getActiveSheet()->setCellValue('F15', $performing_loan_mon_min1);
$objPHPExcel->getActiveSheet()->setCellValue('G15', "=C15-F15");
$objPHPExcel->getActiveSheet()->setCellValue('H15', $performing_loan);
$objPHPExcel->getActiveSheet()->setCellValue('I15', "0");
$objPHPExcel->getActiveSheet()->setCellValue('J15', "=H15-I15");

$objPHPExcel->getActiveSheet()->setCellValue('B16', '-	Non Performing Loan*)	');
$objPHPExcel->getActiveSheet()->setCellValue('C16', $non_performing_loan);
$objPHPExcel->getActiveSheet()->setCellValue('D16', $non_performing_loan_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E16', "=C16-D16");
$objPHPExcel->getActiveSheet()->setCellValue('F16', $non_performing_loan_mon_min1);
$objPHPExcel->getActiveSheet()->setCellValue('G16', "=C16-F16");
$objPHPExcel->getActiveSheet()->setCellValue('H16', $non_performing_loan);
$objPHPExcel->getActiveSheet()->setCellValue('I16', "0");
$objPHPExcel->getActiveSheet()->setCellValue('J16', "=H16-I16");

$objPHPExcel->getActiveSheet()->setCellValue('B17', '-	Allowance For Loan	');
$objPHPExcel->getActiveSheet()->setCellValue('C17', $allowence_for_loan);
$objPHPExcel->getActiveSheet()->setCellValue('D17', $allowence_for_loan3);
$objPHPExcel->getActiveSheet()->setCellValue('E17', $allowence_for_loan5);
$objPHPExcel->getActiveSheet()->setCellValue('F17', $allowence_for_loan4);
$objPHPExcel->getActiveSheet()->setCellValue('G17', $allowence_for_loan6);
$objPHPExcel->getActiveSheet()->setCellValue('H17', $allowence_for_loan);
$objPHPExcel->getActiveSheet()->setCellValue('I17', $budget_allowence_for_loan);
$objPHPExcel->getActiveSheet()->setCellValue('J17', $allowence_for_loan7);

$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Acceptance receivables	');
$objPHPExcel->getActiveSheet()->setCellValue('C18', $acceptance_recevables);
$objPHPExcel->getActiveSheet()->setCellValue('D18', $acceptance_recevables3);
$objPHPExcel->getActiveSheet()->setCellValue('E18', $acceptance_recevables5);
$objPHPExcel->getActiveSheet()->setCellValue('F18', $acceptance_recevables4);
$objPHPExcel->getActiveSheet()->setCellValue('G18', $acceptance_recevables6);
$objPHPExcel->getActiveSheet()->setCellValue('H18', $acceptance_recevables);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $budget_acceptance_recevables);
$objPHPExcel->getActiveSheet()->setCellValue('J18', $acceptance_recevables7);

$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Derivative receivables');
$objPHPExcel->getActiveSheet()->setCellValue('C19', $deferred_receivables);
$objPHPExcel->getActiveSheet()->setCellValue('D19', $deferred_receivables3);
$objPHPExcel->getActiveSheet()->setCellValue('E19', $deferred_receivables5);
$objPHPExcel->getActiveSheet()->setCellValue('F19', $deferred_receivables4);
$objPHPExcel->getActiveSheet()->setCellValue('G19', $deferred_receivables6);
$objPHPExcel->getActiveSheet()->setCellValue('H19', $deferred_receivables);
$objPHPExcel->getActiveSheet()->setCellValue('I19', $budget_deferred_receivables);
$objPHPExcel->getActiveSheet()->setCellValue('J19', $deferred_receivables7);

$objPHPExcel->getActiveSheet()->setCellValue('B20','Fixed assets (Property, Plant Equipment)');
$objPHPExcel->getActiveSheet()->setCellValue('C20',$fixed_assets);
$objPHPExcel->getActiveSheet()->setCellValue('D20',$fixed_assets3);
$objPHPExcel->getActiveSheet()->setCellValue('E20',$fixed_assets5);
$objPHPExcel->getActiveSheet()->setCellValue('F20',$fixed_assets4);
$objPHPExcel->getActiveSheet()->setCellValue('G20',$fixed_assets6);
$objPHPExcel->getActiveSheet()->setCellValue('H20',$fixed_assets);
$objPHPExcel->getActiveSheet()->setCellValue('I20',$budget_fixed_assets);
$objPHPExcel->getActiveSheet()->setCellValue('J20',$fixed_assets7);

$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Deferred taxes	');
$objPHPExcel->getActiveSheet()->setCellValue('C21', $deferred_taxes);
$objPHPExcel->getActiveSheet()->setCellValue('D21', $deferred_taxes3);
$objPHPExcel->getActiveSheet()->setCellValue('E21', $deferred_taxes5);
$objPHPExcel->getActiveSheet()->setCellValue('F21', $deferred_taxes4);
$objPHPExcel->getActiveSheet()->setCellValue('G21', $deferred_taxes6);
$objPHPExcel->getActiveSheet()->setCellValue('H21', $deferred_taxes);
$objPHPExcel->getActiveSheet()->setCellValue('I21', $budget_deferred_taxes);
$objPHPExcel->getActiveSheet()->setCellValue('J21', $deferred_taxes7);

$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Other assets');
$objPHPExcel->getActiveSheet()->setCellValue('C22', '=SUM(C23:C27)');
$objPHPExcel->getActiveSheet()->setCellValue('D22', '=SUM(D23:D27)');
$objPHPExcel->getActiveSheet()->setCellValue('E22', '=SUM(E23:E27)');
$objPHPExcel->getActiveSheet()->setCellValue('F22', '=SUM(F23:F27)');
$objPHPExcel->getActiveSheet()->setCellValue('G22', '=SUM(G23:G27)');
$objPHPExcel->getActiveSheet()->setCellValue('H22', '=SUM(H23:H27)');
$objPHPExcel->getActiveSheet()->setCellValue('I22', '=SUM(I23:I27)');
$objPHPExcel->getActiveSheet()->setCellValue('J22', '=SUM(J23:J27)');


$objPHPExcel->getActiveSheet()->setCellValue('B23', '-	Foreclosed properties');
$objPHPExcel->getActiveSheet()->setCellValue('C23', $foreclose_properties);
$objPHPExcel->getActiveSheet()->setCellValue('D23', $foreclose_properties3);
$objPHPExcel->getActiveSheet()->setCellValue('E23', $foreclose_properties5);
$objPHPExcel->getActiveSheet()->setCellValue('F23', $foreclose_properties4);
$objPHPExcel->getActiveSheet()->setCellValue('G23', $foreclose_properties6);
$objPHPExcel->getActiveSheet()->setCellValue('H23', $foreclose_properties);
$objPHPExcel->getActiveSheet()->setCellValue('I23', $budget_foreclose_properties);
$objPHPExcel->getActiveSheet()->setCellValue('J23', $foreclose_properties7);

$objPHPExcel->getActiveSheet()->setCellValue('B24', '- 	Allowance For Foreclosed properties	');
$objPHPExcel->getActiveSheet()->setCellValue('C24', $allowence_for_fp);
$objPHPExcel->getActiveSheet()->setCellValue('D24', $allowence_for_fp3);
$objPHPExcel->getActiveSheet()->setCellValue('E24', $allowence_for_fp5);
$objPHPExcel->getActiveSheet()->setCellValue('F24', $allowence_for_fp4);
$objPHPExcel->getActiveSheet()->setCellValue('G24', $allowence_for_fp6);
$objPHPExcel->getActiveSheet()->setCellValue('H24', $allowence_for_fp);
$objPHPExcel->getActiveSheet()->setCellValue('I24', $budget_allowence_for_fp);
$objPHPExcel->getActiveSheet()->setCellValue('J24', $allowence_for_fp7);

$objPHPExcel->getActiveSheet()->setCellValue('B25', '-	Account receivable	');
$objPHPExcel->getActiveSheet()->setCellValue('C25', $account_receivable);
$objPHPExcel->getActiveSheet()->setCellValue('D25', $account_receivable3);
$objPHPExcel->getActiveSheet()->setCellValue('E25', $account_receivable5);
$objPHPExcel->getActiveSheet()->setCellValue('F25', $account_receivable4);
$objPHPExcel->getActiveSheet()->setCellValue('G25', $account_receivable6);
$objPHPExcel->getActiveSheet()->setCellValue('H25', $account_receivable);
$objPHPExcel->getActiveSheet()->setCellValue('I25', $budget_account_receivable);
$objPHPExcel->getActiveSheet()->setCellValue('J25', $account_receivable7);

$objPHPExcel->getActiveSheet()->setCellValue('B26', '-	Others');
$objPHPExcel->getActiveSheet()->setCellValue('C26', $others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('D26', $others_assets3);
$objPHPExcel->getActiveSheet()->setCellValue('E26', $others_assets5);
$objPHPExcel->getActiveSheet()->setCellValue('F26', $others_assets4);
$objPHPExcel->getActiveSheet()->setCellValue('G26', $others_assets6);
$objPHPExcel->getActiveSheet()->setCellValue('H26', $others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('I26', $budget_others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('J26', $others_assets7);

$objPHPExcel->getActiveSheet()->setCellValue('B27', '-	Allowances For Suspence Account	');
$objPHPExcel->getActiveSheet()->setCellValue('C27', $allowence_fsa);
$objPHPExcel->getActiveSheet()->setCellValue('D27', $allowence_fsa3);
$objPHPExcel->getActiveSheet()->setCellValue('E27', $allowence_fsa5);
$objPHPExcel->getActiveSheet()->setCellValue('F27', $allowence_fsa4);
$objPHPExcel->getActiveSheet()->setCellValue('G27', $allowence_fsa6);
$objPHPExcel->getActiveSheet()->setCellValue('H27', $allowence_fsa);
$objPHPExcel->getActiveSheet()->setCellValue('I27', $budget_allowence_fsa);
$objPHPExcel->getActiveSheet()->setCellValue('J27', $allowence_fsa7);

$objPHPExcel->getActiveSheet()->setCellValue('B28', 'TOTAL ASSETS');
$objPHPExcel->getActiveSheet()->setCellValue('C28', '=SUM(C8:C14)+SUM(C17:C22)');
$objPHPExcel->getActiveSheet()->setCellValue('D28', '=SUM(D8:D14)+SUM(D17:D22)');
$objPHPExcel->getActiveSheet()->setCellValue('E28', '=SUM(E8:E14)+SUM(E17:E22)');
$objPHPExcel->getActiveSheet()->setCellValue('F28', '=SUM(F8:F14)+SUM(F17:F22)');
$objPHPExcel->getActiveSheet()->setCellValue('G28', '=SUM(G8:G14)+SUM(G17:G22)');
$objPHPExcel->getActiveSheet()->setCellValue('H28', '=SUM(H8:H14)+SUM(H17:H22)');
$objPHPExcel->getActiveSheet()->setCellValue('I28', '=SUM(I8:I14)+SUM(I17:I22)');
$objPHPExcel->getActiveSheet()->setCellValue('J28', '=SUM(J8:J14)+SUM(J17:J22)');

$objPHPExcel->getActiveSheet()->setCellValue('C31', 'For The Month');
$objPHPExcel->getActiveSheet()->setCellValue('H31', $label_bln);

$objPHPExcel->getActiveSheet()->setCellValue('C33', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('D33', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('E33', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('F33', $prev_date);
$objPHPExcel->getActiveSheet()->setCellValue('G33', 'Var MTD');
$objPHPExcel->getActiveSheet()->setCellValue('H33', 'Actual');
$objPHPExcel->getActiveSheet()->setCellValue('I33', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('J33', 'Var');




$objPHPExcel->getActiveSheet()->setCellValue('B31', 'Account of Liabilities & Equity');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B31:B33');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C31:G32');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H31:J32');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Current Account');
$objPHPExcel->getActiveSheet()->setCellValue('C34', -1*($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('D34', -1*($current_account3));
$objPHPExcel->getActiveSheet()->setCellValue('E34', -1*($current_account5));
$objPHPExcel->getActiveSheet()->setCellValue('F34', -1*($current_account4));
$objPHPExcel->getActiveSheet()->setCellValue('G34', -1*($current_account6));
$objPHPExcel->getActiveSheet()->setCellValue('H34', -1*($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('I34', $budget_current_account);
$objPHPExcel->getActiveSheet()->setCellValue('J34', "=H34-I34");

$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Saving Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C35', -1*($saving_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('D35', -1*($saving_deposits3));
$objPHPExcel->getActiveSheet()->setCellValue('E35', -1*($saving_deposits5));
$objPHPExcel->getActiveSheet()->setCellValue('F35', -1*($saving_deposits4));
$objPHPExcel->getActiveSheet()->setCellValue('G35', -1*($saving_deposits6));
$objPHPExcel->getActiveSheet()->setCellValue('H35', -1*($saving_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('I35', $budget_saving_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('J35', "=H35-I35");

$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Time Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C36', -1*($time_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('D36', -1*($time_deposits3));
$objPHPExcel->getActiveSheet()->setCellValue('E36', -1*($time_deposits5));
$objPHPExcel->getActiveSheet()->setCellValue('F36', -1*($time_deposits4));
$objPHPExcel->getActiveSheet()->setCellValue('G36', -1*($time_deposits6));
$objPHPExcel->getActiveSheet()->setCellValue('H36', -1*($time_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('I36', $budget_time_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('J36', "=H36-I36");

$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Total deposits	');
$objPHPExcel->getActiveSheet()->setCellValue('C37', '=SUM(C34:C36)');
$objPHPExcel->getActiveSheet()->setCellValue('D37', '=SUM(D34:D36)');
$objPHPExcel->getActiveSheet()->setCellValue('E37', '=SUM(E34:E36)');
$objPHPExcel->getActiveSheet()->setCellValue('F37', '=SUM(F34:F36)');
$objPHPExcel->getActiveSheet()->setCellValue('G37', '=SUM(G34:G36)');
$objPHPExcel->getActiveSheet()->setCellValue('H37', '=SUM(H34:H36)');
$objPHPExcel->getActiveSheet()->setCellValue('I37', '=SUM(I34:I36)');
$objPHPExcel->getActiveSheet()->setCellValue('J37', '=SUM(J34:J36)');

$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Interbank');
$objPHPExcel->getActiveSheet()->setCellValue('C38', '=SUM(C39:C42)');
$objPHPExcel->getActiveSheet()->setCellValue('D38', '=SUM(D39:D42)');
$objPHPExcel->getActiveSheet()->setCellValue('E38', '=SUM(E39:E42)');
$objPHPExcel->getActiveSheet()->setCellValue('F38', '=SUM(F39:F42)');
$objPHPExcel->getActiveSheet()->setCellValue('G38', '=SUM(G39:G42)');
$objPHPExcel->getActiveSheet()->setCellValue('H38', '=SUM(H39:H42)');
$objPHPExcel->getActiveSheet()->setCellValue('I38', $budget_interbank);
$objPHPExcel->getActiveSheet()->setCellValue('J38', '=H38-I38');

$objPHPExcel->getActiveSheet()->setCellValue('B39', '-	Call Money');
$objPHPExcel->getActiveSheet()->setCellValue('C39', -1*($call_money));
$objPHPExcel->getActiveSheet()->setCellValue('D39', -1*($call_money3));
$objPHPExcel->getActiveSheet()->setCellValue('E39', -1*($call_money5));
$objPHPExcel->getActiveSheet()->setCellValue('F39', -1*($call_money4));
$objPHPExcel->getActiveSheet()->setCellValue('G39', -1*($call_money6));
$objPHPExcel->getActiveSheet()->setCellValue('H39', -1*($call_money));
$objPHPExcel->getActiveSheet()->setCellValue('I39', $budget_call_money);
$objPHPExcel->getActiveSheet()->setCellValue('J39', "=H39-I39");

$objPHPExcel->getActiveSheet()->setCellValue('B40', '-	Bank Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C40', -1*($bank_deposit));
$objPHPExcel->getActiveSheet()->setCellValue('D40', -1*($bank_deposit3));
$objPHPExcel->getActiveSheet()->setCellValue('E40', -1*($bank_deposit5));
$objPHPExcel->getActiveSheet()->setCellValue('F40', -1*($bank_deposit4));
$objPHPExcel->getActiveSheet()->setCellValue('G40', -1*($bank_deposit6));
$objPHPExcel->getActiveSheet()->setCellValue('H40', -1*($bank_deposit));
$objPHPExcel->getActiveSheet()->setCellValue('I40', $budget_bank_deposit);
$objPHPExcel->getActiveSheet()->setCellValue('J40', "=H40-I40");

 

$objPHPExcel->getActiveSheet()->setCellValue('B41', '-	Current account	');
$objPHPExcel->getActiveSheet()->setCellValue('C41', -1*($current_account_interbank2));
$objPHPExcel->getActiveSheet()->setCellValue('D41', -1*($current_account_interbank32));
$objPHPExcel->getActiveSheet()->setCellValue('E41', "=C41-D41");
$objPHPExcel->getActiveSheet()->setCellValue('F41', -1*($current_account_interbank42));
$objPHPExcel->getActiveSheet()->setCellValue('G41', -1*($current_account_interbank62));
$objPHPExcel->getActiveSheet()->setCellValue('H41', -1*($current_account_interbank2));
$objPHPExcel->getActiveSheet()->setCellValue('I41', $budget_current_account_interbank72);
$objPHPExcel->getActiveSheet()->setCellValue('J41', "=H41-I41");

$objPHPExcel->getActiveSheet()->setCellValue('B42', '-	Saving Account	');
$objPHPExcel->getActiveSheet()->setCellValue('C42', -1*($saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('D42', -1*($saving_account3));
$objPHPExcel->getActiveSheet()->setCellValue('E42', -1*($saving_account5));
$objPHPExcel->getActiveSheet()->setCellValue('F42', -1*($saving_account4));
$objPHPExcel->getActiveSheet()->setCellValue('G42', -1*($saving_account6));
$objPHPExcel->getActiveSheet()->setCellValue('H42', -1*($saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('I42', $budget_saving_account);
$objPHPExcel->getActiveSheet()->setCellValue('J42', "=H42-I42");

$objPHPExcel->getActiveSheet()->setCellValue('B43', 'Derivative payable	');
$objPHPExcel->getActiveSheet()->setCellValue('C43', -1*($derivative_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D43', -1*($derivative_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E43', -1*($derivative_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F43', -1*($derivative_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G43', -1*($derivative_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H43', -1*($derivative_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I43', $budget_derivative_payable);
$objPHPExcel->getActiveSheet()->setCellValue('J43', "=H43-I43");

$objPHPExcel->getActiveSheet()->setCellValue('B44', 'Acceptance payable	');
$objPHPExcel->getActiveSheet()->setCellValue('C44', -1*($acceptance_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D44', -1*($acceptance_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E44', -1*($acceptance_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F44', -1*($acceptance_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G44', -1*($acceptance_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H44', -1*($acceptance_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I44', $budget_acceptance_payable);
$objPHPExcel->getActiveSheet()->setCellValue('J44', "=H44-I44");

$objPHPExcel->getActiveSheet()->setCellValue('B45', 'KLBI Payable');
$objPHPExcel->getActiveSheet()->setCellValue('C45', -1*($klbi_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D45', -1*($klbi_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E45', -1*($klbi_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F45', -1*($klbi_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G45', -1*($klbi_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H45', -1*($klbi_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I45', $budget_klbi_payable);
$objPHPExcel->getActiveSheet()->setCellValue('J45', "=H45-I45");

$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Mandatory Convertible Bonds');
$objPHPExcel->getActiveSheet()->setCellValue('C46', -1*($mandatory_convertible_bonds));
$objPHPExcel->getActiveSheet()->setCellValue('D46', -1*($mandatory_convertible_bonds3));
$objPHPExcel->getActiveSheet()->setCellValue('E46', -1*($mandatory_convertible_bonds5));
$objPHPExcel->getActiveSheet()->setCellValue('F46', -1*($mandatory_convertible_bonds4));
$objPHPExcel->getActiveSheet()->setCellValue('G46', -1*($mandatory_convertible_bonds6));
$objPHPExcel->getActiveSheet()->setCellValue('H46', -1*($mandatory_convertible_bonds));
$objPHPExcel->getActiveSheet()->setCellValue('I46', $budget_mandatory_convertible_bonds);
$objPHPExcel->getActiveSheet()->setCellValue('J46', "=H46-I46");

$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Securities sold with agreement to repurchase');
$objPHPExcel->getActiveSheet()->setCellValue('C47', -1*($scurities_sold_watr));
$objPHPExcel->getActiveSheet()->setCellValue('D47', -1*($scurities_sold_watr3));
$objPHPExcel->getActiveSheet()->setCellValue('E47', -1*($scurities_sold_watr5));
$objPHPExcel->getActiveSheet()->setCellValue('F47', -1*($scurities_sold_watr4));
$objPHPExcel->getActiveSheet()->setCellValue('G47', -1*($scurities_sold_watr6));
$objPHPExcel->getActiveSheet()->setCellValue('H47', -1*($scurities_sold_watr));
$objPHPExcel->getActiveSheet()->setCellValue('I47', $budget_scurities_sold_watr);
$objPHPExcel->getActiveSheet()->setCellValue('J47', "=H47-I47");

$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('C48', -1*($others_liabilities));
$objPHPExcel->getActiveSheet()->setCellValue('D48', -1*($others_liabilities3));
$objPHPExcel->getActiveSheet()->setCellValue('E48', -1*($others_liabilities5));
$objPHPExcel->getActiveSheet()->setCellValue('F48', -1*($others_liabilities4));
$objPHPExcel->getActiveSheet()->setCellValue('G48', -1*($others_liabilities6));
$objPHPExcel->getActiveSheet()->setCellValue('H48', -1*($others_liabilities));
$objPHPExcel->getActiveSheet()->setCellValue('I48', $budget_others_liabilities);
$objPHPExcel->getActiveSheet()->setCellValue('J48', "=H48-I48");

$objPHPExcel->getActiveSheet()->setCellValue('B49', 'Total Other Liabilities');
$objPHPExcel->getActiveSheet()->setCellValue('C49', "=SUM(C39:C48)");
$objPHPExcel->getActiveSheet()->setCellValue('D49', "=SUM(D39:D48)");
$objPHPExcel->getActiveSheet()->setCellValue('E49', "=SUM(E39:E48)");
$objPHPExcel->getActiveSheet()->setCellValue('F49', "=SUM(F39:F48)");
$objPHPExcel->getActiveSheet()->setCellValue('G49', "=SUM(G39:G48)");
$objPHPExcel->getActiveSheet()->setCellValue('H49', "=SUM(H39:H48)");
$objPHPExcel->getActiveSheet()->setCellValue('I49', "=+I38+SUM(I43:I48)");
$objPHPExcel->getActiveSheet()->setCellValue('J49', "=SUM(J39:J48)");



//$colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue()

for ($i=8;$i<29;$i++) {
	//$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=8;$i<29;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
//=========
for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}

for ($i=34;$i<58;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
		$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}

for ($i=61;$i<64;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}

for ($i=65;$i<71;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}

//$objPHPExcel->getActiveSheet()->getStyle('C34:H57')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C34:H57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I34:I57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('J34:J57')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J34:J57')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


//$objPHPExcel->getActiveSheet()->getStyle('C61:C62')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('C65:C70')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C60:C63')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('C65:C72')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->setCellValue('B50', 'Paid in capital');
$objPHPExcel->getActiveSheet()->setCellValue('C50', -1*($paid_in_capital));
$objPHPExcel->getActiveSheet()->setCellValue('D50', -1*($paid_in_capital3));
$objPHPExcel->getActiveSheet()->setCellValue('E50', -1*($paid_in_capital5));
$objPHPExcel->getActiveSheet()->setCellValue('F50', -1*($paid_in_capital4));
$objPHPExcel->getActiveSheet()->setCellValue('G50', -1*($paid_in_capital6));
$objPHPExcel->getActiveSheet()->setCellValue('H50', -1*($paid_in_capital));
$objPHPExcel->getActiveSheet()->setCellValue('I50', $budget_paid_in_capital);
$objPHPExcel->getActiveSheet()->setCellValue('J50', "=H50-I50");

$objPHPExcel->getActiveSheet()->setCellValue('B51', 'Agio ( disagio)');
$objPHPExcel->getActiveSheet()->setCellValue('C51', -1*($agio_disagio));
$objPHPExcel->getActiveSheet()->setCellValue('D51', -1*($agio_disagio3));
$objPHPExcel->getActiveSheet()->setCellValue('E51', -1*($agio_disagio5));
$objPHPExcel->getActiveSheet()->setCellValue('F51', -1*($agio_disagio4));
$objPHPExcel->getActiveSheet()->setCellValue('G51', -1*($agio_disagio6));
$objPHPExcel->getActiveSheet()->setCellValue('H51', -1*($agio_disagio));
$objPHPExcel->getActiveSheet()->setCellValue('I51', $budget_agio_disagio);
$objPHPExcel->getActiveSheet()->setCellValue('J51', "=H51-I51");

$objPHPExcel->getActiveSheet()->setCellValue('B52', 'General reserve');
$objPHPExcel->getActiveSheet()->setCellValue('C52', -1*($general_reserve));
$objPHPExcel->getActiveSheet()->setCellValue('D52', -1*($general_reserve3));
$objPHPExcel->getActiveSheet()->setCellValue('E52', -1*($general_reserve5));
$objPHPExcel->getActiveSheet()->setCellValue('F52', -1*($general_reserve4));
$objPHPExcel->getActiveSheet()->setCellValue('G52', -1*($general_reserve6));
$objPHPExcel->getActiveSheet()->setCellValue('H52', -1*($general_reserve));
$objPHPExcel->getActiveSheet()->setCellValue('I52', $budget_general_reserve);
$objPHPExcel->getActiveSheet()->setCellValue('J52', "=H52-I52");

$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Available for sale securities - net');
$objPHPExcel->getActiveSheet()->setCellValue('C53', -1*($available_fss_net));
$objPHPExcel->getActiveSheet()->setCellValue('D53', -1*($available_fss_net3));
$objPHPExcel->getActiveSheet()->setCellValue('E53', -1*($available_fss_net5));
$objPHPExcel->getActiveSheet()->setCellValue('F53', -1*($available_fss_net4));
$objPHPExcel->getActiveSheet()->setCellValue('G53', -1*($available_fss_net6));
$objPHPExcel->getActiveSheet()->setCellValue('H53', -1*($available_fss_net));
$objPHPExcel->getActiveSheet()->setCellValue('I53', $budget_available_fss_net);
$objPHPExcel->getActiveSheet()->setCellValue('J53', "=H53-I53");

$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Retained earnings');
$objPHPExcel->getActiveSheet()->setCellValue('C54', -1*($retained_earning));
$objPHPExcel->getActiveSheet()->setCellValue('D54', -1*($retained_earning3));
$objPHPExcel->getActiveSheet()->setCellValue('E54', -1*($retained_earning5));
$objPHPExcel->getActiveSheet()->setCellValue('F54', -1*($retained_earning4));
$objPHPExcel->getActiveSheet()->setCellValue('G54', -1*($retained_earning6));
$objPHPExcel->getActiveSheet()->setCellValue('H54', -1*($retained_earning));
$objPHPExcel->getActiveSheet()->setCellValue('I54', $budget_retained_earning);
$objPHPExcel->getActiveSheet()->setCellValue('J54', "=H54-I54");

$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Profit/loss current year');



//objPHPExcel->getActiveSheet()->setCellValue('C55', "=IS_Flash_Report!M61");
$objPHPExcel->getActiveSheet()->setCellValue('C55', $c55);
//$objPHPExcel->getActiveSheet()->setCellValue('D55', $m61_is_report_day_min1);
$objPHPExcel->getActiveSheet()->setCellValue('D55', $d55);
$objPHPExcel->getActiveSheet()->setCellValue('E55', "=C55-D55");
//$objPHPExcel->getActiveSheet()->setCellValue('F55', "=IS_Flash_Report!T61");
$objPHPExcel->getActiveSheet()->setCellValue('F55', $c55);
$objPHPExcel->getActiveSheet()->setCellValue('G55', "=C55-F55");
//$objPHPExcel->getActiveSheet()->setCellValue('H55', "=IS_Flash_Report!M61");
$objPHPExcel->getActiveSheet()->setCellValue('H55', $c55);
$objPHPExcel->getActiveSheet()->setCellValue('I55', $budget_profit_los);
$objPHPExcel->getActiveSheet()->setCellValue('J55', "=H55-I55");

$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Total Equity');
$objPHPExcel->getActiveSheet()->setCellValue('C56', '=SUM(C50:C55)');
$objPHPExcel->getActiveSheet()->setCellValue('D56', '=SUM(D50:D55)');
$objPHPExcel->getActiveSheet()->setCellValue('E56', '=SUM(E50:E55)');
$objPHPExcel->getActiveSheet()->setCellValue('F56', '=SUM(F50:F55)');
$objPHPExcel->getActiveSheet()->setCellValue('G56', '=SUM(G50:G55)');
$objPHPExcel->getActiveSheet()->setCellValue('H56', '=SUM(H50:H55)');
$objPHPExcel->getActiveSheet()->setCellValue('I56', '=SUM(I50:I55)');
$objPHPExcel->getActiveSheet()->setCellValue('J56', '=SUM(J50:J55)');


$objPHPExcel->getActiveSheet()->setCellValue('B57', 'TOTAL LIABILITIES & EQUITY');
$objPHPExcel->getActiveSheet()->setCellValue('C57', '=C37+C49+C56');
$objPHPExcel->getActiveSheet()->setCellValue('D57', '=D37+D49+D56');
$objPHPExcel->getActiveSheet()->setCellValue('E57', '=E37+E49+E56');
$objPHPExcel->getActiveSheet()->setCellValue('F57', '=F37+F49+F56');
$objPHPExcel->getActiveSheet()->setCellValue('G57', '=G37+G49+G56');
$objPHPExcel->getActiveSheet()->setCellValue('H57', '=H37+H49+H56');
$objPHPExcel->getActiveSheet()->setCellValue('I57', '=I37+I49+I56');
$objPHPExcel->getActiveSheet()->setCellValue('J57', '=J37+J49+J56');
	




#########################################################
#       NPL QUERY
#########################################################
# TOTAL NPL

//remark 2016-12-19
/*
$q_tot_npl =" SELECT 'NPL $curr_tgl' Ket, SUM(a.JumlahKreditPeriodeLaporan)NPL FROM Dm_AsetKredit_$var_tabel a ";
$q_tot_npl.=" WHERE a.DataDate='$curr_tgl' AND a.Kolektibilitas IN('3','4','5') AND a.[Status] NOT IN ('2','8')  ";
$res_tot_npl=odbc_exec($connection2, $q_tot_npl);
$row_total_npl=odbc_fetch_array($res_tot_npl);
$total_npl=$row_total_npl['NPL'];

$q_tot_npl =" SELECT 'NPL $curr_tgl_min1' Ket, SUM(a.JumlahKreditPeriodeLaporan)NPL FROM Dm_AsetKredit_$var_tabel_min1 a ";
$q_tot_npl.=" WHERE a.DataDate='$curr_tgl_min1' AND a.Kolektibilitas IN('3','4','5') AND a.[Status] NOT IN ('2','8')  ";
$res_tot_npl=odbc_exec($connection2, $q_tot_npl);
$row_total_npl=odbc_fetch_array($res_tot_npl);
$total_npl_min1=$row_total_npl['NPL'];


odbc_exec($connection2, " drop table temp_nplx ");
$q_sp=" exec CTABLE_NPL @datadate1='$curr_tgl_min1', @datadate2='$curr_tgl' ";
odbc_exec($connection2, $q_sp);

$q_kat_npl=" select kat, sum(nplMovement) as jum from temp_nplx group by kat order by kat asc ";
$res_kat_npl=odbc_exec($connection2, $q_kat_npl);
while ($row_kat_npl=odbc_fetch_array($res_kat_npl)){
# NEW_NPL
if ($row_kat_npl['kat']=='01-New NPL') {
    $new_npl=$row_kat_npl['jum'];
}

# NPL Reclass to PL
if ($row_kat_npl['kat']=='02-NPL Reclass to PL') {
    $npl_reclass=$row_kat_npl['jum'];
}

# NPL Paid Off
if ($row_kat_npl['kat']=='03-NPL Paid Off') {
    $npl_paid_off=$row_kat_npl['jum'];
}

# NPL Payment
if ($row_kat_npl['kat']=='04-NPL Payment') {
    $npl_payment=$row_kat_npl['jum'];
}

# NPL Added
if ($row_kat_npl['kat']=='05-NPL Added') {
    $npl_add=$row_kat_npl['jum'];
}
}
*/



$objPHPExcel->getActiveSheet()->setCellValue('B60', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('C60', $total_npl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'New NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('C61', '0');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Penambah_OS_NPL');
$objPHPExcel->getActiveSheet()->setCellValue('C62', '0');
$objPHPExcel->getActiveSheet()->setCellValue('B63', 'Total New NPL');
$objPHPExcel->getActiveSheet()->setCellValue('C63', $new_npl);
$objPHPExcel->getActiveSheet()->setCellValue('B64', '');	
$objPHPExcel->getActiveSheet()->setCellValue('B65', 'NPL to PL (Reklass) ');
$objPHPExcel->getActiveSheet()->setCellValue('C65', $npl_reclass);
$objPHPExcel->getActiveSheet()->setCellValue('B66', 'NPL Paid Off');
$objPHPExcel->getActiveSheet()->setCellValue('C66', $npl_paid_off);
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'Reverse Saldo NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('C67', '0');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'NPL Payment');
$objPHPExcel->getActiveSheet()->setCellValue('C68', $npl_payment);
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'NPL Exchange Rate');
$objPHPExcel->getActiveSheet()->setCellValue('C69', '0');
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'NPL Credit Card');
$objPHPExcel->getActiveSheet()->setCellValue('C70', '0');
$objPHPExcel->getActiveSheet()->setCellValue('B71', 'NPL Added');
$objPHPExcel->getActiveSheet()->setCellValue('C71', $npl_add);
$objPHPExcel->getActiveSheet()->setCellValue('B72', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C72', $total_npl);

// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('BS_Flash Report');

//=======BORDER
$styleArray = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$objPHPExcel->getActiveSheet()->getStyle('B5:J28')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B31:J57')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B60:C72')->applyFromArray($styleArray);
//=======END BORDER





$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1); 

//width dimension
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);

$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(20);
//$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(20);

//MERGE
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B5:B6');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C5:L5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('M5:Q5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('R5:S5');

// CENTER
$objPHPExcel->getActiveSheet()->getStyle('B5:T6')->applyFromArray($styleArrayAlignmentCenter);
$objPHPExcel->getActiveSheet()->getStyle('B5:T6')->applyFromArray($styleArrayAlignmentCenter2);
//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArrayAlignmentCenter2);


// NUMBER FORMAT===========
//$objPHPExcel->getActiveSheet()->getStyle('C8:J61')->getNumberFormat()->setFormatCode('#,##0.00,,;(#,##0.00,,)');
$objPHPExcel->getActiveSheet()->getStyle('C8:D23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F8:H23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J8:K23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('P8:P23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getStyle('I8:I23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L8:L23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('M8:O23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('Q8:Q23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('S8:T23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('R8:R23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('I26:I39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->getStyle('C26:D39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F26:H39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J26:K39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('F26:T39')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I26:I39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('P26:P39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L26:L39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('M26:O39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('Q26:Q39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('S26:T39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('R26:R39')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget

//$objPHPExcel->getActiveSheet()->getStyle('C42:D47')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('F42:T47')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');

$objPHPExcel->getActiveSheet()->getStyle('C42:D47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F42:H47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I42:I47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J42:K47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L42:L47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('M42:O47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('Q42:Q47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('S42:T47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('R42:R47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('P42:P47')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

//$objPHPExcel->getActiveSheet()->getStyle('C49:D61')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('F49:T61')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');


$objPHPExcel->getActiveSheet()->getStyle('C49:D61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('F49:H61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('J49:K61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('I49:I61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('L49:L61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('M49:O61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('Q49:Q61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('S49:T61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('R49:R61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"'); // untuk budget
$objPHPExcel->getActiveSheet()->getStyle('P49:P61')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');



// BOLD

$objPHPExcel->getActiveSheet()->getStyle('B5:T6')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B1:B3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B14')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B19')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B23')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B39')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B40')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B61')->applyFromArray($styleArrayFontBold);
//$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayFontBold);
// background

$objPHPExcel->getActiveSheet()->getStyle('B5:T6')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('A9A9A9');
$objPHPExcel->getActiveSheet()->getStyle('A1:Z4')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:A100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('B62:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('U5:Z61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');



$objPHPExcel->getActiveSheet()->getStyle('B8:T8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B14:T14')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B19:T19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B23:T23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B38:T38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B39:T39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B45:T45')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B47:T47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B59:T59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');
$objPHPExcel->getActiveSheet()->getStyle('B61:T61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');

//#A9A9A9
//$objPHPExcel->getActiveSheet()->getStyle('I5:I61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#FFFF00');
//$objPHPExcel->getActiveSheet()->getStyle('P5:P61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#FFFF00');


//$objPHPExcel->getActiveSheet()->getStyle('B59:T59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');
//$objPHPExcel->getActiveSheet()->getStyle('B61:T61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');
////$objPHPExcel->getActiveSheet()->getStyle('B47:T47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');
//$objPHPExcel->getActiveSheet()->getStyle('B45:T45')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');
//$objPHPExcel->getActiveSheet()->getStyle('B39:T39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');
//$objPHPExcel->getActiveSheet()->getStyle('B23:T23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('#A9A9A9');


//59,61,47,45,39,23
//	#FFFF00
// Add some data to the second sheet, resembling some different data types

$objPHPExcel->getActiveSheet()->setCellValue('B1', 'PT BANK MNC INTERNASIONAL TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'INCOME STATEMENTS - MTD AND YTD '.$label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('B3', '(Amounts in Rp millions)');

$objPHPExcel->getActiveSheet()->setCellValue('B5', 'Description');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Interest Income');
$objPHPExcel->getActiveSheet()->setCellValue('C8', "=SUM(C9:C13)");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "=SUM(D9:D13)");
$objPHPExcel->getActiveSheet()->setCellValue('E8', "=+IF(D8=0,0,IF(C8=0,0,(D8/C8)))");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "=SUM(F9:F13)");
$objPHPExcel->getActiveSheet()->setCellValue('G8', "=SUM(G9:G13)");
$objPHPExcel->getActiveSheet()->setCellValue('H8', "=SUM(H9:H13)");
$objPHPExcel->getActiveSheet()->setCellValue('M8', "=SUM(M9:M13)");
$objPHPExcel->getActiveSheet()->setCellValue('L8', "=SUM(L9:L13)");
$objPHPExcel->getActiveSheet()->setCellValue('Q8', "=SUM(Q9:Q13)");
$objPHPExcel->getActiveSheet()->setCellValue('J8', "=SUM(J9:J13)");
$objPHPExcel->getActiveSheet()->setCellValue('N8', "=SUM(N9:N13)");
$objPHPExcel->getActiveSheet()->setCellValue('K8', "=+IF(J8=0,0,(J8/L8))");
$objPHPExcel->getActiveSheet()->setCellValue('O8', "=+IF(N8=0,0,(N8/Q8))");
$objPHPExcel->getActiveSheet()->setCellValue('S8', "=+IF(M8=0,0,IF(R8=0,0,(M8/R8)))");


$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Loan');

$objPHPExcel->getActiveSheet()->setCellValue('C9', $is_loans4-$is_loans_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D9', "=(F9-C9)");
//$objPHPExcel->getActiveSheet()->setCellValue('E9', "=(D9/C9)");
$objPHPExcel->getActiveSheet()->setCellValue('E9', "=+IF(D9=0,0,IF(C9=0,0,(D9/C9)))");
$objPHPExcel->getActiveSheet()->setCellValue('F9', "=(M9-T9)");
$objPHPExcel->getActiveSheet()->setCellValue('G9', "=$is_loans3-T9");
$objPHPExcel->getActiveSheet()->setCellValue('H9', "=+F9-G9");
$objPHPExcel->getActiveSheet()->setCellValue('M9', $is_loans);
$objPHPExcel->getActiveSheet()->setCellValue('L9', $budget_is_loans);
$objPHPExcel->getActiveSheet()->setCellValue('Q9', $budget_is_loans2);
$objPHPExcel->getActiveSheet()->setCellValue('J9', "=(F9-L9)");
$objPHPExcel->getActiveSheet()->setCellValue('N9', "=(M9-Q9)");
$objPHPExcel->getActiveSheet()->setCellValue('K9', "=+IF(J9=0,0,(J9/L9))");
$objPHPExcel->getActiveSheet()->setCellValue('O9', "=+IF(N9=0,0,(N9/Q9))");
$objPHPExcel->getActiveSheet()->setCellValue('S9', "=+IF(M9=0,0,IF(R9=0,0,(M9/R9)))");

$objPHPExcel->getActiveSheet()->setCellValue('T9', "$is_loans4");


$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Treasury bills');
$objPHPExcel->getActiveSheet()->setCellValue('C10', $is_treasury4-$is_treasury_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D10', "=(F10-C10)");
//$objPHPExcel->getActiveSheet()->setCellValue('E10', "=(D10/C10)");
$objPHPExcel->getActiveSheet()->setCellValue('E10', "=+IF(D10=0,0,IF(C10=0,0,(D10/C10)))");
$objPHPExcel->getActiveSheet()->setCellValue('F10', "=(M10-T10)");
$objPHPExcel->getActiveSheet()->setCellValue('G10', "=($is_treasury3-T10)");
$objPHPExcel->getActiveSheet()->setCellValue('H10', "=+F10-G10");
$objPHPExcel->getActiveSheet()->setCellValue('M10', $is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('L10', $budget_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('Q10', $budget_is_treasury2);
$objPHPExcel->getActiveSheet()->setCellValue('J10', "=(F10-L10)");
$objPHPExcel->getActiveSheet()->setCellValue('N10', "=(M10-Q10)");
$objPHPExcel->getActiveSheet()->setCellValue('K10', "=+IF(J10=0,0,(J10/L10))");
$objPHPExcel->getActiveSheet()->setCellValue('O10', "=+IF(N10=0,0,(N10/Q10))");
$objPHPExcel->getActiveSheet()->setCellValue('S10', "=+IF(M10=0,0,IF(R10=0,0,(M10/R10)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T10', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T10', $is_treasury4);

$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Interbank placements');

$objPHPExcel->getActiveSheet()->setCellValue('C11', $is_interbank_placement4-$is_interbank_placement_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D11', "=(F11-C11)");
//$objPHPExcel->getActiveSheet()->setCellValue('E11', "=(D11/C11)");
$objPHPExcel->getActiveSheet()->setCellValue('E11', "=+IF(D11=0,0,IF(C11=0,0,(D11/C11)))");
$objPHPExcel->getActiveSheet()->setCellValue('F11', "=(M11-T11)");
$objPHPExcel->getActiveSheet()->setCellValue('G11', $is_interbank_placement3-$is_interbank_placement4);
$objPHPExcel->getActiveSheet()->setCellValue('H11', "=+F11-G11");
$objPHPExcel->getActiveSheet()->setCellValue('M11', $is_interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('L11', $budget_is_interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('Q11', $budget_is_interbank_placement2);
$objPHPExcel->getActiveSheet()->setCellValue('J11', "=(F11-L11)");
$objPHPExcel->getActiveSheet()->setCellValue('N11', "=(M11-Q11)");
$objPHPExcel->getActiveSheet()->setCellValue('K11', "=+IF(J11=0,0,(J11/L11))");
$objPHPExcel->getActiveSheet()->setCellValue('O11', "=+IF(N11=0,0,(N11/Q11))");
$objPHPExcel->getActiveSheet()->setCellValue('S11', "=+IF(M11=0,0,IF(R11=0,0,(M11/R11)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T11', $acc_is_interbank_placement);
$objPHPExcel->getActiveSheet()->setCellValue('T11', $is_interbank_placement4);

$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Placement with BI');
$objPHPExcel->getActiveSheet()->setCellValue('C12', $is_placement_wbi4-$is_placement_wbi_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D12', "=(F12-C12)");
$objPHPExcel->getActiveSheet()->setCellValue('E12', "=+IF(D12=0,0,IF(C12=0,0,(D12/C12)))");
$objPHPExcel->getActiveSheet()->setCellValue('F12', "=(M12-T12)");
$objPHPExcel->getActiveSheet()->setCellValue('G12', $is_placement_wbi3-$is_placement_wbi4);
$objPHPExcel->getActiveSheet()->setCellValue('H12', "=+F12-G12");
$objPHPExcel->getActiveSheet()->setCellValue('M12', $is_placement_wbi);
$objPHPExcel->getActiveSheet()->setCellValue('L12', $budget_is_placement_wbi);
$objPHPExcel->getActiveSheet()->setCellValue('Q12', $budget_is_placement_wbi2);
$objPHPExcel->getActiveSheet()->setCellValue('J12', "=(F12-L12)");
$objPHPExcel->getActiveSheet()->setCellValue('N12', "=(M12-Q12)");
$objPHPExcel->getActiveSheet()->setCellValue('K12', "=+IF(J12=0,0,(J12/L21))");
$objPHPExcel->getActiveSheet()->setCellValue('O12', "=+IF(N12=0,0,(N12/Q12))");
$objPHPExcel->getActiveSheet()->setCellValue('S21', "=+IF(M12=0,0,IF(R12=0,0,(M12/R12)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T12', $acc_is_placement_wbi);
$objPHPExcel->getActiveSheet()->setCellValue('T12', $is_placement_wbi4);

$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('C13', $is_ii_others_assets4-$is_ii_others_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D13', "=(F13-C13)");
$objPHPExcel->getActiveSheet()->setCellValue('E13', "=+IF(D13=0,0,IF(C13=0,0,(D13/C13)))");
$objPHPExcel->getActiveSheet()->setCellValue('F13', "=(M13-T13)");
$objPHPExcel->getActiveSheet()->setCellValue('G13', $is_ii_others_assets3-$is_ii_others_assets4);
$objPHPExcel->getActiveSheet()->setCellValue('H13', "=+F13-G13");
$objPHPExcel->getActiveSheet()->setCellValue('M13', $is_ii_others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('L13', $budget_is_ii_others);
$objPHPExcel->getActiveSheet()->setCellValue('Q13', $budget_is_ii_others2);
$objPHPExcel->getActiveSheet()->setCellValue('J13', "=(F13-L13)");
$objPHPExcel->getActiveSheet()->setCellValue('N13', "=(M13-Q13)");
$objPHPExcel->getActiveSheet()->setCellValue('K13', "=+IF(J13=0,0,(J13/L13))");
$objPHPExcel->getActiveSheet()->setCellValue('O13', "=+IF(N13=0,0,(N13/Q13))");
$objPHPExcel->getActiveSheet()->setCellValue('S13', "=+IF(M13=0,0,IF(R13=0,0,(M13/R13)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T13', $acc_is_ii_others);
$objPHPExcel->getActiveSheet()->setCellValue('T13', $is_ii_others_assets4);



$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Interest Expense Funding');
$objPHPExcel->getActiveSheet()->setCellValue('C14', "=SUM(C15:C18)");
$objPHPExcel->getActiveSheet()->setCellValue('D14', "=SUM(D15:D18)");
$objPHPExcel->getActiveSheet()->setCellValue('E14', "=+IF(D14=0,0,IF(C14=0,0,(D14/C14)))");
$objPHPExcel->getActiveSheet()->setCellValue('F14', "=SUM(F15:F18)");
$objPHPExcel->getActiveSheet()->setCellValue('G14', "=SUM(G15:G18)");
$objPHPExcel->getActiveSheet()->setCellValue('H14', "=SUM(H15:H18)");
$objPHPExcel->getActiveSheet()->setCellValue('M14', "=SUM(M15:M18)");
$objPHPExcel->getActiveSheet()->setCellValue('L14', "=SUM(L15:L18)");
$objPHPExcel->getActiveSheet()->setCellValue('Q14', "=SUM(Q15:Q18)");
$objPHPExcel->getActiveSheet()->setCellValue('J14', "=SUM(J15:J18)");
$objPHPExcel->getActiveSheet()->setCellValue('N14', "=SUM(N15:N18)");
$objPHPExcel->getActiveSheet()->setCellValue('K14', "=+IF(J9=14,0,(J14/L14))");
$objPHPExcel->getActiveSheet()->setCellValue('O14', "=+IF(N8=14,0,(N14/Q14))");
$objPHPExcel->getActiveSheet()->setCellValue('S14', "=+IF(M8=14,0,IF(R14=0,0,(M14/R14)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T14', "=SUM(T15:T18)");
$objPHPExcel->getActiveSheet()->setCellValue('T14', "=C14");

$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Current accounts');
$objPHPExcel->getActiveSheet()->setCellValue('C15', $is_current_account4-$is_current_account_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D15', "=(F15-C15)");
$objPHPExcel->getActiveSheet()->setCellValue('E15', "=+IF(D15=0,0,IF(C15=0,0,(D15/C15)))");
$objPHPExcel->getActiveSheet()->setCellValue('F15', "=(M15-T15)");
$objPHPExcel->getActiveSheet()->setCellValue('G15', $is_current_account3-$is_current_account4);
$objPHPExcel->getActiveSheet()->setCellValue('H15', "=+F15-G15");
$objPHPExcel->getActiveSheet()->setCellValue('M15', $is_current_account);
$objPHPExcel->getActiveSheet()->setCellValue('L15', $budget_is_current_account);
$objPHPExcel->getActiveSheet()->setCellValue('Q15', $budget_is_current_account2);
$objPHPExcel->getActiveSheet()->setCellValue('J15', "=(F15-L15)");
$objPHPExcel->getActiveSheet()->setCellValue('N15', "=(M15-Q15)");
$objPHPExcel->getActiveSheet()->setCellValue('K15', "=+IF(J15=0,0,(J15/L15))");
$objPHPExcel->getActiveSheet()->setCellValue('O15', "=+IF(N15=0,0,(N15/Q15))");
$objPHPExcel->getActiveSheet()->setCellValue('S15', "=+IF(M15=0,0,IF(R15=0,0,(M15/R15)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T15', $acc_is_current_account);
$objPHPExcel->getActiveSheet()->setCellValue('T15', $is_current_account4);

$objPHPExcel->getActiveSheet()->setCellValue('B16', 'Saving accounts');
$objPHPExcel->getActiveSheet()->setCellValue('C16', $is_saving_account4-$is_saving_account_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D16', "=(F16-C16)");
$objPHPExcel->getActiveSheet()->setCellValue('E16', "=+IF(D16=0,0,IF(C16=0,0,(D16/C16)))");
$objPHPExcel->getActiveSheet()->setCellValue('F16', "=(M16-T16)");
$objPHPExcel->getActiveSheet()->setCellValue('G16', $is_saving_account3-$is_saving_account4);
$objPHPExcel->getActiveSheet()->setCellValue('H16', "=+F16-G16");
$objPHPExcel->getActiveSheet()->setCellValue('M16', $is_saving_account);
$objPHPExcel->getActiveSheet()->setCellValue('L16', $budget_is_saving_account);
$objPHPExcel->getActiveSheet()->setCellValue('Q16', $budget_is_saving_account2);
$objPHPExcel->getActiveSheet()->setCellValue('J16', "=(F16-L16)");
$objPHPExcel->getActiveSheet()->setCellValue('N16', "=(M16-Q16)");
$objPHPExcel->getActiveSheet()->setCellValue('K16', "=+IF(J16=0,0,(J16/L16))");
$objPHPExcel->getActiveSheet()->setCellValue('O16', "=+IF(N16=0,0,(N16/Q16))");
$objPHPExcel->getActiveSheet()->setCellValue('S16', "=+IF(M16=0,0,IF(R16=0,0,(M16/R16)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T16', $acc_is_saving_account);
$objPHPExcel->getActiveSheet()->setCellValue('T16', $is_saving_account4);

$objPHPExcel->getActiveSheet()->setCellValue('B17', 'Time deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C17', $is_time_deposits4-$is_time_deposits_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D17', "=(F17-C17)");
$objPHPExcel->getActiveSheet()->setCellValue('E17', "=+IF(D17=0,0,IF(C17=0,0,(D17/C17)))");
$objPHPExcel->getActiveSheet()->setCellValue('F17', "=(M17-T17)");
$objPHPExcel->getActiveSheet()->setCellValue('G17', $is_time_deposits3-$is_time_deposits4);
$objPHPExcel->getActiveSheet()->setCellValue('H17', "=+F17-G17");
$objPHPExcel->getActiveSheet()->setCellValue('M17', $is_time_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('L17', $budget_is_time_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('Q17', $budget_is_time_deposits2);
$objPHPExcel->getActiveSheet()->setCellValue('J17', "=(F17-L17)");
$objPHPExcel->getActiveSheet()->setCellValue('N17', "=(M17-Q17)");
$objPHPExcel->getActiveSheet()->setCellValue('K17', "=+IF(J17=0,0,(J17/L17))");
$objPHPExcel->getActiveSheet()->setCellValue('O17', "=+IF(N17=0,0,(N17/Q17))");
$objPHPExcel->getActiveSheet()->setCellValue('S17', "=+IF(M17=0,0,IF(R17=0,0,(M17/R17)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T17', $acc_is_time_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('T17', $is_time_deposits4);

$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Bank deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C18', $is_bank_deposit4-$is_bank_deposit_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D18', "=(F18-C18)");
$objPHPExcel->getActiveSheet()->setCellValue('E18', "=+IF(D18=0,0,IF(C18=0,0,(D18/C18)))");
$objPHPExcel->getActiveSheet()->setCellValue('F18', "=(M18-T18)");
$objPHPExcel->getActiveSheet()->setCellValue('G18', $is_bank_deposit3-$is_bank_deposit4);
$objPHPExcel->getActiveSheet()->setCellValue('H18', "=+F18-G18");
$objPHPExcel->getActiveSheet()->setCellValue('M18', $is_bank_deposit);
$objPHPExcel->getActiveSheet()->setCellValue('L18', $budget_is_bank_deposit);
$objPHPExcel->getActiveSheet()->setCellValue('Q18', $budget_is_bank_deposit2);
$objPHPExcel->getActiveSheet()->setCellValue('J18', "=(F18-L18)");
$objPHPExcel->getActiveSheet()->setCellValue('N18', "=(M18-Q18)");
$objPHPExcel->getActiveSheet()->setCellValue('K18', "=+IF(J18=0,0,(J18/L18))");
$objPHPExcel->getActiveSheet()->setCellValue('O18', "=+IF(N18=0,0,(N18/Q18))");
$objPHPExcel->getActiveSheet()->setCellValue('S18', "=+IF(M18=0,0,IF(R18=0,0,(M18/R18)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T18', $acc_is_bank_deposit);
$objPHPExcel->getActiveSheet()->setCellValue('T18', $is_bank_deposit4);

$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Other Interest Expense ');
$objPHPExcel->getActiveSheet()->setCellValue('C19', "=SUM(C20:C22)");
$objPHPExcel->getActiveSheet()->setCellValue('D19', "=SUM(D20:D22)");
$objPHPExcel->getActiveSheet()->setCellValue('E19', "=+IF(D19=0,0,IF(C19=0,0,(D19/C19)))");
$objPHPExcel->getActiveSheet()->setCellValue('F19', "=SUM(F20:F22)");
$objPHPExcel->getActiveSheet()->setCellValue('G19', "=SUM(G20:G22)");
$objPHPExcel->getActiveSheet()->setCellValue('H19', "=SUM(H20:H22)");
$objPHPExcel->getActiveSheet()->setCellValue('M19', "=SUM(M20:M22)");
$objPHPExcel->getActiveSheet()->setCellValue('L19', "=SUM(L20:L22)");
$objPHPExcel->getActiveSheet()->setCellValue('Q19', "=SUM(Q20:Q22)");
$objPHPExcel->getActiveSheet()->setCellValue('J19', "=SUM(J20:J22)");
$objPHPExcel->getActiveSheet()->setCellValue('N19', "=SUM(N20:N22)");
$objPHPExcel->getActiveSheet()->setCellValue('K19', "=+IF(J19=0,0,(J19/L19))");
$objPHPExcel->getActiveSheet()->setCellValue('O19', "=+IF(N19=0,0,(N19/Q19))");
$objPHPExcel->getActiveSheet()->setCellValue('S19', "=+IF(M19=0,0,IF(R19=0,0,(M19/R19)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T19', "=SUM(T20:T22)");
$objPHPExcel->getActiveSheet()->setCellValue('T19', "=C19");
//$objPHPExcel->getActiveSheet()->setCellValue('');

$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Borrowings (MCB)');
$objPHPExcel->getActiveSheet()->setCellValue('C20', $is_borrowings4-$is_borrowings_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D20', "=(F20-C20)");
$objPHPExcel->getActiveSheet()->setCellValue('E20', "=+IF(D20=0,0,IF(C20=0,0,(D20/C20)))");
$objPHPExcel->getActiveSheet()->setCellValue('F20', "=(M20-T20)");
$objPHPExcel->getActiveSheet()->setCellValue('G20', $is_borrowings3-$is_borrowings4);
$objPHPExcel->getActiveSheet()->setCellValue('H20', "=+F20-G20");
$objPHPExcel->getActiveSheet()->setCellValue('M20', $is_borrowings);
$objPHPExcel->getActiveSheet()->setCellValue('L20', $budget_is_borrowings);
$objPHPExcel->getActiveSheet()->setCellValue('Q20', $budget_is_borrowings2);
$objPHPExcel->getActiveSheet()->setCellValue('J20', "=(F20-L20)");
$objPHPExcel->getActiveSheet()->setCellValue('N20', "=(M20-Q20)");
$objPHPExcel->getActiveSheet()->setCellValue('K20', "=+IF(J20=0,0,(J20/L20))");
$objPHPExcel->getActiveSheet()->setCellValue('O20', "=+IF(N20=0,0,(N20/Q20))");
$objPHPExcel->getActiveSheet()->setCellValue('S20', "=+IF(M20=0,0,IF(R20=0,0,(M20/R20)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T20', $acc_is_borrowings);
$objPHPExcel->getActiveSheet()->setCellValue('T20', $is_borrowings4);

$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Guaranteed premium');
$objPHPExcel->getActiveSheet()->setCellValue('C21', $is_guaranteed4-$is_guaranteed_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D21', "=(F21-C21)");
$objPHPExcel->getActiveSheet()->setCellValue('E21', "=+IF(D21=0,0,IF(C21=0,0,(D21/C21)))");
$objPHPExcel->getActiveSheet()->setCellValue('F21', "=(M21-T21)");
$objPHPExcel->getActiveSheet()->setCellValue('G21', $is_guaranteed3-$is_guaranteed4);
$objPHPExcel->getActiveSheet()->setCellValue('H21', "=+F21-G21");
$objPHPExcel->getActiveSheet()->setCellValue('M21', $is_guaranteed);
$objPHPExcel->getActiveSheet()->setCellValue('L21', $budget_is_guaranteed);
$objPHPExcel->getActiveSheet()->setCellValue('Q21', $budget_is_guaranteed2);
$objPHPExcel->getActiveSheet()->setCellValue('J21', "=(F21-L21)");
$objPHPExcel->getActiveSheet()->setCellValue('N21', "=(M21-Q21)");
$objPHPExcel->getActiveSheet()->setCellValue('K21', "=+IF(J21=0,0,(J21/L21))");
$objPHPExcel->getActiveSheet()->setCellValue('O21', "=+IF(N21=0,0,(N21/Q21))");
$objPHPExcel->getActiveSheet()->setCellValue('S21', "=+IF(M21=0,0,IF(R21=0,0,(M21/21)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T21', $acc_is_guaranteed);
$objPHPExcel->getActiveSheet()->setCellValue('T21', $is_guaranteed4);

$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('C22', $is_ie_others_assets4-$is_ie_others_assets_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D22', "=(F22-C22)");
$objPHPExcel->getActiveSheet()->setCellValue('E22', "=+IF(D22=0,0,IF(C22=0,0,(D22/C22)))");
$objPHPExcel->getActiveSheet()->setCellValue('F22', "=(M22-T22)");
$objPHPExcel->getActiveSheet()->setCellValue('G22', $is_ie_others_assets3-$is_ie_others_assets4);
$objPHPExcel->getActiveSheet()->setCellValue('H22', "=+F22-G22");
$objPHPExcel->getActiveSheet()->setCellValue('M22', $is_ie_others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('L22', $budget_is_ie_others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('Q22', $budget_is_ie_others_assets2);
$objPHPExcel->getActiveSheet()->setCellValue('J22', "=(F22-L22)");
$objPHPExcel->getActiveSheet()->setCellValue('N22', "=(M22-Q22)");
$objPHPExcel->getActiveSheet()->setCellValue('K22', "=+IF(J22=0,0,(J22/L22))");
$objPHPExcel->getActiveSheet()->setCellValue('O22', "=+IF(N22=0,0,(N22/Q22))");
$objPHPExcel->getActiveSheet()->setCellValue('S22', "=+IF(M22=0,0,IF(R22=0,0,(M22/R22)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T22', $acc_is_ie_others_assets);
$objPHPExcel->getActiveSheet()->setCellValue('T22', $is_ie_others_assets4);

$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Net Interest Income');
$objPHPExcel->getActiveSheet()->setCellValue('C23', "=SUM(C8+C14+C19)");
$objPHPExcel->getActiveSheet()->setCellValue('D23', "=SUM(D8+D14+D19)");
$objPHPExcel->getActiveSheet()->setCellValue('E23', "=+IF(D23=0,0,IF(C23=0,0,(D23/C23)))");
$objPHPExcel->getActiveSheet()->setCellValue('F23', "=SUM(F8+F14+F19)");
$objPHPExcel->getActiveSheet()->setCellValue('G23', "=SUM(G8+G14+G19)");
$objPHPExcel->getActiveSheet()->setCellValue('H23', "=SUM(H8+H14+H19)");
$objPHPExcel->getActiveSheet()->setCellValue('M23', "=SUM(M8+M14+M19)");
$objPHPExcel->getActiveSheet()->setCellValue('L23', "=SUM(L8+L14+L19)");
$objPHPExcel->getActiveSheet()->setCellValue('Q23', "=SUM(Q8+Q14+Q19)");
$objPHPExcel->getActiveSheet()->setCellValue('J23', "=SUM(J8+J14+J19)");
$objPHPExcel->getActiveSheet()->setCellValue('N23', "=SUM(N8+N14+N19)");
$objPHPExcel->getActiveSheet()->setCellValue('K23', "=+IF(J23=0,0,(J23/L23))");
$objPHPExcel->getActiveSheet()->setCellValue('O23', "=+IF(N23=0,0,(N23/Q23))");
$objPHPExcel->getActiveSheet()->setCellValue('S23', "=+IF(M23=0,0,IF(R23=0,0,(M23/R23)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T23', "=SUM(T8+T14+T19)");
$objPHPExcel->getActiveSheet()->setCellValue('T23', "=C23");

$objPHPExcel->getActiveSheet()->setCellValue('B24', '');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'Other Income:');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'Forex gain/(loss) on transactions');
$objPHPExcel->getActiveSheet()->setCellValue('C26', $forex_gain4-$forex_gain_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D26', "=(F26-C26)");
$objPHPExcel->getActiveSheet()->setCellValue('E26', "=+IF(D26=0,0,IF(C26=0,0,(D26/C26)))");
$objPHPExcel->getActiveSheet()->setCellValue('F26', "=(M26-T26)");
$objPHPExcel->getActiveSheet()->setCellValue('G26', $forex_gain3-$forex_gain4);
$objPHPExcel->getActiveSheet()->setCellValue('H26', "=+F26-G26");
$objPHPExcel->getActiveSheet()->setCellValue('M26', $forex_gain);
$objPHPExcel->getActiveSheet()->setCellValue('L26', $budget_forex_gain);
$objPHPExcel->getActiveSheet()->setCellValue('Q26', $budget_forex_gain2);
$objPHPExcel->getActiveSheet()->setCellValue('J26', "=(F26-L26)");
$objPHPExcel->getActiveSheet()->setCellValue('N26', "=(M26-Q26)");
$objPHPExcel->getActiveSheet()->setCellValue('K26', "=+IF(J26=0,0,(J26/L26))");
$objPHPExcel->getActiveSheet()->setCellValue('O26', "=+IF(N26=0,0,(N26/Q26))");
$objPHPExcel->getActiveSheet()->setCellValue('S26', "=+IF(M26=0,0,IF(R26=0,0,(M26/R26)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T26', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T26', $forex_gain4);

$objPHPExcel->getActiveSheet()->setCellValue('B27', 'Gain/(Loss) on sale of securities/bonds');
$objPHPExcel->getActiveSheet()->setCellValue('C27', $gain_loss4-$gain_loss_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D27', "=(F27-C27)");
$objPHPExcel->getActiveSheet()->setCellValue('E27', "=+IF(D27=0,0,IF(C27=0,0,(D27/C27)))");
$objPHPExcel->getActiveSheet()->setCellValue('F27', "=(M27-T27)");
$objPHPExcel->getActiveSheet()->setCellValue('G27', $gain_loss3-$gain_loss4);
$objPHPExcel->getActiveSheet()->setCellValue('H27', "=+F27-G27");
$objPHPExcel->getActiveSheet()->setCellValue('M27', $gain_loss);
$objPHPExcel->getActiveSheet()->setCellValue('L27', $budget_gain_loss);
$objPHPExcel->getActiveSheet()->setCellValue('Q27', $budget_gain_loss2);
$objPHPExcel->getActiveSheet()->setCellValue('J27', "=(F27-L27)");
$objPHPExcel->getActiveSheet()->setCellValue('N27', "=(M27-Q27)");
$objPHPExcel->getActiveSheet()->setCellValue('K27', "=+IF(J27=0,0,(J27/L27))");
$objPHPExcel->getActiveSheet()->setCellValue('O27', "=+IF(N27=0,0,(N27/Q27))");
$objPHPExcel->getActiveSheet()->setCellValue('S27', "=+IF(M27=0,0,IF(R27=0,0,(M27/R27)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T27', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T27', $gain_loss4);

$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Remittance fee');
$objPHPExcel->getActiveSheet()->setCellValue('C28', $remittance_fee4-$remittance_fee_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D28', "=(F28-C28)");
$objPHPExcel->getActiveSheet()->setCellValue('E28', "=+IF(D28=0,0,IF(C28=0,0,(D28/C28)))");
$objPHPExcel->getActiveSheet()->setCellValue('F28', "=(M28-T28)");
$objPHPExcel->getActiveSheet()->setCellValue('G28', $remittance_fee3-$remittance_fee4);
$objPHPExcel->getActiveSheet()->setCellValue('H28', "=+F28-G128");
$objPHPExcel->getActiveSheet()->setCellValue('M28', $remittance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('L28', $budget_remittance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('Q28', $budget_remittance_fee2);
$objPHPExcel->getActiveSheet()->setCellValue('J28', "=(F28-L28)");
$objPHPExcel->getActiveSheet()->setCellValue('N28', "=(M28-Q28)");
$objPHPExcel->getActiveSheet()->setCellValue('K28', "=+IF(J28=0,0,(J28/L28))");
$objPHPExcel->getActiveSheet()->setCellValue('O28', "=+IF(N28=0,0,(N28/Q28))");
$objPHPExcel->getActiveSheet()->setCellValue('S28', "=+IF(M28=0,0,IF(R28=0,0,(M28/R28)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T28', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T28', $remittance_fee4);

$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Trade Finance fee');
$objPHPExcel->getActiveSheet()->setCellValue('C29', $trade_finance_fee4-$trade_finance_fee_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D29', "=(F29-C29)");
$objPHPExcel->getActiveSheet()->setCellValue('E29', "=+IF(D29=0,0,IF(C29=0,0,(D29/C29)))");
$objPHPExcel->getActiveSheet()->setCellValue('F29', "=(M29-T29)");
$objPHPExcel->getActiveSheet()->setCellValue('G29', $trade_finance_fee3-$trade_finance_fee4);
$objPHPExcel->getActiveSheet()->setCellValue('H29', "=+F29-G29");
$objPHPExcel->getActiveSheet()->setCellValue('M29', $trade_finance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('L29', $budget_trade_finance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('Q29', $budget_trade_finance_fee2);
$objPHPExcel->getActiveSheet()->setCellValue('J29', "=(F29-L29)");
$objPHPExcel->getActiveSheet()->setCellValue('N29', "=(M29-Q29)");
$objPHPExcel->getActiveSheet()->setCellValue('K29', "=+IF(J29=0,0,(J29/L29))");
$objPHPExcel->getActiveSheet()->setCellValue('O29', "=+IF(N29=0,0,(N29/Q29))");
$objPHPExcel->getActiveSheet()->setCellValue('S29', "=+IF(M29=0,0,IF(R29=0,0,(M29/R29)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T29', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T29', $trade_finance_fee4);

$objPHPExcel->getActiveSheet()->setCellValue('B30', 'Processing fee');
$objPHPExcel->getActiveSheet()->setCellValue('C30', $processing_fee4-$processing_fee_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D30', "=(F30-C30)");
$objPHPExcel->getActiveSheet()->setCellValue('E30', "=+IF(D30=0,0,IF(C30=0,0,(D30/C30)))");
$objPHPExcel->getActiveSheet()->setCellValue('F30', "=(M30-T30)");
$objPHPExcel->getActiveSheet()->setCellValue('G30', $processing_fee3-$processing_fee4);
$objPHPExcel->getActiveSheet()->setCellValue('H30', "=+F30-G30");
$objPHPExcel->getActiveSheet()->setCellValue('M30', $processing_fee);
$objPHPExcel->getActiveSheet()->setCellValue('L30', $budget_processing_fee);
$objPHPExcel->getActiveSheet()->setCellValue('Q30', $budget_processing_fee2);
$objPHPExcel->getActiveSheet()->setCellValue('J30', "=(F30-L30)");
$objPHPExcel->getActiveSheet()->setCellValue('N30', "=(M30-Q30)");
$objPHPExcel->getActiveSheet()->setCellValue('K30', "=+IF(J30=0,0,(J30/L30))");
$objPHPExcel->getActiveSheet()->setCellValue('O30', "=+IF(N30=0,0,(N30/Q30))");
$objPHPExcel->getActiveSheet()->setCellValue('S30', "=+IF(M30=0,0,IF(R30=0,0,(M30/R30)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T30', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T30', $processing_fee4);

$objPHPExcel->getActiveSheet()->setCellValue('B31', 'Credit Card fee');
$objPHPExcel->getActiveSheet()->setCellValue('C31', $credit_card_fee4-$credit_card_fee_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D31', "=(F31-C31)");
$objPHPExcel->getActiveSheet()->setCellValue('E31', "=+IF(D31=0,0,IF(C31=0,0,(D31/C31)))");
$objPHPExcel->getActiveSheet()->setCellValue('F31', "=(M31-T31)");
$objPHPExcel->getActiveSheet()->setCellValue('G31', $credit_card_fee3-$credit_card_fee4);
$objPHPExcel->getActiveSheet()->setCellValue('H31', "=+F31-G31");
$objPHPExcel->getActiveSheet()->setCellValue('M31', $credit_card_fee);
$objPHPExcel->getActiveSheet()->setCellValue('L31', $budget_credit_card_fee);
$objPHPExcel->getActiveSheet()->setCellValue('Q31', $budget_credit_card_fee2);
$objPHPExcel->getActiveSheet()->setCellValue('J31', "=(F31-L31)");
$objPHPExcel->getActiveSheet()->setCellValue('N31', "=(M31-Q31)");
$objPHPExcel->getActiveSheet()->setCellValue('K31', "=+IF(J31=0,0,(J31/L31))");
$objPHPExcel->getActiveSheet()->setCellValue('O31', "=+IF(N31=0,0,(N31/Q31))");
$objPHPExcel->getActiveSheet()->setCellValue('S31', "=+IF(M31=0,0,IF(R31=0,0,(M31/R31)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T31', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T31', $credit_card_fee4);

$objPHPExcel->getActiveSheet()->setCellValue('B32', 'Insurance Fee');
$objPHPExcel->getActiveSheet()->setCellValue('C32', $insurance_fee4-$insurance_fee_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D32', "=(F32-C32)");
$objPHPExcel->getActiveSheet()->setCellValue('E32', "=+IF(D32=0,0,IF(C32=0,0,(D32/C32)))");
$objPHPExcel->getActiveSheet()->setCellValue('F32', "=(M32-T32)");
$objPHPExcel->getActiveSheet()->setCellValue('G32', $insurance_fee3-$insurance_fee4);
$objPHPExcel->getActiveSheet()->setCellValue('H32', "=+F32-G32");
$objPHPExcel->getActiveSheet()->setCellValue('M32', $insurance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('L32', $budget_insurance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('Q32', $budget_insurance_fee2);
$objPHPExcel->getActiveSheet()->setCellValue('J32', "=(F32-L32)");
$objPHPExcel->getActiveSheet()->setCellValue('N32', "=(M32-Q32)");
$objPHPExcel->getActiveSheet()->setCellValue('K32', "=+IF(J32=0,0,(J32/L32))");
$objPHPExcel->getActiveSheet()->setCellValue('O32', "=+IF(N32=0,0,(N32/Q32))");
$objPHPExcel->getActiveSheet()->setCellValue('S32', "=+IF(M32=0,0,IF(R32=0,0,(M32/R32)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T32', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T32', $insurance_fee4);

$objPHPExcel->getActiveSheet()->setCellValue('B33', 'Service Charges');
$objPHPExcel->getActiveSheet()->setCellValue('C33', $service_charges4-$service_charges_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D33', "=(F33-C33)");
$objPHPExcel->getActiveSheet()->setCellValue('E33', "=+IF(D33=0,0,IF(C33=0,0,(D33/C33)))");
$objPHPExcel->getActiveSheet()->setCellValue('F33', "=(M33-T33)");
$objPHPExcel->getActiveSheet()->setCellValue('G33', $service_charges3-$service_charges4);
$objPHPExcel->getActiveSheet()->setCellValue('H33', "=+F33-G33");
$objPHPExcel->getActiveSheet()->setCellValue('M33', $service_charges);
$objPHPExcel->getActiveSheet()->setCellValue('L33', $budget_service_charges);
$objPHPExcel->getActiveSheet()->setCellValue('Q33', $budget_service_charges2);
$objPHPExcel->getActiveSheet()->setCellValue('J33', "=(F33-L33)");
$objPHPExcel->getActiveSheet()->setCellValue('N33', "=(M33-Q33)");
$objPHPExcel->getActiveSheet()->setCellValue('K33', "=+IF(J33=0,0,(J33/L33))");
$objPHPExcel->getActiveSheet()->setCellValue('O33', "=+IF(N33=0,0,(N33/Q33))");
$objPHPExcel->getActiveSheet()->setCellValue('S33', "=+IF(M33=0,0,IF(R33=0,0,(M33/R33)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T33', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T33', $service_charges4);

$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Other Commission & Fee ');
$objPHPExcel->getActiveSheet()->setCellValue('C34', $other_cf4-$other_cf_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D34', "=(F34-C34)");
$objPHPExcel->getActiveSheet()->setCellValue('E34', "=+IF(D34=0,0,IF(C34=0,0,(D34/C34)))");
$objPHPExcel->getActiveSheet()->setCellValue('F34', "=(M34-T34)");
$objPHPExcel->getActiveSheet()->setCellValue('G34', $other_cf3-$other_cf4);
$objPHPExcel->getActiveSheet()->setCellValue('H34', "=+F34-G34");
$objPHPExcel->getActiveSheet()->setCellValue('M34', $other_cf);
$objPHPExcel->getActiveSheet()->setCellValue('L34', $budget_other_cf);
$objPHPExcel->getActiveSheet()->setCellValue('Q34', $budget_other_cf2);
$objPHPExcel->getActiveSheet()->setCellValue('J34', "=(F34-L34)");
$objPHPExcel->getActiveSheet()->setCellValue('N34', "=(M34-Q34)");
$objPHPExcel->getActiveSheet()->setCellValue('K34', "=+IF(J34=0,0,(J34/L34))");
$objPHPExcel->getActiveSheet()->setCellValue('O34', "=+IF(N34=0,0,(N34/Q34))");
$objPHPExcel->getActiveSheet()->setCellValue('S34', "=+IF(M34=0,0,IF(R34=0,0,(M34/R34)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T34', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T34', $other_cf4);


$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Penalty');
$objPHPExcel->getActiveSheet()->setCellValue('C35', $penalty4-$penalty_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D35', "=(F35-C35)");
$objPHPExcel->getActiveSheet()->setCellValue('E35', "=+IF(D35=0,0,IF(C35=0,0,(D35/C35)))");
$objPHPExcel->getActiveSheet()->setCellValue('F35', "=(M35-T35)");
$objPHPExcel->getActiveSheet()->setCellValue('G35', $penalty3-$penalty4);
$objPHPExcel->getActiveSheet()->setCellValue('H35', "=+F35-G35");
$objPHPExcel->getActiveSheet()->setCellValue('M35', $penalty);
$objPHPExcel->getActiveSheet()->setCellValue('L35', $budget_penalty);
$objPHPExcel->getActiveSheet()->setCellValue('Q35', $budget_penalty2);
$objPHPExcel->getActiveSheet()->setCellValue('J35', "=(F35-L35)");
$objPHPExcel->getActiveSheet()->setCellValue('N35', "=(M35-Q35)");
$objPHPExcel->getActiveSheet()->setCellValue('K35', "=+IF(J35=0,0,(J35/L35))");
$objPHPExcel->getActiveSheet()->setCellValue('O35', "=+IF(N35=0,0,(N35/Q35))");
$objPHPExcel->getActiveSheet()->setCellValue('S35', "=+IF(M35=0,0,IF(R35=0,0,(M35/R35)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T35', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T35', $penalty4);

$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Write Offs Recovered');
$objPHPExcel->getActiveSheet()->setCellValue('C36', $write_orr4-$write_orr_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D36', "=(F36-C36)");
$objPHPExcel->getActiveSheet()->setCellValue('E36', "=+IF(D36=0,0,IF(C36=0,0,(D36/C36)))");
$objPHPExcel->getActiveSheet()->setCellValue('F36', "=(M36-T36)");
$objPHPExcel->getActiveSheet()->setCellValue('G36', $write_orr3-$write_orr4);
$objPHPExcel->getActiveSheet()->setCellValue('H36', "=+F36-G36");
$objPHPExcel->getActiveSheet()->setCellValue('M36', $write_orr);
$objPHPExcel->getActiveSheet()->setCellValue('L36', $budget_write_orr);
$objPHPExcel->getActiveSheet()->setCellValue('Q36', $budget_write_orr2);
$objPHPExcel->getActiveSheet()->setCellValue('J36', "=(F36-L36)");
$objPHPExcel->getActiveSheet()->setCellValue('N36', "=(M36-Q36)");
$objPHPExcel->getActiveSheet()->setCellValue('K36', "=+IF(J36=0,0,IF(L36=0,0,(J36/L36)))");
$objPHPExcel->getActiveSheet()->setCellValue('O36', "=+IF(N36=0,0,IF(Q36=0,0,(N36/Q36)))");
$objPHPExcel->getActiveSheet()->setCellValue('S36', "=+IF(M36=0,0,IF(R36=0,0,(M36/R36)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T36', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T36', $write_orr4);

$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Other Income');
$objPHPExcel->getActiveSheet()->setCellValue('C37', $other_income4-$other_income_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D37', "=(F37-C37)");
$objPHPExcel->getActiveSheet()->setCellValue('E37', "=+IF(D37=0,0,IF(C37=0,0,(D37/C37)))");
$objPHPExcel->getActiveSheet()->setCellValue('F37', "=(M37-T37)");
$objPHPExcel->getActiveSheet()->setCellValue('G37', $other_income3-$other_income4);
$objPHPExcel->getActiveSheet()->setCellValue('H37', "=+F37-G37");
$objPHPExcel->getActiveSheet()->setCellValue('M37', $other_income);
$objPHPExcel->getActiveSheet()->setCellValue('L37', $budget_other_income);
$objPHPExcel->getActiveSheet()->setCellValue('Q37', $budget_other_income2);
$objPHPExcel->getActiveSheet()->setCellValue('J37', "=(F37-L37)");
$objPHPExcel->getActiveSheet()->setCellValue('N37', "=(M37-Q37)");
$objPHPExcel->getActiveSheet()->setCellValue('K37', "=+IF(J37=0,0,IF(L37=0,0,(J37/L37)))");
$objPHPExcel->getActiveSheet()->setCellValue('O37', "=+IF(N37=0,0,IF(Q37=0,0,(N37/Q37)))");
$objPHPExcel->getActiveSheet()->setCellValue('S37', "=+IF(M37=0,0,IF(R37=0,0,(M37/R37)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T37', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T37', $other_income4);

/*
$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Total Other Income');
$objPHPExcel->getActiveSheet()->setCellValue('C38', $total_other_income4-$total_other_income_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D38', "=(F38-C38)");
$objPHPExcel->getActiveSheet()->setCellValue('E38', "=+IF(D38=0,0,IF(C38=0,0,(D38/C38)))");
$objPHPExcel->getActiveSheet()->setCellValue('F38', "=(M38-T38)");
$objPHPExcel->getActiveSheet()->setCellValue('G38', $total_other_income3-$total_other_income4);
$objPHPExcel->getActiveSheet()->setCellValue('H38', "=+F38-G38");
$objPHPExcel->getActiveSheet()->setCellValue('M38', $total_other_income);
$objPHPExcel->getActiveSheet()->setCellValue('L38', $budget_total_other_income);
$objPHPExcel->getActiveSheet()->setCellValue('Q38', $budget_total_other_income2);
$objPHPExcel->getActiveSheet()->setCellValue('J38', "=(F38-L38)");
$objPHPExcel->getActiveSheet()->setCellValue('N38', "=(M38-Q38)");
$objPHPExcel->getActiveSheet()->setCellValue('K38', "=+IF(J38=0,0,(J38/L38))");
$objPHPExcel->getActiveSheet()->setCellValue('O38', "=+IF(N38=0,0,(N38/Q38))");
$objPHPExcel->getActiveSheet()->setCellValue('S38', "=+IF(M38=0,0,IF(R38=0,0,(M38/R38)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T38', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T38', $total_other_income4);
*/
$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Total Other Income');
$objPHPExcel->getActiveSheet()->setCellValue('C38', "=SUM(C26:C37)");
$objPHPExcel->getActiveSheet()->setCellValue('D38', "=SUM(D26:D37)");
$objPHPExcel->getActiveSheet()->setCellValue('E38', "=+IF(D38=0,0,IF(C38=0,0,(D38/C38)))");
$objPHPExcel->getActiveSheet()->setCellValue('F38', "=SUM(F26:F37)");
$objPHPExcel->getActiveSheet()->setCellValue('G38', "=SUM(G26:G37)");
$objPHPExcel->getActiveSheet()->setCellValue('H38', "=SUM(H26:H37)");
$objPHPExcel->getActiveSheet()->setCellValue('M38', "=SUM(M26:M37)");
$objPHPExcel->getActiveSheet()->setCellValue('L38', "=SUM(L26:L37)");
$objPHPExcel->getActiveSheet()->setCellValue('Q38', "=SUM(Q26:Q37)");
$objPHPExcel->getActiveSheet()->setCellValue('J38', "=SUM(J26:J37)");
$objPHPExcel->getActiveSheet()->setCellValue('N38', "=SUM(N26:N37)");
$objPHPExcel->getActiveSheet()->setCellValue('K38', "=+IF(J38=0,0,(J38/L38))");
$objPHPExcel->getActiveSheet()->setCellValue('O38', "=SUM(O26:O37)");
$objPHPExcel->getActiveSheet()->setCellValue('S38', "=+IF(M38=0,0,IF(R38=0,0,(M38/R38)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T38', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T38', "=SUM(T26:T37)");



$objPHPExcel->getActiveSheet()->setCellValue('B39', 'Operating Revenue');
$objPHPExcel->getActiveSheet()->setCellValue('C39', "=SUM(C23:C38)");
$objPHPExcel->getActiveSheet()->setCellValue('D39', "=SUM(D23:D38)");
$objPHPExcel->getActiveSheet()->setCellValue('E39', "=+IF(D39=0,0,IF(C39=0,0,(D39/C39)))");
$objPHPExcel->getActiveSheet()->setCellValue('F39', "=SUM(F23:F38)");
$objPHPExcel->getActiveSheet()->setCellValue('G39', "=SUM(G23:G38)");
$objPHPExcel->getActiveSheet()->setCellValue('H39', "=SUM(H23:H38)");
$objPHPExcel->getActiveSheet()->setCellValue('M39', "=SUM(M23:M38)");
$objPHPExcel->getActiveSheet()->setCellValue('L39', "=SUM(L23:L38)");
$objPHPExcel->getActiveSheet()->setCellValue('Q39', "=SUM(Q23:Q38)");
$objPHPExcel->getActiveSheet()->setCellValue('J39', "=(F39-L39)");
$objPHPExcel->getActiveSheet()->setCellValue('N39', "=(M39-Q39)");
$objPHPExcel->getActiveSheet()->setCellValue('K39', "=+IF(J39=0,0,(J39/L39))");
$objPHPExcel->getActiveSheet()->setCellValue('O39', "=+IF(N39=0,0,(N39/Q39))");
$objPHPExcel->getActiveSheet()->setCellValue('S39', "=+IF(M39=0,0,IF(R39=0,0,(M39/R39)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T39', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T39', "=C39");

$objPHPExcel->getActiveSheet()->setCellValue('B40', '');
$objPHPExcel->getActiveSheet()->setCellValue('B41', 'Operating Expenses:');
$objPHPExcel->getActiveSheet()->setCellValue('B42', 'Staff Cost');
$objPHPExcel->getActiveSheet()->setCellValue('C42', $staff_cost4-$staff_cost_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D42', "=(F42-C42)");
$objPHPExcel->getActiveSheet()->setCellValue('E42', "=+IF(D42=0,0,IF(C42=0,0,(D42/C42)))");
$objPHPExcel->getActiveSheet()->setCellValue('F42', "=(M42-T42)");
$objPHPExcel->getActiveSheet()->setCellValue('G42', $staff_cost3-$staff_cost4);
$objPHPExcel->getActiveSheet()->setCellValue('H42', "=+F42-G42");
$objPHPExcel->getActiveSheet()->setCellValue('M42', $staff_cost);
$objPHPExcel->getActiveSheet()->setCellValue('L42', $budget_staff_cost);
$objPHPExcel->getActiveSheet()->setCellValue('Q42', $budget_staff_cost2);
$objPHPExcel->getActiveSheet()->setCellValue('J42', "=(F42-L42)");
$objPHPExcel->getActiveSheet()->setCellValue('N42', "=(M42-Q42)");
$objPHPExcel->getActiveSheet()->setCellValue('K42', "=+IF(J42=0,0,(J42/L42))");
$objPHPExcel->getActiveSheet()->setCellValue('O42', "=+IF(N42=0,0,(N42/Q42))");
$objPHPExcel->getActiveSheet()->setCellValue('S42', "=+IF(M42=0,0,IF(R42=0,0,(M42/R42)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T42', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T42', $staff_cost4);

$objPHPExcel->getActiveSheet()->setCellValue('B43', 'General & Administrative Expenses');
$objPHPExcel->getActiveSheet()->setCellValue('C43', $general_ae4-$general_ae_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D43', "=(F43-C43)");
$objPHPExcel->getActiveSheet()->setCellValue('E43', "=+IF(D43=0,0,IF(C43=0,0,(D43/C43)))");
$objPHPExcel->getActiveSheet()->setCellValue('F43', "=(M43-T43)");
$objPHPExcel->getActiveSheet()->setCellValue('G43', $general_ae3-$general_ae4);
$objPHPExcel->getActiveSheet()->setCellValue('H43', "=+F43-G43");
$objPHPExcel->getActiveSheet()->setCellValue('M43', $general_ae);
$objPHPExcel->getActiveSheet()->setCellValue('L43', $budget_general_ae);
$objPHPExcel->getActiveSheet()->setCellValue('Q43', $budget_general_ae2);
$objPHPExcel->getActiveSheet()->setCellValue('J43', "=(F43-L43)");
$objPHPExcel->getActiveSheet()->setCellValue('N43', "=(M43-Q43)");
$objPHPExcel->getActiveSheet()->setCellValue('K43', "=+IF(J43=0,0,(J43/L43))");
$objPHPExcel->getActiveSheet()->setCellValue('O43', "=+IF(N43=0,0,(N43/Q43))");
$objPHPExcel->getActiveSheet()->setCellValue('S43', "=+IF(M43=0,0,IF(R43=0,0,(M43/R43)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T43', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T43', $general_ae4);

$objPHPExcel->getActiveSheet()->setCellValue('B44', 'Depreciation');
$objPHPExcel->getActiveSheet()->setCellValue('C44', $depreciation4-$depreciation_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D44', "=(F44-C44)");
$objPHPExcel->getActiveSheet()->setCellValue('E44', "=+IF(D44=0,0,IF(C44=0,0,(D44/C44)))");
$objPHPExcel->getActiveSheet()->setCellValue('F44', "=(M44-T44)");
$objPHPExcel->getActiveSheet()->setCellValue('G44', $depreciation3-$depreciation4);
$objPHPExcel->getActiveSheet()->setCellValue('H44', "=+F44-G44");
$objPHPExcel->getActiveSheet()->setCellValue('M44', $depreciation);
$objPHPExcel->getActiveSheet()->setCellValue('L44', $budget_depreciation);
$objPHPExcel->getActiveSheet()->setCellValue('Q44', $budget_depreciation2);
$objPHPExcel->getActiveSheet()->setCellValue('J44', "=(F44-L44)");
$objPHPExcel->getActiveSheet()->setCellValue('N44', "=(M44-Q44)");
$objPHPExcel->getActiveSheet()->setCellValue('K44', "=+IF(J44=0,0,(J44/L44))");
$objPHPExcel->getActiveSheet()->setCellValue('O44', "=+IF(N44=0,0,(N44/Q44))");
$objPHPExcel->getActiveSheet()->setCellValue('S44', "=+IF(M44=0,0,IF(R44=0,0,(M44/R44)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T44', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T44', $depreciation4);

$objPHPExcel->getActiveSheet()->setCellValue('B45', 'Total Operating Expenses');
$objPHPExcel->getActiveSheet()->setCellValue('C45', '=SUM(C42:C44)');
$objPHPExcel->getActiveSheet()->setCellValue('D45', '=SUM(D42:D44)');
$objPHPExcel->getActiveSheet()->setCellValue('E45', "=+IF(D45=0,0,IF(C45=0,0,(D45/C45)))");
$objPHPExcel->getActiveSheet()->setCellValue('F45', '=SUM(F42:F44)');
$objPHPExcel->getActiveSheet()->setCellValue('G45', '=SUM(G42:G44)');
$objPHPExcel->getActiveSheet()->setCellValue('H45', '=SUM(H42:H44)');
$objPHPExcel->getActiveSheet()->setCellValue('M45', '=SUM(M42:M44)');
$objPHPExcel->getActiveSheet()->setCellValue('L45', '=SUM(L42:L44)');
$objPHPExcel->getActiveSheet()->setCellValue('Q45', '=SUM(Q42:Q44)');
$objPHPExcel->getActiveSheet()->setCellValue('J45', "=(F45-L45)");
$objPHPExcel->getActiveSheet()->setCellValue('N45', "=(M45-Q45)");
$objPHPExcel->getActiveSheet()->setCellValue('K45', "=+IF(J45=0,0,(J45/L45))");
$objPHPExcel->getActiveSheet()->setCellValue('O45', "=+IF(N45=0,0,(N45/Q45))");
$objPHPExcel->getActiveSheet()->setCellValue('S45', "=+IF(M45=0,0,IF(R45=0,0,(M45/R45)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T45', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T45', "=C45");

$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Other Operating Expense/Income');
$objPHPExcel->getActiveSheet()->setCellValue('C46', $other_oei4-$other_oei_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D46', "=(F46-C46)");
$objPHPExcel->getActiveSheet()->setCellValue('E46', "=+IF(D46=0,0,IF(C46=0,0,(D46/C46)))");
$objPHPExcel->getActiveSheet()->setCellValue('F46', "=(M46-T46)");
$objPHPExcel->getActiveSheet()->setCellValue('G46', $other_oei3-$other_oei4);
$objPHPExcel->getActiveSheet()->setCellValue('H46', "=+F46-G46");
$objPHPExcel->getActiveSheet()->setCellValue('M46', $other_oei);
$objPHPExcel->getActiveSheet()->setCellValue('L46', $budget_other_oei);
$objPHPExcel->getActiveSheet()->setCellValue('Q46', $budget_other_oei2);
$objPHPExcel->getActiveSheet()->setCellValue('J46', "=(F46-L46)");
$objPHPExcel->getActiveSheet()->setCellValue('N46', "=(M46-Q46)");
$objPHPExcel->getActiveSheet()->setCellValue('K46', "=+IF(J46=0,0,(J46/L46))");
$objPHPExcel->getActiveSheet()->setCellValue('O46', "=+IF(N46=0,0,(N46/Q46))");
$objPHPExcel->getActiveSheet()->setCellValue('S46', "=+IF(M46=0,0,IF(R46=0,0,(M46/R46)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T44', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T46', $other_oei4);





$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Operating Profit');
$objPHPExcel->getActiveSheet()->setCellValue('C47', '=C39+C45+C46');
$objPHPExcel->getActiveSheet()->setCellValue('D47', '=D39+D45+D46');
$objPHPExcel->getActiveSheet()->setCellValue('E47', '=E39+E45+E46');
$objPHPExcel->getActiveSheet()->setCellValue('F47', '=F39+F45+F46');
$objPHPExcel->getActiveSheet()->setCellValue('G47', '=G39+G45+G46');
$objPHPExcel->getActiveSheet()->setCellValue('H47', '=H39+H45+H46');
$objPHPExcel->getActiveSheet()->setCellValue('M47', '=M39+M45+M46');
$objPHPExcel->getActiveSheet()->setCellValue('L47', '=L39+L45+L46');
$objPHPExcel->getActiveSheet()->setCellValue('Q47', '=Q39+Q45+Q46');
$objPHPExcel->getActiveSheet()->setCellValue('J47', "=J39+J45+J46");
$objPHPExcel->getActiveSheet()->setCellValue('N47', "=N39+N45+N46");
$objPHPExcel->getActiveSheet()->setCellValue('K47', "=+IF(J47=0,0,(J47/L47))");
$objPHPExcel->getActiveSheet()->setCellValue('O47', "=+IF(N47=0,0,(N47/Q47))");
$objPHPExcel->getActiveSheet()->setCellValue('S47', "=+IF(M47=0,0,IF(R47=0,0,(M47/R47)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T47', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T47', "=C47");

$objPHPExcel->getActiveSheet()->setCellValue('B48', '');
$objPHPExcel->getActiveSheet()->setCellValue('B49', 'Provision');





$objPHPExcel->getActiveSheet()->setCellValue('B50', 'General Provision');
$objPHPExcel->getActiveSheet()->setCellValue('C50', $general_provision4-$general_provision_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D50', "=(F50-C50)");
$objPHPExcel->getActiveSheet()->setCellValue('E50', "=+IF(D50=0,0,IF(C50=0,0,(D50/C50)))");
$objPHPExcel->getActiveSheet()->setCellValue('F50', "=(M50-T50)");
$objPHPExcel->getActiveSheet()->setCellValue('G50', $general_provision3-$general_provision4);
$objPHPExcel->getActiveSheet()->setCellValue('H50', "=+F50-G50");
$objPHPExcel->getActiveSheet()->setCellValue('M50', $general_provision);
$objPHPExcel->getActiveSheet()->setCellValue('L50', $budget_general_provision);
$objPHPExcel->getActiveSheet()->setCellValue('Q50', $budget_general_provision2);
$objPHPExcel->getActiveSheet()->setCellValue('J50', "=(F50-L50)");
$objPHPExcel->getActiveSheet()->setCellValue('N50', "=(M50-Q50)");
$objPHPExcel->getActiveSheet()->setCellValue('K50', "=+IF(J50=0,0,(J50/L50))");
$objPHPExcel->getActiveSheet()->setCellValue('O50', "=+IF(N50=0,0,(N50/Q50))");
$objPHPExcel->getActiveSheet()->setCellValue('S50', "=+IF(M50=0,0,IF(R50=0,0,(M50/R50)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T50', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T50', $general_provision4);

$objPHPExcel->getActiveSheet()->setCellValue('B51', 'Specific Provision Charged');
$objPHPExcel->getActiveSheet()->setCellValue('C51', $specific_pc4-$specific_pc_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D51', "=(F51-C51)");
$objPHPExcel->getActiveSheet()->setCellValue('E51', "=+IF(D51=0,0,IF(C51=0,0,(D51/C51)))");
$objPHPExcel->getActiveSheet()->setCellValue('F51', "=(M51-T51)");
$objPHPExcel->getActiveSheet()->setCellValue('G51', $specific_pc3-$specific_pc4);
$objPHPExcel->getActiveSheet()->setCellValue('H51', "=+F51-G51");
$objPHPExcel->getActiveSheet()->setCellValue('M51', $specific_pc);
$objPHPExcel->getActiveSheet()->setCellValue('L51', $budget_specific_pc);
$objPHPExcel->getActiveSheet()->setCellValue('Q51', $budget_specific_pc2);
$objPHPExcel->getActiveSheet()->setCellValue('J51', "=(F51-L51)");
$objPHPExcel->getActiveSheet()->setCellValue('N51', "=(M15-Q51)");
$objPHPExcel->getActiveSheet()->setCellValue('K51', "=+IF(J51=0,0,(J51/L51))");
$objPHPExcel->getActiveSheet()->setCellValue('O51', "=+IF(N51=0,0,(N51/Q51))");
$objPHPExcel->getActiveSheet()->setCellValue('S51', "=+IF(M51=0,0,IF(R51=0,0,(M51/R51)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T51', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T51', $specific_pc4);

$objPHPExcel->getActiveSheet()->setCellValue('B52', 'Specific Provision Recovery  ');
$objPHPExcel->getActiveSheet()->setCellValue('C52', $specific_pr4-$specific_pr_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D52', "=(F52-C52)");
$objPHPExcel->getActiveSheet()->setCellValue('E52', "=+IF(D52=0,0,IF(C52=0,0,(D52/C52)))");
$objPHPExcel->getActiveSheet()->setCellValue('F52', "=(M52-T52)");
$objPHPExcel->getActiveSheet()->setCellValue('G52', $specific_pr3-$specific_pr4);
$objPHPExcel->getActiveSheet()->setCellValue('H52', "=+F52-G52");
$objPHPExcel->getActiveSheet()->setCellValue('M52', $specific_pr);
$objPHPExcel->getActiveSheet()->setCellValue('L52', $budget_specific_pr);
$objPHPExcel->getActiveSheet()->setCellValue('Q52', $budget_specific_pr2);
$objPHPExcel->getActiveSheet()->setCellValue('J52', "=(F52-L52)");
$objPHPExcel->getActiveSheet()->setCellValue('N52', "=(M52-Q52)");
$objPHPExcel->getActiveSheet()->setCellValue('K52', "=+IF(J52=0,0,(J52/L52))");
$objPHPExcel->getActiveSheet()->setCellValue('O52', "=+IF(N52=0,0,(N52/Q52))");
$objPHPExcel->getActiveSheet()->setCellValue('S52', "=+IF(M52=0,0,IF(R52=0,0,(M52/R52)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T52', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T52', $specific_pr4);

$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Write Offs Charged');
$objPHPExcel->getActiveSheet()->setCellValue('C53', $write_off_ch4-$write_off_ch_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D53', "=(F53-C53)");
$objPHPExcel->getActiveSheet()->setCellValue('E53', "=+IF(D53=0,0,IF(C53=0,0,(D53/C53)))");
$objPHPExcel->getActiveSheet()->setCellValue('F53', "=(M53-T53)");
$objPHPExcel->getActiveSheet()->setCellValue('G53', $write_off_ch3-$write_off_ch4);
$objPHPExcel->getActiveSheet()->setCellValue('H53', "=+F53-G53");
$objPHPExcel->getActiveSheet()->setCellValue('M53', $write_off_ch);
$objPHPExcel->getActiveSheet()->setCellValue('L53', $budget_write_off_ch);
$objPHPExcel->getActiveSheet()->setCellValue('Q53', $budget_write_off_ch2);
$objPHPExcel->getActiveSheet()->setCellValue('J53', "=(F53-L53)");
$objPHPExcel->getActiveSheet()->setCellValue('N53', "=(M53-Q53)");
$objPHPExcel->getActiveSheet()->setCellValue('K53', "=+IF(J53=0,0,(J53/L53))");
$objPHPExcel->getActiveSheet()->setCellValue('O53', "=+IF(N53=0,0,(N53/Q53))");
$objPHPExcel->getActiveSheet()->setCellValue('S53', "=+IF(M53=0,0,IF(R53=0,0,(M53/R53)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T53', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T53', $write_off_ch4);

$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Write Offs Recovered');
$objPHPExcel->getActiveSheet()->setCellValue('C54', $write_off_rec4-$write_off_rec_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D54', "=(F54-C54)");
$objPHPExcel->getActiveSheet()->setCellValue('E54', "=+IF(D54=0,0,IF(C54=0,0,(D54/C54)))");
$objPHPExcel->getActiveSheet()->setCellValue('F54', "=(M54-T54)");
$objPHPExcel->getActiveSheet()->setCellValue('G54', $write_off_rec3-$write_off_rec4);
$objPHPExcel->getActiveSheet()->setCellValue('H54', "=+F54-G54");
$objPHPExcel->getActiveSheet()->setCellValue('M54', $write_off_rec);
$objPHPExcel->getActiveSheet()->setCellValue('L54', $budget_write_off_rec);
$objPHPExcel->getActiveSheet()->setCellValue('Q54', $budget_write_off_rec2);
$objPHPExcel->getActiveSheet()->setCellValue('J54', "=(F54-L54)");
$objPHPExcel->getActiveSheet()->setCellValue('N54', "=(M54-Q54)");
$objPHPExcel->getActiveSheet()->setCellValue('K54', "=+IF(J54=0,0,(J54/L54))");
$objPHPExcel->getActiveSheet()->setCellValue('O54', "=+IF(N54=0,0,(N54/Q54))");
$objPHPExcel->getActiveSheet()->setCellValue('S54', "=+IF(M54=0,0,IF(R54=0,0,(M54/R54)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T54', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T54', $write_off_rec4);


$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Foreclose Properties Provision');
$objPHPExcel->getActiveSheet()->setCellValue('C55', $foreclose_pp4-$foreclose_pp_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D55', "=(F55-C55)");
$objPHPExcel->getActiveSheet()->setCellValue('E55', "=+IF(D55=0,0,IF(C55=0,0,(D55/C55)))");
$objPHPExcel->getActiveSheet()->setCellValue('F55', "=(M55-T55)");
$objPHPExcel->getActiveSheet()->setCellValue('G55', $foreclose_pp3-$foreclose_pp4);
$objPHPExcel->getActiveSheet()->setCellValue('H55', "=+F55-G55");
$objPHPExcel->getActiveSheet()->setCellValue('M55', $foreclose_pp);
$objPHPExcel->getActiveSheet()->setCellValue('L55', $budget_foreclose_pp);
$objPHPExcel->getActiveSheet()->setCellValue('Q55', $budget_foreclose_pp2);
$objPHPExcel->getActiveSheet()->setCellValue('J55', "=(F55-L55)");
$objPHPExcel->getActiveSheet()->setCellValue('N55', "=(M55-Q55)");
$objPHPExcel->getActiveSheet()->setCellValue('K55', "=+IF(J55=0,0,(J55/L55))");
$objPHPExcel->getActiveSheet()->setCellValue('O55', "=+IF(N55=0,0,(N55/Q55))");
$objPHPExcel->getActiveSheet()->setCellValue('S55', "=+IF(M55=0,0,IF(R55=0,0,(M55/R55)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T55', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T55', $foreclose_pp4);

$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('C56', $other_provision4-$other_provision_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D56', "=(F56-C56)");
$objPHPExcel->getActiveSheet()->setCellValue('E56', "=+IF(D56=0,0,IF(C56=0,0,(D56/C56)))");
$objPHPExcel->getActiveSheet()->setCellValue('F56', "=(M56-T56)");
$objPHPExcel->getActiveSheet()->setCellValue('G56', $other_provision3-$other_provision4);
$objPHPExcel->getActiveSheet()->setCellValue('H56', "=+F56-G56");
$objPHPExcel->getActiveSheet()->setCellValue('M56', $other_provision);
$objPHPExcel->getActiveSheet()->setCellValue('L56', $budget_other_provision);
$objPHPExcel->getActiveSheet()->setCellValue('Q56', $budget_other_provision2);
$objPHPExcel->getActiveSheet()->setCellValue('J56', "=(F56-L56)");
$objPHPExcel->getActiveSheet()->setCellValue('N56', "=(M56-Q56)");
$objPHPExcel->getActiveSheet()->setCellValue('K56', "=+IF(J56=0,0,(J56/L56))");
$objPHPExcel->getActiveSheet()->setCellValue('O56', "=+IF(N56=0,0,(N56/Q56))");
$objPHPExcel->getActiveSheet()->setCellValue('S56', "=+IF(M56=0,0,IF(R56=0,0,(M56/R56)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T56', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T56', $other_provision4);

$objPHPExcel->getActiveSheet()->setCellValue('B57', 'Total Provision');
$objPHPExcel->getActiveSheet()->setCellValue('C57', $tot_provision4-$tot_provision_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D57', "=SUM(D50:D56)");
$objPHPExcel->getActiveSheet()->setCellValue('E57', "=+IF(D57=0,0,IF(C57=0,0,(D57/C57)))");
$objPHPExcel->getActiveSheet()->setCellValue('F57', "=SUM(F50:F56)");
$objPHPExcel->getActiveSheet()->setCellValue('G57', $tot_provision3-$tot_provision4);
$objPHPExcel->getActiveSheet()->setCellValue('H57', "=SUM(H50:H56)");
$objPHPExcel->getActiveSheet()->setCellValue('M57', "=SUM(M50:M56)");
$objPHPExcel->getActiveSheet()->setCellValue('L57', "=SUM(L50:L56)");
$objPHPExcel->getActiveSheet()->setCellValue('Q57', "=SUM(Q50:Q56)");
$objPHPExcel->getActiveSheet()->setCellValue('J57', "=SUM(J50:J56)");
$objPHPExcel->getActiveSheet()->setCellValue('N57', "=SUM(N50:N56)");
$objPHPExcel->getActiveSheet()->setCellValue('K57', "=+IF(J57=0,0,(J57/L57))");
$objPHPExcel->getActiveSheet()->setCellValue('O57', "=+IF(N57=0,0,(N57/Q57))");
$objPHPExcel->getActiveSheet()->setCellValue('S57', "=+IF(M57=0,0,IF(R57=0,0,(M57/R57)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T57', $acc_is_treasury);
$objPHPExcel->getActiveSheet()->setCellValue('T57', $tot_provision4);

$objPHPExcel->getActiveSheet()->setCellValue('B58', 'Extraordinary item');
$objPHPExcel->getActiveSheet()->setCellValue('B59', 'Profit (Loss) Before Tax');
$objPHPExcel->getActiveSheet()->setCellValue('C59', '=C47+C57+C58');
$objPHPExcel->getActiveSheet()->setCellValue('D59', '=D47+D57+D58');
$objPHPExcel->getActiveSheet()->setCellValue('F59', '=E47+E57+E58');
$objPHPExcel->getActiveSheet()->setCellValue('G59', '=F47+F57+F58');
$objPHPExcel->getActiveSheet()->setCellValue('H59', '=H47+H57+H58');
$objPHPExcel->getActiveSheet()->setCellValue('M59', '=M47+M57+M58');
$objPHPExcel->getActiveSheet()->setCellValue('L59', '=L47+L57+L58');
$objPHPExcel->getActiveSheet()->setCellValue('Q59', '=Q47+Q57+Q58');
$objPHPExcel->getActiveSheet()->setCellValue('J59', "=J47+J57+J58");
$objPHPExcel->getActiveSheet()->setCellValue('N59', "=N47+N57+N58");
$objPHPExcel->getActiveSheet()->setCellValue('K59', "=+IF(J59=0,0,(J59/L59))");
$objPHPExcel->getActiveSheet()->setCellValue('O59', "=+IF(N59=0,0,(N59/Q59))");
$objPHPExcel->getActiveSheet()->setCellValue('S59', "=+IF(M59=0,0,IF(R59=0,0,(M59/R59)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T59', "=T47+T57+T58");
$objPHPExcel->getActiveSheet()->setCellValue('T59', "=C59");

$objPHPExcel->getActiveSheet()->setCellValue('B60', 'Corporate Tax');
$objPHPExcel->getActiveSheet()->setCellValue('C60', $corporate_tax4-$corporate_tax_m2);
$objPHPExcel->getActiveSheet()->setCellValue('D60', "=(F60-C60)");
$objPHPExcel->getActiveSheet()->setCellValue('E60', "=+IF(D60=0,0,IF(C60=0,0,(D60/C60)))");
$objPHPExcel->getActiveSheet()->setCellValue('F60', "=(M60-T60)");
$objPHPExcel->getActiveSheet()->setCellValue('G60', $corporate_tax3-$corporate_tax4);
$objPHPExcel->getActiveSheet()->setCellValue('H60', "=+F60-G60");
//$objPHPExcel->getActiveSheet()->setCellValue('M60', $corporate_tax*$curr_pajak);
$objPHPExcel->getActiveSheet()->setCellValue('M60', "=M59*($curr_pajak)");
$objPHPExcel->getActiveSheet()->setCellValue('L60', $budget_corporate_tax);
$objPHPExcel->getActiveSheet()->setCellValue('Q60', $budget_corporate_tax2);
$objPHPExcel->getActiveSheet()->setCellValue('J60', "=(F60-L60)");
$objPHPExcel->getActiveSheet()->setCellValue('N60', "=(M60-Q60)");
$objPHPExcel->getActiveSheet()->setCellValue('K60', "=+IF(J60=0,0,(J60/L60))");
$objPHPExcel->getActiveSheet()->setCellValue('O60', "=+IF(N60=0,0,(N60/Q60))");
$objPHPExcel->getActiveSheet()->setCellValue('S60', "=+IF(M60=0,0,IF(R60=0,0,(M60/R60)))");
//$objPHPExcel->getActiveSheet()->setCellValue('T60', 0);
$objPHPExcel->getActiveSheet()->setCellValue('T60', $corporate_tax4);


##### ONLY PRINT  ###########################

//$objPHPExcel->getActiveSheet()->setCellValue('M600', $corporate_tax);
//$objPHPExcel->getActiveSheet()->setCellValue('M601', $curr_pajak);

#############################################



$objPHPExcel->getActiveSheet()->setCellValue('B61', 'Profit (Loss) After Tax');
$objPHPExcel->getActiveSheet()->setCellValue('C61', '=SUM(C59:C60)');
$objPHPExcel->getActiveSheet()->setCellValue('D61', '=SUM(D59:D60)');
$objPHPExcel->getActiveSheet()->setCellValue('E61', '=SUM(E59:E60)');
$objPHPExcel->getActiveSheet()->setCellValue('F61', '=SUM(F59:F60)');
$objPHPExcel->getActiveSheet()->setCellValue('G61', '=SUM(G59:G60)');
$objPHPExcel->getActiveSheet()->setCellValue('H61', '=SUM(H59:H60)');
$objPHPExcel->getActiveSheet()->setCellValue('M61', '=SUM(M59:M60)');
$objPHPExcel->getActiveSheet()->setCellValue('L61', '=SUM(L59:L60)');
$objPHPExcel->getActiveSheet()->setCellValue('Q61', '=SUM(Q59:Q60)');
$objPHPExcel->getActiveSheet()->setCellValue('J61', "=SUM(J59:J60)");
$objPHPExcel->getActiveSheet()->setCellValue('N61', "=SUM(N59:N60)");
$objPHPExcel->getActiveSheet()->setCellValue('K61', "=SUM(K59:K60)");
$objPHPExcel->getActiveSheet()->setCellValue('O61', "=SUM(O59:O60)");
$objPHPExcel->getActiveSheet()->setCellValue('S61', "=SUM(S59:S60)");
//$objPHPExcel->getActiveSheet()->setCellValue('T61', "=SUM(T59:T60)");
$objPHPExcel->getActiveSheet()->setCellValue('T61', "=C61");


$objPHPExcel->getActiveSheet()->setCellValue('B62', '');
$objPHPExcel->getActiveSheet()->setCellValue('B63', '');
$objPHPExcel->getActiveSheet()->setCellValue('B64', '');

$objPHPExcel->getActiveSheet()->setCellValue('T5', 'YTD');

$objPHPExcel->getActiveSheet()->setCellValue('C5', 'MTD');
$objPHPExcel->getActiveSheet()->setCellValue('M5', 'YTD');
$objPHPExcel->getActiveSheet()->setCellValue('R5', 'Whole Year  Budget');



$objPHPExcel->getActiveSheet()->setCellValue('C6', $prev_date);
$objPHPExcel->getActiveSheet()->setCellValue('D6', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('E6', '%');
$objPHPExcel->getActiveSheet()->setCellValue('F6', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('G6', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('H6', 'Var');
$objPHPExcel->getActiveSheet()->setCellValue('I6', 'MTD Forecast '.$label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('J6', 'Var MTD v.s Budget');
$objPHPExcel->getActiveSheet()->setCellValue('K6', '%');
$objPHPExcel->getActiveSheet()->setCellValue('L6', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('M6', $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('N6', 'Var YTD v.s Budget');
$objPHPExcel->getActiveSheet()->setCellValue('O6', '%');
$objPHPExcel->getActiveSheet()->setCellValue('P6', 'YTD Forecast '. $label_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('Q6', 'Budget');
$objPHPExcel->getActiveSheet()->setCellValue('R6', 'Rp');
$objPHPExcel->getActiveSheet()->setCellValue('S6', '%');
$objPHPExcel->getActiveSheet()->setCellValue('T6', $prev_date);
//$objPHPExcel->getActiveSheet()->setCellValue('U6', $label_tgl_year_min1);

#------------------------ QUERY FORECAST---------------------------------------------------------
$query_forecast=" select Proyeksi_MTD, Proyeksi_YTD from Master_Forecast2 WHERE ";
$tgl_forecast=" Data_Date ='$curr_tgl' ";

#-----------  FLASH101000007  Loan
        $level_forecast=" and FLASH_Level_3='FLASH201000002' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $loan_r_mtd=$row_forecast['Proyeksi_MTD'];
        $loan_r_ytd=$row_forecast['Proyeksi_YTD'];
#-----------  FLASH201000003  Treasury bills
        $level_forecast=" and FLASH_Level_3='FLASH201000003' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $treasury_r_mtd=$row_forecast['Proyeksi_MTD'];
        $treasury_r_ytd=$row_forecast['Proyeksi_YTD'];
//echo $query_forecast.$tgl_forecast.$level_forecast;
//die();
#-----------  FLASH101000004  Interbank placements
        $level_forecast=" and FLASH_Level_3='FLASH201000004' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $interbank_r_placement_mtd=$row_forecast['Proyeksi_MTD'];
        $interbank_r_placement_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000005  Placement with BI
        $level_forecast=" and FLASH_Level_3='FLASH201000005' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $placement_wbi_r_mtd=$row_forecast['Proyeksi_MTD'];
        $placement_wbi_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH101000019  Others
        $level_forecast=" and FLASH_Level_3='FLASH201000006' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others1_r_mtd=$row_forecast['Proyeksi_MTD'];
        $others1_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH102000001  Current accounts
        $level_forecast=" and FLASH_Level_3='FLASH202000002' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $current_account_r_mtd=$row_forecast['Proyeksi_MTD'];
        $current_account_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH102000002  Saving accounts
        $level_forecast=" and FLASH_Level_3='FLASH202000003' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $saving_account_r_mtd=$row_forecast['Proyeksi_MTD'];
        $saving_account_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH102000003  Time deposits
        $level_forecast=" and FLASH_Level_3='FLASH202000004' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $time_deposits_r_mtd=$row_forecast['Proyeksi_MTD'];
        $time_deposits_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH102000006  Bank deposits
        $level_forecast=" and FLASH_Level_3='FLASH202000005' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $bank_deposits_r_mtd=$row_forecast['Proyeksi_MTD'];
        $bank_deposits_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000007  Borrowings (MCB)
        $level_forecast=" and FLASH_Level_3='FLASH202000007' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $borrowings_mcb_r_mtd=$row_forecast['Proyeksi_MTD'];
        $borrowings_mcb_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000008  Guaranteed premium
        $level_forecast=" and FLASH_Level_3='FLASH202000008' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $guaranted_premium_r_mtd=$row_forecast['Proyeksi_MTD'];
        $guaranted_premium_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000006  Others
        $level_forecast=" and FLASH_Level_3='FLASH202000009' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others2_r_mtd=$row_forecast['Proyeksi_MTD'];
        $others2_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000008  Forex gain/(loss) on transactions
        $level_forecast=" and FLASH_Level_3='FLASH201000008' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $forex_gain_r_mtd=$row_forecast['Proyeksi_MTD'];
        $forex_gain_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000009  Gain/(Loss) on sale of securities/bonds
        $level_forecast=" and FLASH_Level_3='FLASH201000009' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $gain_loss_r_mtd=$row_forecast['Proyeksi_MTD'];
        $gain_loss_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000010  Remittance fee
        $level_forecast=" and FLASH_Level_3='FLASH201000010' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $remittance_fee_r_mtd=$row_forecast['Proyeksi_MTD'];
        $remittance_fee_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000011  Trade Finance fee
        $level_forecast=" and FLASH_Level_3='FLASH201000011' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $trade_finance_r_mtd=$row_forecast['Proyeksi_MTD'];
        $trade_finance_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000012  Processing fee
        $level_forecast=" and FLASH_Level_3='FLASH201000012' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $processing_fee_r_mtd=$row_forecast['Proyeksi_MTD'];
        $processing_fee_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000013  Credit Card fee
        $level_forecast=" and FLASH_Level_3='FLASH201000013' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $credit_card_fee_r_mtd=$row_forecast['Proyeksi_MTD'];
        $credit_card_fee_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000014  Insurance Fee
        $level_forecast=" and FLASH_Level_3='FLASH201000014' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $insurance_fee_r_mtd=$row_forecast['Proyeksi_MTD'];
        $insurance_fee_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000015  Service Charges
        $level_forecast=" and FLASH_Level_3='FLASH201000015' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $service_charge_r_mtd=$row_forecast['Proyeksi_MTD'];
        $service_charge_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000016  Other Commission & Fee
        $level_forecast=" and FLASH_Level_3='FLASH201000016' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others_commision_r_mtd=$row_forecast['Proyeksi_MTD'];
        $others_commision_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000017  Penalty
        $level_forecast=" and FLASH_Level_3='FLASH201000017' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $penalty_r_mtd=$row_forecast['Proyeksi_MTD'];
        $penalty_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000018  Write Offs Recovered
        $level_forecast=" and FLASH_Level_3='FLASH201000018' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_r_mtd=$row_forecast['Proyeksi_MTD'];
        $write_off_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH201000019  Other Income
        $level_forecast=" and FLASH_Level_3='FLASH201000019' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $other_income_r_mtd=$row_forecast['Proyeksi_MTD'];
        $other_income_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000010  Staff Cost
        $level_forecast=" and FLASH_Level_3='FLASH202000010' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $staff_cost_r_mtd=$row_forecast['Proyeksi_MTD'];
        $staff_cost_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000011  General & Administrative Expenses
        $level_forecast=" and FLASH_Level_3='FLASH202000011' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $general_administrative_r_mtd=$row_forecast['Proyeksi_MTD'];
        $general_administrative_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000012  Depreciation
        $level_forecast=" and FLASH_Level_3='FLASH202000012' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $depreciation_r_mtd=$row_forecast['Proyeksi_MTD'];
        $depreciation_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000014  Other Operating Expense/income
        $level_forecast=" and FLASH_Level_3='FLASH202000014' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $other_operating_r_mtd=$row_forecast['Proyeksi_MTD'];
        $other_operating_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000015  General Provision
        $level_forecast=" and FLASH_Level_3='FLASH202000015' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $general_provision_r_mtd=$row_forecast['Proyeksi_MTD'];
        $general_provision_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000016  Specific Provision Charged
        $level_forecast=" and FLASH_Level_3='FLASH202000016' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $specific_provision_c_r_mtd=$row_forecast['Proyeksi_MTD'];
        $specific_provision_c_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000017  Specific Provision Recovery  
        $level_forecast=" and FLASH_Level_3='FLASH202000017' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $specific_provision_r_r_mtd=$row_forecast['Proyeksi_MTD'];
        $specific_provision_r_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000018  Write Offs Charged
        $level_forecast=" and FLASH_Level_3='FLASH202000018' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_charge_r_mtd=$row_forecast['Proyeksi_MTD'];
        $write_off_charge_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000019  Write Offs Recovered
        $level_forecast=" and FLASH_Level_3='FLASH202000019' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_recover_r_mtd=$row_forecast['Proyeksi_MTD'];
        $write_off_recover_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000020  Foreclose Properties Provision
        $level_forecast=" and FLASH_Level_3='FLASH202000020' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $foreclose_propertis_r_mtd=$row_forecast['Proyeksi_MTD'];
        $foreclose_propertis_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000021  Others
        $level_forecast=" and FLASH_Level_3='FLASH202000021' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others3_r_mtd=$row_forecast['Proyeksi_MTD'];
        $others3_r_ytd=$row_forecast['Proyeksi_YTD'];

#-----------  FLASH202000023  Corporate Tax
        $level_forecast=" and FLASH_Level_3='FLASH202000023' ";
        $result_forecast=odbc_exec($connection2, $query_forecast.$tgl_forecast.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $corporate_tax_r_mtd=$row_forecast['Proyeksi_MTD'];
        $corporate_tax_r_ytd=$row_forecast['Proyeksi_YTD'];

#-------------  Output FORECAST MTD -------------------------
$objPHPExcel->getActiveSheet()->setCellValue('I9', $loan_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I10', $treasury_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I11', $interbank_r_placement_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('I12', $placement_wbi_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('I13', $others1_r_mtd);

$objPHPExcel->getActiveSheet()->setCellValue('I15', $current_account_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I16', $saving_account_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I17', $time_deposits_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $bank_deposits_r_mtd);

$objPHPExcel->getActiveSheet()->setCellValue('I20', $borrowings_mcb_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I21', $guaranted_premium_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I22', $others2_r_mtd);

$objPHPExcel->getActiveSheet()->setCellValue('I26', $forex_gain_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I27', $gain_loss_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I28', $remittance_fee_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I29', $trade_finance_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I30', $processing_fee_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I31', $credit_card_fee_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I32', $insurance_fee_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I33', $service_charge_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I34', $others_commision_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I35', $penalty_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I36', $write_off_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I37', $other_income_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I38', "=SUM(C26:C37)");

$objPHPExcel->getActiveSheet()->setCellValue('I42', $staff_cost_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I43', $general_administrative_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I44', $depreciation_r_mtd);

$objPHPExcel->getActiveSheet()->setCellValue('I44', $other_operating_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I50', $general_provision_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I51', $specific_provision_c_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I52', $specific_provision_r_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I53', $write_off_charge_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I54', $write_off_recover_r_mtd);
$objPHPExcel->getActiveSheet()->setCellValue('I55', $foreclose_propertis_r_mtd);

$objPHPExcel->getActiveSheet()->setCellValue('I57', "=SUM(I50:I56)");

$objPHPExcel->getActiveSheet()->setCellValue('I60', $corporate_tax_r_mtd);

#-------------  Output FORECAST YTD -------------------------
$objPHPExcel->getActiveSheet()->setCellValue('P9', $loan_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P10', $treasury_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P11', $interbank_r_placement_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P12', $placement_wbi_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P13', $others1_r_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('P15', $current_account_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P16', $saving_account_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P17', $time_deposits_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P18', $bank_deposits_r_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('P20', $borrowings_mcb_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P21', $guaranted_premium_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P22', $others2_r_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('P26', $forex_gain_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P27', $gain_loss_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P28', $remittance_fee_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P29', $trade_finance_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P30', $processing_fee_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P31', $credit_card_fee_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P32', $insurance_fee_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P33', $service_charge_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P34', $others_commision_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P35', $penalty_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P36', $write_off_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P37', $other_income_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P38', "=SUM(C26:C37)");

$objPHPExcel->getActiveSheet()->setCellValue('P42', $staff_cost_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P43', $general_administrative_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P44', $depreciation_r_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('P44', $other_operating_r_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('P50', $general_provision_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P51', $specific_provision_c_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P52', $specific_provision_r_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P53', $write_off_charge_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P54', $write_off_recover_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P55', $foreclose_propertis_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('P57', "=SUM(P50:P56)");

$objPHPExcel->getActiveSheet()->setCellValue('P60', $corporate_tax_r_ytd);





#------------------------ BUDGET PL (R) BULAN TERAKHIR ---------------------------------------------------------
$query_budget_r=" select Budget_MTD,Budget_YTD from Budget_PL where datepart(month,DataDate) ='12' and datepart(year,datadate) = '$year_budget' ";
//$tgl_forecast=" Data_Date ='$curr_tgl' ";

#-----------  FLASH101000007  Loan
        $level_forecast=" and FLASH_Level_3='FLASH101000007' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $loan_rp_mtd=$row_forecast['Budget_MTD'];
        $loan_rp_ytd=$row_forecast['Budget_YTD'];
#-----------  FLASH201000003  Treasury bills
        $level_forecast=" and FLASH_Level_3='FLASH201000003' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $treasury_rp_mtd=$row_forecast['Budget_MTD'];
        $treasury_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH101000004  Interbank placements
        $level_forecast=" and FLASH_Level_3='FLASH101000004' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $interbank_rp_placement_mtd=$row_forecast['Budget_MTD'];
        $interbank_rp_placement_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000005  Placement with BI
        $level_forecast=" and FLASH_Level_3='FLASH201000005' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $placement_wbi_rp_mtd=$row_forecast['Budget_MTD'];
        $placement_wbi_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH101000019  Others
        $level_forecast=" and FLASH_Level_3='FLASH101000019' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others1_rp_mtd=$row_forecast['Budget_MTD'];
        $others1_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH102000001  Current accounts
        $level_forecast=" and FLASH_Level_3='FLASH102000001' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $current_account_rp_mtd=$row_forecast['Budget_MTD'];
        $current_account_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH102000002  Saving accounts
        $level_forecast=" and FLASH_Level_3='FLASH102000002' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $saving_account_rp_mtd=$row_forecast['Budget_MTD'];
        $saving_account_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH102000003  Time deposits
        $level_forecast=" and FLASH_Level_3='FLASH102000003' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $time_deposits_rp_mtd=$row_forecast['Budget_MTD'];
        $time_deposits_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH102000006  Bank deposits
        $level_forecast=" and FLASH_Level_3='FLASH102000006' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $bank_deposits_rp_mtd=$row_forecast['Budget_MTD'];
        $bank_deposits_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000007  Borrowings (MCB)
        $level_forecast=" and FLASH_Level_3='FLASH202000007' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $borrowings_mcb_rp_mtd=$row_forecast['Budget_MTD'];
        $borrowings_mcb_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000008  Guaranteed premium
        $level_forecast=" and FLASH_Level_3='FLASH202000008' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $guaranted_premium_rp_mtd=$row_forecast['Budget_MTD'];
        $guaranted_premium_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000006  Others
        $level_forecast=" and FLASH_Level_3='FLASH202000006' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others2_rp_mtd=$row_forecast['Budget_MTD'];
        $others2_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000008  Forex gain/(loss) on transactions
        $level_forecast=" and FLASH_Level_3='FLASH201000008' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $forex_gain_rp_mtd=$row_forecast['Budget_MTD'];
        $forex_gain_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000009  Gain/(Loss) on sale of securities/bonds
        $level_forecast=" and FLASH_Level_3='FLASH201000009' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $gain_loss_rp_mtd=$row_forecast['Budget_MTD'];
        $gain_loss_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000010  Remittance fee
        $level_forecast=" and FLASH_Level_3='FLASH201000010' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $remittance_fee_rp_mtd=$row_forecast['Budget_MTD'];
        $remittance_fee_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000011  Trade Finance fee
        $level_forecast=" and FLASH_Level_3='FLASH201000011' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $trade_finance_rp_mtd=$row_forecast['Budget_MTD'];
        $trade_finance_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000012  Processing fee
        $level_forecast=" and FLASH_Level_3='FLASH201000012' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $processing_fee_rp_mtd=$row_forecast['Budget_MTD'];
        $processing_fee_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000013  Credit Card fee
        $level_forecast=" and FLASH_Level_3='FLASH201000013' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $credit_card_fee_rp_mtd=$row_forecast['Budget_MTD'];
        $credit_card_fee_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000014  Insurance Fee
        $level_forecast=" and FLASH_Level_3='FLASH201000014' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $insurance_fee_rp_mtd=$row_forecast['Budget_MTD'];
        $insurance_fee_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000015  Service Charges
        $level_forecast=" and FLASH_Level_3='FLASH201000015' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $service_charge_rp_mtd=$row_forecast['Budget_MTD'];
        $service_charge_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000016  Other Commission & Fee
        $level_forecast=" and FLASH_Level_3='FLASH201000016' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others_commision_rp_mtd=$row_forecast['Budget_MTD'];
        $others_commision_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000017  Penalty
        $level_forecast=" and FLASH_Level_3='FLASH201000017' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $penalty_rp_mtd=$row_forecast['Budget_MTD'];
        $penalty_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000018  Write Offs Recovered
        $level_forecast=" and FLASH_Level_3='FLASH201000018' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_rp_mtd=$row_forecast['Budget_MTD'];
        $write_off_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH201000019  Other Income
        $level_forecast=" and FLASH_Level_3='FLASH201000019' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $other_income_rp_mtd=$row_forecast['Budget_MTD'];
        $other_income_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000010  Staff Cost
        $level_forecast=" and FLASH_Level_3='FLASH202000010' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $staff_cost_rp_mtd=$row_forecast['Budget_MTD'];
        $staff_cost_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000011  General & Administrative Expenses
        $level_forecast=" and FLASH_Level_3='FLASH202000011' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $general_administrative_rp_mtd=$row_forecast['Budget_MTD'];
        $general_administrative_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000012  Depreciation
        $level_forecast=" and FLASH_Level_3='FLASH202000012' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $depreciation_rp_mtd=$row_forecast['Budget_MTD'];
        $depreciation_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000014  Other Operating Expense/income
        $level_forecast=" and FLASH_Level_3='FLASH202000014' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $other_operating_rp_mtd=$row_forecast['Budget_MTD'];
        $other_operating_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000015  General Provision
        $level_forecast=" and FLASH_Level_3='FLASH202000015' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $general_provision_rp_mtd=$row_forecast['Budget_MTD'];
        $general_provision_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000016  Specific Provision Charged
        $level_forecast=" and FLASH_Level_3='FLASH202000016' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $specific_provision_c_rp_mtd=$row_forecast['Budget_MTD'];
        $specific_provision_c_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000017  Specific Provision Recovery  
        $level_forecast=" and FLASH_Level_3='FLASH202000017' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $specific_provision_rp_r_mtd=$row_forecast['Budget_MTD'];
        $specific_provision_rp_r_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000018  Write Offs Charged
        $level_forecast=" and FLASH_Level_3='FLASH202000018' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_charge_rp_mtd=$row_forecast['Budget_MTD'];
        $write_off_charge_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000019  Write Offs Recovered
        $level_forecast=" and FLASH_Level_3='FLASH202000019' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $write_off_recover_rp_mtd=$row_forecast['Budget_MTD'];
        $write_off_recover_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000020  Foreclose Properties Provision
        $level_forecast=" and FLASH_Level_3='FLASH202000020' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $foreclose_propertis_rp_mtd=$row_forecast['Budget_MTD'];
        $foreclose_propertis_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000021  Others
        $level_forecast=" and FLASH_Level_3='FLASH202000021' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $others3_rp_mtd=$row_forecast['Budget_MTD'];
        $others3_rp_ytd=$row_forecast['Budget_YTD'];

#-----------  FLASH202000023  Corporate Tax
        $level_forecast=" and FLASH_Level_3='FLASH202000023' ";
        $result_forecast=odbc_exec($connection2, $query_budget_r.$level_forecast);
        $row_forecast=odbc_fetch_array($result_forecast);
        $corporate_tax_rp_mtd=$row_forecast['Budget_MTD'];
        $corporate_tax_rp_ytd=$row_forecast['Budget_YTD'];

#-------------  Output FORECAST YTD -------------------------
$objPHPExcel->getActiveSheet()->setCellValue('R9', $loan_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R10', $treasury_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R11', $interbank_rp_placement_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R12', $placement_wbi_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R13', $others1_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R14', "=SUM(R15:R18)");
$objPHPExcel->getActiveSheet()->setCellValue('R15', $current_account_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R16', $saving_account_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R17', $time_deposits_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R18', $bank_deposits_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R19', "=SUM(R20:R22)");
$objPHPExcel->getActiveSheet()->setCellValue('R20', $borrowings_mcb_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R21', $guaranted_premium_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R22', $others2_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R23', "=+R8+R14+R19");

$objPHPExcel->getActiveSheet()->setCellValue('R26', $forex_gain_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R27', $gain_loss_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R28', $remittance_fee_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R29', $trade_finance_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R30', $processing_fee_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R31', $credit_card_fee_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R32', $insurance_fee_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R33', $service_charge_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R34', $others_commision_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R35', $penalty_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R36', $write_off_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R37', $other_income_rp_ytd);
//$objPHPExcel->getActiveSheet()->setCellValue('R38', "");    
$objPHPExcel->getActiveSheet()->setCellValue('R38', "=SUM(R26:R37)");
$objPHPExcel->getActiveSheet()->setCellValue('R39', "=R23+R38");

$objPHPExcel->getActiveSheet()->setCellValue('R42', $staff_cost_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R43', $general_administrative_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R44', $depreciation_rp_ytd);    
$objPHPExcel->getActiveSheet()->setCellValue('R45', "=SUM(R42:R44)");
$objPHPExcel->getActiveSheet()->setCellValue('R47', "=R39+R45+R46");

$objPHPExcel->getActiveSheet()->setCellValue('R46', $other_operating_rp_ytd);

$objPHPExcel->getActiveSheet()->setCellValue('R50', $general_provision_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R51', $specific_provision_c_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R52', $specific_provision_rp_r_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R53', $write_off_charge_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R54', $write_off_recover_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R55', $foreclose_propertis_rp_ytd);
$objPHPExcel->getActiveSheet()->setCellValue('R57', "=SUM(R50:R56)");
$objPHPExcel->getActiveSheet()->setCellValue('R59', "=R47+R57+R58");
$objPHPExcel->getActiveSheet()->setCellValue('R61', "=R59+R60");



$objPHPExcel->getActiveSheet()->setCellValue('R60', $corporate_tax_rp_ytd);

for ($i=8;$i<24;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}


for ($i=8;$i<24;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('L'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, 0);
    }
}

for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('N'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('N'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('O'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('P'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('P'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('Q'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('R'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('R'.$i, 0);
    }
}

for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('S'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, 0);
    }
}
for ($i=8;$i<24;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('T'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('T'.$i, 0);
    }
}


$objPHPExcel->getActiveSheet()->getStyle('E8:E23')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('E26:E39')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('E42:E47')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('E50:E61')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));

$objPHPExcel->getActiveSheet()->getStyle('K8:K23')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('K26:K39')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('K42:K47')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('K50:K61')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));

$objPHPExcel->getActiveSheet()->getStyle('O8:O23')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('O26:O39')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('O42:O47')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('O50:O61')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));

$objPHPExcel->getActiveSheet()->getStyle('S8:S23')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('S26:S39')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('S42:S47')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));
$objPHPExcel->getActiveSheet()->getStyle('S50:S61')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));


// Rename 2nd sheet
$objPHPExcel->getActiveSheet()->setTitle('IS_Flash_Report');
//border
$objPHPExcel->getActiveSheet()->getStyle('B5:T61')->applyFromArray($styleArray);


// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.ms-excel');
//header('Content-Disposition: attachment;filename="Flash_Report_'.$label_tgl.'.xls"');
//header('Cache-Control: max-age=0');


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('php://output');
//$objWriter->save(str_replace(__FILE__,'/path/to/save/filename.extension',__FILE__));
//$objWriter->save($_SERVER['DOCUMENT_ROOT'] .'/filenamexxxxxxxx.xls');
$objWriter->save("download/Flash_Report_".$label_tgl."_".$file_eksport.".xls");
?>
<!--
<div align="center">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Flash_Report_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel  <i class="fa fa-arrow-circle-o-down"></i> </a> </div><br><br>
-->
                    <div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> List Flash Report
                            </div>
							<!--
                            <div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>
                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div> -->
                        </div>
                        <div class="portlet-body">
                            <h4><b>PT Bank MNC Internasional, Tbk</b></h4>
                            <br>
							<div class="tabbable-line">
														
							<?php
							//echo "<div class='alert alert-info alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Tax is on month $waktu_pajak </strong> </div>";
							echo "<div class='alert alert-danger'><strong> Data using the tax in $waktu_pajak </strong> </div>";
							//echo "<div class='alert alert-danger alert-dismissable' ><b> Data menggunakan Pajak Bulan $waktu_pajak </b> </div>";
							?>
							
							<div class="pull-right" style="font-size:12px">
								<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Flash_Report_".$label_tgl."_".$file_eksport.".xls";?>" 
									class="btn btn-sm green"> Download Excel  <i class="fa fa-arrow-circle-o-down"></i> </a> 
							</div> 
							<br>
							<br>
							    <ul class="nav nav-tabs ">
                                    <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        BS Flash Report </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        IS Flash Report </a>
                                    </li>
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                        <!--<p>
                                            <h5><b>BS Flash Report</b></h5>
                                        </p>
                                        -->
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-bordered table-hover" width="100%">
                                                <thead>
                                                
                                                </thead>
                                                <tbody>
                                                <tr class="active">
                                                <td width="20%" rowspan="2" align="center"> Account Of Assets </td>
                                                <td width="50%" colspan="5" align="center">For The Month</td>
                                                <td width="30%" colspan="3" align="center"> <?php echo $label_bln;?> </td>
                                                </tr>
       
                                                <tr class="active">
                                                <td width="10%"> <?php echo $label_tgl;?> </td>
                                                <td width="10%"> <?php echo $label_tgl_min1;?> </td>
                                                <td width="10%"> Var </td>
                                                <td width="10%"> <?php echo $prev_date;?> </td>
                                                <td width="10%"> Var MTD </td>
                                                <td width="10%"> Actual </td>
                                                <td width="10%"> Budget </td>
                                                <td width="10%"> Var </td>
                                                </tr>

                                                <tr class="warning">
                                                <td width="20%"> Cash </td>
                                                <td width="10%"><?=nice_number($cash)?></td>
                                                <td width="10%"><?=nice_number($cash3)?></td>
                                                <td width="10%"><?=nice_number($cash5)?></td>
                                                <td width="10%"><?=nice_number($cash4)?></td>
                                                <td width="10%"><?=nice_number($cash6)?></td>
                                                <td width="10%"><?=nice_number($cash)?></td>
                                                <td width="10%"><?=numberFormat($budget_cash)?></td>
                                                <td width="10%"><?=nice_number($cash7)?></td>
                                                </tr>
               
                                                <tr class="warning">
                                                <td width="20%"> Current account - Bank Indonesia </td>
                                                <td width="10%"><?=nice_number($current_account_bi)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi3)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi5)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi4)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi6)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi)?></td>
                                                <td width="10%"><?=numberFormat($budget_current_account_bi)?></td>
                                                <td width="10%"><?=nice_number($current_account_bi7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Certificate of bank Indonesia (SBI & BI call money)     </td>
                                                <td width="10%"><?=nice_number($certificate_bi)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi3)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi5)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi4)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi6)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi)?></td>
                                                <td width="10%"><?=numberFormat($budget_certificate_bi)?></td>
                                                <td width="10%"><?=nice_number($certificate_bi7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Interbank Placement </td>
                                                <td width="10%"><?=nice_number($interbank_placement)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement3)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement5)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement4)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement6)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement)?></td>
                                                <td width="10%"><?=numberFormat($budget_interbank_placement)?></td>
                                                <td width="10%"><?=nice_number($interbank_placement7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Securities </td>
                                                <td width="10%"><?=nice_number($scurities)?></td>
                                                <td width="10%"><?=nice_number($scurities3)?></td>
                                                <td width="10%"><?=nice_number($scurities5)?></td>
                                                <td width="10%"><?=nice_number($scurities4)?></td>
                                                <td width="10%"><?=nice_number($scurities6)?></td>
                                                <td width="10%"><?=nice_number($scurities)?></td>
                                                <td width="10%"><?=numberFormat($budget_scurities)?></td>
                                                <td width="10%"><?=nice_number($scurities7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> -  Allowance For Securities </td>
                                                <td width="10%"><?=nice_number($allowence_fs)?></td>
                                                <td width="10%"><?=nice_number($allowence_fs3)?></td>
                                                <td width="10%"><?=nice_number($allowence_fs5)?></td>
                                                <td width="10%"><?=nice_number($allowence_fs4)?></td>
                                                <td width="10%"><?=nice_number($allowence_fs6)?></td>
                                                <td width="10%"><?=nice_number($allowence_fs)?></td>
                                                <td width="10%"></td>
                                                <td width="10%"><?=nice_number($allowence_fs7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Loans </td>
                                                <td width="10%"><?=nice_number($loans)?></td>
                                                <td width="10%"><?=nice_number($loans3)?></td>
                                                <td width="10%"><?=nice_number($loans5)?></td>
                                                <td width="10%"><?=nice_number($loans4)?></td>
                                                <td width="10%"><?=nice_number($loans6)?></td>
                                                <td width="10%"><?=nice_number($loans)?></td>
                                                <td width="10%"><?=numberFormat($budget_loans)?></td>
                                                <td width="10%"><?=nice_number($loans7)?></td>
                                                </tr>
 
                                                <tr class="warning">
                                                <td width="20%"> -  Performing Loan </td>
                                                <td width="10%"><?=nice_number($performing_loan)?></td>
                                                <td width="10%"><?=nice_number($performing_loan3)?></td>
                                                <td width="10%"><?=nice_number($performing_loan5)?></td>
                                                <td width="10%"><?=nice_number($performing_loan4)?></td>
                                                <td width="10%"><?=nice_number($performing_loan6)?></td>
                                                <td width="10%"><?=nice_number($performing_loan)?></td>
                                                <td width="10%"></td>
                                                <td width="10%"><?=nice_number($performing_loan7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> -  Non Performing Loan*) </td>
                                               <td width="10%"><?=nice_number($non_performing_loan)?></td>
                                                <td width="10%"><?=nice_number($non_performing_loan3)?></td>
                                                <td width="10%"><?=nice_number($non_performing_loan5)?></td>
                                                <td width="10%"><?=nice_number($non_performing_loan4)?></td>
                                                <td width="10%"><?=nice_number($non_performing_loan6)?></td>
                                                <td width="10%"><?=nice_number($non_performing_loan)?></td>
                                                <td width="10%"></td>
                                                <td width="10%"><?=nice_number($non_performing_loan7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> -  Allowance For Loan </td>
                                                <td width="10%"><?=nice_number($allowence_for_loan)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan3)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan5)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan4)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan6)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan)?></td>
                                                <td width="10%"><?=numberFormat($budget_allowence_for_loan)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_loan7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Acceptance receivables  </td>
                                                <td width="10%"><?=nice_number($acceptance_recevables)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables3)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables5)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables4)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables6)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables)?></td>
                                                <td width="10%"><?=numberFormat($budget_acceptance_recevables)?></td>
                                                <td width="10%"><?=nice_number($acceptance_recevables7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Derivative receivables </td>
                                                <td width="10%"><?=nice_number($deferred_receivables)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables3)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables5)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables4)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables6)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables)?></td>
                                                <td width="10%"><?=numberFormat($budget_deferred_receivables)?></td>
                                                <td width="10%"><?=nice_number($deferred_receivables7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Fixed assets (Property, Plant Equipment) </td>
                                                <td width="10%"><?=nice_number($fixed_assets)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets3)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets5)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets4)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets6)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets)?></td>
                                                <td width="10%"><?=numberFormat($budget_fixed_assets)?></td>
                                                <td width="10%"><?=nice_number($fixed_assets7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> Deferred taxes  </td>
                                                <td width="10%"><?=nice_number($deferred_taxes)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes3)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes5)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes4)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes6)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes)?></td>
                                                <td width="10%"><?=numberFormat($budget_deferred_taxes)?></td>
                                                <td width="10%"><?=nice_number($deferred_taxes7)?></td>
                                                </tr>


<?php
#total others assets ------------------
$dash_others_assets=$foreclose_properties+$allowence_for_fp+$account_receivable+$others_assets+$allowence_fsa;
$dash_others_assets3=$foreclose_properties3+$allowence_for_fp3+$account_receivable3+$others_assets3+$allowence_fsa3;
$dash_others_assets5=$foreclose_properties5+$allowence_for_fp5+$account_receivable5+$others_assets5+$allowence_fsa5;
$dash_others_assets4=$foreclose_properties4+$allowence_for_fp4+$account_receivable4+$others_assets4+$allowence_fsa4;
$dash_others_assets6=$foreclose_properties6+$allowence_for_fp6+$account_receivable6+$others_assets6+$allowence_fsa6;
$budget_dash_others_assets=$budget_foreclose_properties+$budget_allowence_for_fp+$budget_account_receivable+$budget_others_assets+$budget_allowence_fsa;
$dash_others_assets7=$foreclose_properties7+$allowence_for_fp7+$account_receivable7+$others_assets7+$allowence_fsa7;

?>

                                                <tr class="warning">
                                                <td width="20%"> <b>Other assets  </b></td>
                                                <td width="10%"><?=nice_number($dash_others_assets)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets3)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets5)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets4)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets6)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_others_assets)?></td>
                                                <td width="10%"><?=nice_number($dash_others_assets7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> -  Foreclosed properties  </td>
                                                <td width="10%"><?=nice_number($foreclose_properties)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties3)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties5)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties4)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties6)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties)?></td>
                                                <td width="10%"><?=numberFormat($budget_foreclose_properties)?></td>
                                                <td width="10%"><?=nice_number($foreclose_properties7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%"> -  Allowance For Foreclosed properties  </td>
                                                <td width="10%"><?=nice_number($allowence_for_fp)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp3)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp5)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp4)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp6)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp)?></td>
                                                <td width="10%"><?=numberFormat($budget_allowence_for_fp)?></td>
                                                <td width="10%"><?=nice_number($allowence_for_fp7)?></td>
                                                </tr>

                                                <tr class="warning">
                                                <td width="20%">- Account receivable  </td>
                                                <td width="10%"><?=nice_number($account_receivable)?></td>
                                                <td width="10%"><?=nice_number($account_receivable3)?></td>
                                                <td width="10%"><?=nice_number($account_receivable5)?></td>
                                                <td width="10%"><?=nice_number($account_receivable4)?></td>
                                                <td width="10%"><?=nice_number($account_receivable6)?></td>
                                                <td width="10%"><?=nice_number($account_receivable)?></td>
                                                <td width="10%"><?=numberFormat($budget_account_receivable)?></td>
                                                <td width="10%"><?=nice_number($account_receivable7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%">- Others  </td>
                                                <td width="10%"><?=nice_number($others_assets)?></td>
                                                <td width="10%"><?=nice_number($others_assets3)?></td>
                                                <td width="10%"><?=nice_number($others_assets5)?></td>
                                                <td width="10%"><?=nice_number($others_assets4)?></td>
                                                <td width="10%"><?=nice_number($others_assets6)?></td>
                                                <td width="10%"><?=nice_number($others_assets)?></td>
                                                <td width="10%"><?=numberFormat($budget_others_assets)?></td>
                                                <td width="10%"><?=nice_number($others_assets7)?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="20%">- Allowances For Suspence Account   </td>
                                                <td width="10%"><?=nice_number($allowence_fsa)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa3)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa5)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa4)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa6)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa)?></td>
                                                <td width="10%"><?=numberFormat($budget_allowence_fsa)?></td>
                                                <td width="10%"><?=nice_number($allowence_fsa7)?></td>
                                                </tr>
                                                <?php
$dash_total_assets=$cash+$current_account_bi+$certificate_bi+$interbank_placement+$scurities+$allowence_fs+$loans+$performing_loan+$non_performing_loan+$allowence_for_loan+$acceptance_recevables+$deferred_receivables+$fixed_assets+$deferred_taxes+$dash_others_assets;
$dash_total_assets3=$cash3+$current_account_bi3+$certificate_bi3+$interbank_placement3+$scurities3+$allowence_fs3+$loans3+$performing_loan3+$non_performing_loan3+$allowence_for_loan3+$acceptance_recevables3+$deferred_receivables3+$fixed_assets3+$deferred_taxes3+$dash_others_assets3;
$dash_total_assets5=$cash5+$current_account_bi5+$certificate_bi5+$interbank_placement5+$scurities5+$allowence_fs5+$loans5+$performing_loan5+$non_performing_loan5+$allowence_for_loan5+$acceptance_recevables5+$deferred_receivables5+$fixed_assets5+$deferred_taxes5+$dash_others_assets5;
$dash_total_assets4=$cash4+$current_account_bi4+$certificate_bi4+$interbank_placement4+$scurities4+$allowence_fs4+$loans4+$performing_loan4+$non_performing_loan4+$allowence_for_loan4+$acceptance_recevables4+$deferred_receivables4+$fixed_assets4+$deferred_taxes4+$dash_others_assets4;
$dash_total_assets6=$cash6+$current_account_bi6+$certificate_bi6+$interbank_placement6+$scurities6+$allowence_fs6+$loans6+$performing_loan6+$non_performing_loan6+$allowence_for_loan6+$acceptance_recevables6+$deferred_receivables6+$fixed_assets6+$deferred_taxes6+$dash_others_assets6;
$budget_dash_total_assets=$budget_cash+$budget_current_account_bi+$budget_certificate_bi+$budget_interbank_placement+$budget_scurities+$budget_allowence_fs+$budget_loans+$budget_performing_loan+$budget_non_performing_loan+$budget_allowence_for_loan+$budget_acceptance_recevables+$budget_deferred_receivables+$budget_fixed_assets+$budget_deferred_taxes+$budget_dash_others_assets;
$dash_total_assets7=$cash7+$current_account_bi7+$certificate_bi7+$interbank_placement7+$scurities7+$allowence_fs7+$loans7+$performing_loan7+$non_performing_loan7+$allowence_for_loan7+$acceptance_recevables7+$deferred_receivables7+$fixed_assets7+$deferred_taxes7+$dash_others_assets7;
                                                ?>
                                                <tr class="warning">
                                                <td width="20%"><b>TOTAL ASSETS </b>   </td>
                                                <td width="10%"><?=nice_number($dash_total_assets)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets3)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets5)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets4)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets6)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_total_assets)?></td>
                                                <td width="10%"><?=nice_number($dash_total_assets7)?></td>
                                                </tr>
                                                </tbody>
                                            </table>
                                            <br>
                                             <table class="table table-bordered table-hover" width="100%">
                                                <thead>
                                                
                                                </thead>
                                                <tbody>
                                                <tr class="active">
                                                <td width="20%" rowspan="2" align="center"> Account of Liabilities & Equity </td>
                                                <td width="50%" colspan="5" align="center">For The Month</td>
                                                <td width="30%" colspan="3" align="center"> <?php echo $label_bln;?> </td>
                                                </tr>
       
                                                <tr class="active">
                                                <td width="10%"> <?php echo $label_tgl;?> </td>
                                                <td width="10%"> <?php echo $label_tgl_min1;?> </td>
                                                <td width="10%"> Var </td>
                                                <td width="10%"> <?php echo $prev_date;?> </td>
                                                <td width="10%"> Var MTD </td>
                                                <td width="10%"> Actual </td>
                                                <td width="10%"> Budget </td>
                                                <td width="10%"> Var </td>
                                                </tr>

                                                <tr class="warning">
                                                <td width="20%">Current Account   </td>
                                                <td width="10%"><?=nice_number($current_account)?></td>
                                                <td width="10%"><?=nice_number($current_account3)?></td>
                                                <td width="10%"><?=nice_number($current_account5)?></td>
                                                <td width="10%"><?=nice_number($current_account4)?></td>
                                                <td width="10%"><?=nice_number($current_account6)?></td>
                                                <td width="10%"><?=nice_number($current_account)?></td>
                                                <td width="10%"><?=numberFormat($budget_current_account)?></td>
                                                <td width="10%"><?=nice_number($current_account7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">Saving Deposits   </td>
                                                <td width="10%"><?=nice_number($saving_deposits)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits3)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits5)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits4)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits6)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits)?></td>
                                                <td width="10%"><?=numberFormat($budget_saving_deposits)?></td>
                                                <td width="10%"><?=nice_number($saving_deposits7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">Time Deposits   </td>
                                                <td width="10%"><?=nice_number($time_deposits)?></td>
                                                <td width="10%"><?=nice_number($time_deposits3)?></td>
                                                <td width="10%"><?=nice_number($time_deposits5)?></td>
                                                <td width="10%"><?=nice_number($time_deposits4)?></td>
                                                <td width="10%"><?=nice_number($time_deposits6)?></td>
                                                <td width="10%"><?=nice_number($time_deposits)?></td>
                                                <td width="10%"><?=numberFormat($budget_time_deposits)?></td>
                                                <td width="10%"><?=nice_number($time_deposits7)?></td>
                                                </tr><?php
 $dash_total_deposits=$current_account+$saving_deposits+$time_deposits;
 $dash_total_deposits3=$current_account3+$saving_deposits3+$time_deposits3;
 $dash_total_deposits5=$current_account5+$saving_deposits5+$time_deposits5;
 $dash_total_deposits4=$current_account4+$saving_deposits4+$time_deposits4;
 $dash_total_deposits6=$current_account6+$saving_deposits6+$time_deposits6;
 $dash_total_deposits7=$current_account7+$saving_deposits7+$time_deposits7;
 $budget_dash_total_deposits=$budget_current_account+$budget_saving_deposits+$budget_time_deposits;                                               

                                                ?>
                                               
<tr class="warning">
                                                <td width="20%">Total deposits   </td>
                                                <td width="10%"><?=nice_number($dash_total_deposits)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits3)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits5)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits4)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits6)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_total_deposits)?></td>
                                                <td width="10%"><?=nice_number($dash_total_deposits7)?></td>
                                                </tr>
                                                <?php
$dash_interbank=$call_money+$bank_deposit+$current_account+$saving_account;
$dash_interbank3=$call_money3+$bank_deposit3+$current_account3+$saving_account3;
$dash_interbank5=$call_money5+$bank_deposit5+$current_account5+$saving_account5;
$dash_interbank4=$call_money4+$bank_deposit4+$current_account4+$saving_account4;
$dash_interbank6=$call_money6+$bank_deposit6+$current_account6+$saving_account6;
$dash_interbank7=$call_money7+$bank_deposit7+$current_account7+$saving_account7;
$budget_dash_interbank=$budget_call_money+$budget_bank_deposit+$budget_current_account+$budget_saving_account;

                                                ?>
                                                

<tr class="warning">
                                                <td width="20%">Interbank   </td>
                                                <td width="10%"><?=nice_number($dash_interbank)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank3)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank5)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank4)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank6)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_interbank)?></td>
                                                <td width="10%"><?=nice_number($dash_interbank7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">- Call Money   </td>
                                                <td width="10%"><?=nice_number($call_money)?></td>
                                                <td width="10%"><?=nice_number($call_money3)?></td>
                                                <td width="10%"><?=nice_number($call_money5)?></td>
                                                <td width="10%"><?=nice_number($call_money4)?></td>
                                                <td width="10%"><?=nice_number($call_money6)?></td>
                                                <td width="10%"><?=nice_number($call_money)?></td>
                                                <td width="10%"><?=numberFormat($budget_call_money)?></td>
                                                <td width="10%"><?=nice_number($acall_money7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">- Bank Deposits   </td>
                                                <td width="10%"><?=nice_number($bank_deposit)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit3)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit5)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit4)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit6)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit)?></td>
                                                <td width="10%"><?=numberFormat($budget_bank_deposit)?></td>
                                                <td width="10%"><?=nice_number($bank_deposit7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">- Current account   </td>
                                                <td width="10%"><?=nice_number($current_account)?></td>
                                                <td width="10%"><?=nice_number($current_account3)?></td>
                                                <td width="10%"><?=nice_number($current_account5)?></td>
                                                <td width="10%"><?=nice_number($current_account4)?></td>
                                                <td width="10%"><?=nice_number($current_account6)?></td>
                                                <td width="10%"><?=nice_number($current_account)?></td>
                                                <td width="10%"><?=numberFormat($budget_current_account)?></td>
                                                <td width="10%"><?=nice_number($current_account7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">- Saving Account   </td>
                                                <td width="10%"><?=nice_number($saving_account)?></td>
                                                <td width="10%"><?=nice_number($saving_account3)?></td>
                                                <td width="10%"><?=nice_number($saving_account5)?></td>
                                                <td width="10%"><?=nice_number($saving_account4)?></td>
                                                <td width="10%"><?=nice_number($saving_account6)?></td>
                                                <td width="10%"><?=nice_number($saving_account)?></td>
                                                <td width="10%"><?=numberFormat($budget_saving_account)?></td>
                                                <td width="10%"><?=nice_number($saving_account7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">Derivative payable   </td>
                                                <td width="10%"><?=nice_number($derivative_payable)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable3)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable5)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable4)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable6)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable)?></td>
                                                <td width="10%"><?=numberFormat($budget_derivative_payable)?></td>
                                                <td width="10%"><?=nice_number($derivative_payable7)?></td>
                                                </tr>
<tr class="warning">
                                                <td width="20%">Acceptance payable</td>
                                                <td width="10%"><?=nice_number($acceptance_payable)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable3)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable5)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable4)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable6)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable)?></td>
                                                <td width="10%"><?=numberFormat($budget_acceptance_payable)?></td>
                                                <td width="10%"><?=nice_number($acceptance_payable7)?></td>
                                                </tr>



<tr class="warning">
                                                <td width="20%">KLBI Payable   </td>
                                                <td width="10%"><?=nice_number($klbi_payable)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable3)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable5)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable4)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable6)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable)?></td>
                                                <td width="10%"><?=numberFormat($budget_klbi_payable)?></td>
                                                <td width="10%"><?=nice_number($klbi_payable7)?></td>
                                                </tr>



<tr class="warning">
                                                <td width="20%">Mandatory Convertible Bonds   </td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds3)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds5)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds4)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds6)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds)?></td>
                                                <td width="10%"><?=numberFormat($budget_mandatory_convertible_bonds)?></td>
                                                <td width="10%"><?=nice_number($mandatory_convertible_bonds7)?></td>
                                                </tr>

<tr class="warning">
                                                <td width="20%">Securities sold with agreement to repurchase  </td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr3)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr5)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr4)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr6)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr)?></td>
                                                <td width="10%"><?=numberFormat($budget_scurities_sold_watr)?></td>
                                                <td width="10%"><?=nice_number($scurities_sold_watr7)?></td>
                                                </tr>
<!-- batas ssssssssssssssssssssssss  48 -->







<tr class="warning">
                                                <td width="20%">Others   </td>
                                                <td width="10%"><?=nice_number($others_liabilities)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities3)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities5)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities4)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities6)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities)?></td>
                                                <td width="10%"><?=numberFormat($budget_others_liabilities)?></td>
                                                <td width="10%"><?=nice_number($others_liabilities7)?></td>
                                                </tr>
                                                <!-- call money sd others -->
                                              
<?php
$dash_tol=$call_money+$bank_deposit+$current_account+$saving_account+$derivative_payable+$acceptance_payable+$klbi_payable+$mandatory_convertible_bonds+$scurities_sold_watr+$others_liabilities;
$dash_tol3=$call_money3+$bank_deposit3+$current_account3+$saving_account3+$derivative_payable3+$acceptance_payable3+$klbi_payable3+$mandatory_convertible_bonds3+$scurities_sold_watr3+$others_liabilities3;
$dash_tol5=$call_money5+$bank_deposit5+$current_account5+$saving_account5+$derivative_payable5+$acceptance_payable5+$klbi_payable5+$mandatory_convertible_bonds5+$scurities_sold_watr5+$others_liabilities5;
$dash_tol4=$call_money4+$bank_deposit4+$current_account4+$saving_account4+$derivative_payable4+$acceptance_payable4+$klbi_payable4+$mandatory_convertible_bonds4+$scurities_sold_watr4+$others_liabilities4;
$dash_tol6=$call_money6+$bank_deposit6+$current_account6+$saving_account6+$derivative_payable6+$acceptance_payable6+$klbi_payable6+$mandatory_convertible_bonds6+$scurities_sold_watr6+$others_liabilities6;
$dash_tol7=$call_money7+$bank_deposit7+$current_account7+$saving_account7+$derivative_payable7+$acceptance_payable7+$klbi_payable7+$mandatory_convertible_bonds7+$scurities_sold_watr7+$others_liabilities7;
$budget_dash_tol=$budget_call_money+$budget_bank_deposit+$budget_current_account+$budget_saving_account+$budget_derivative_payable+$budget_acceptance_payable+$budget_klbi_payable+$budget_mandatory_convertible_bonds+$budget_scurities_sold_watr+$budget_others_liabilities;

?>

<tr class="warning">
                                                <td width="20%">Total Other Liabilities   </td>
                                                <td width="10%"><?=nice_number($dash_tol)?></td>
                                                <td width="10%"><?=nice_number($dash_tol3)?></td>
                                                <td width="10%"><?=nice_number($dash_tol5)?></td>
                                                <td width="10%"><?=nice_number($dash_tol4)?></td>
                                                <td width="10%"><?=nice_number($dash_tol6)?></td>
                                                <td width="10%"><?=nice_number($dash_tol)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_tol)?></td>
                                                <td width="10%"><?=nice_number($dash_tol7)?></td>
                                                </tr>

<tr class="warning">
                                                <td width="20%">Paid in capital   </td>
                                                <td width="10%"><?=nice_number($paid_in_capital)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital3)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital5)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital4)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital6)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital)?></td>
                                                <td width="10%"><?=numberFormat($budget_paid_in_capital)?></td>
                                                <td width="10%"><?=nice_number($paid_in_capital7)?></td>
                                                </tr>





<tr class="warning">
                                                <td width="20%">Agio ( disagio)   </td>
                                                <td width="10%"><?=nice_number($agio_disagio)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio3)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio5)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio4)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio6)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio)?></td>
                                                <td width="10%"><?=numberFormat($budget_agio_disagio)?></td>
                                                <td width="10%"><?=nice_number($agio_disagio7)?></td>
                                                </tr>

<tr class="warning">
                                                <td width="20%">General reserve   </td>
                                                <td width="10%"><?=nice_number($general_reserve)?></td>
                                                <td width="10%"><?=nice_number($general_reserve3)?></td>
                                                <td width="10%"><?=nice_number($general_reserve5)?></td>
                                                <td width="10%"><?=nice_number($general_reserve4)?></td>
                                                <td width="10%"><?=nice_number($general_reserve6)?></td>
                                                <td width="10%"><?=nice_number($general_reserve)?></td>
                                                <td width="10%"><?=numberFormat($budget_general_reserve)?></td>
                                                <td width="10%"><?=nice_number($general_reserve7)?></td>
                                                </tr>

<tr class="warning">
                                                <td width="20%">Available for sale securities - net   </td>
                                                <td width="10%"><?=nice_number($available_fss_net)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net3)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net5)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net4)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net6)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net)?></td>
                                                <td width="10%"><?=numberFormat($budget_available_fss_net)?></td>
                                                <td width="10%"><?=nice_number($available_fss_net7)?></td>
                                                </tr>




<tr class="warning">
                                                <td width="20%">Retained earnings   </td>
                                                <td width="10%"><?=nice_number($retained_earning)?></td>
                                                <td width="10%"><?=nice_number($retained_earning3)?></td>
                                                <td width="10%"><?=nice_number($retained_earning5)?></td>
                                                <td width="10%"><?=nice_number($retained_earning4)?></td>
                                                <td width="10%"><?=nice_number($retained_earning6)?></td>
                                                <td width="10%"><?=nice_number($retained_earning)?></td>
                                                <td width="10%"><?=numberFormat($budget_retained_earning)?></td>
                                                <td width="10%"><?=nice_number($retained_earning7)?></td>
                                                </tr>
                                               

<tr class="warning">
                                                <td width="20%">Profit/loss current year   </td>
                                                <td width="10%"><?=nice_number($c55)?></td>
                                                <td width="10%"><?=nice_number($d55)?></td>
                                                <td width="10%"><?=nice_number($profit_los5)?></td>
                                                <td width="10%"><?=nice_number($profit_los4)?></td>
                                                <td width="10%"><?=nice_number($profit_los6)?></td>
                                                <td width="10%"><?=nice_number($c55)?></td>
                                                <td width="10%"><?=numberFormat($budget_profit_los)?></td>
                                                <td width="10%"><?=nice_number($profit_los7)?></td>
                                                </tr>
<?php
$dah_equity=$paid_in_capital+$agio_disagio+$general_reserve+$available_fss_net+$retained_earning+$c55;
$dah_equity3=$paid_in_capital3+$agio_disagio3+$general_reserve3+$available_fss_net3+$retained_earning3+$d55;
$dah_equity5=$paid_in_capital5+$agio_disagio5+$general_reserve5+$available_fss_net5+$retained_earning5+$profit_los5;
$dah_equity4=$paid_in_capital4+$agio_disagio4+$general_reserve4+$available_fss_net4+$retained_earning4+$profit_los4;
$dah_equity6=$paid_in_capital6+$agio_disagio6+$general_reserve6+$available_fss_net6+$retained_earning6+$profit_los6;
$dah_equity7=$paid_in_capital7+$agio_disagio7+$general_reserve7+$available_fss_net7+$retained_earning7+$profit_los7;
$budget_dah_equity=$budget_paid_in_capital+$budget_agio_disagio+$budget_general_reserve+$budget_available_fss_net+$budget_retained_earning+$budget_profit_los;


?>
                                                <!-- Paid in capital s/d Profit/loss current year
                                                $objPHPExcel->getActiveSheet()->setCellValue('B56', 'Total Equity');
$objPHPExcel->getActiveSheet()->setCellValue('C56', '=SUM(C50:C55)');
$objPHPExcel->getActiveSheet()->setCellValue('D56', '=SUM(D50:D55)');
$objPHPExcel->getActiveSheet()->setCellValue('E56', '=SUM(E50:E55)');
$objPHPExcel->getActiveSheet()->setCellValue('F56', '=SUM(F50:F55)');
$objPHPExcel->getActiveSheet()->setCellValue('G56', '=SUM(G50:G55)');
$objPHPExcel->getActiveSheet()->setCellValue('H56', '=SUM(H50:H55)');
$objPHPExcel->getActiveSheet()->setCellValue('I56', '=SUM(I50:I55)');
$objPHPExcel->getActiveSheet()->setCellValue('J56', '=SUM(J50:J55)');
-->
<tr class="warning">
                                                <td width="20%">Total Equity  </td>
                                                <td width="10%"><?=nice_number($dah_equity)?></td>
                                                <td width="10%"><?=nice_number($dah_equity3)?></td>
                                                <td width="10%"><?=nice_number($dah_equity5)?></td>
                                                <td width="10%"><?=nice_number($dah_equity4)?></td>
                                                <td width="10%"><?=nice_number($dah_equity6)?></td>
                                                <td width="10%"><?=nice_number($dah_equity)?></td>
                                                <td width="10%"><?=numberFormat($budget_dah_equity)?></td>
                                                <td width="10%"><?=nice_number($dah_equity7)?></td>
                                                </tr>
                                                <?php
$dash_tle=$dash_total_deposits+$dash_tol+$dah_equity;
$dash_tle3=$dash_total_deposits3+$dash_tol3+$dah_equity3;
$dash_tle5=$dash_total_deposits5+$dash_tol5+$dah_equity5;
$dash_tle4=$dash_total_deposits4+$dash_tol4+$dah_equity4;
$dash_tle6=$dash_total_deposits6+$dash_tol6+$dah_equity6;
$dash_tle7=$dash_total_deposits7+$dash_tol7+$dah_equity7;
$budget_dash_tle=$budget_dash_total_deposits+$budget_dash_tol+$budget_dah_equity;


                                                ?>
<!--                                                $objPHPExcel->getActiveSheet()->setCellValue('B57', 'TOTAL LIABILITIES & EQUITY');
$objPHPExcel->getActiveSheet()->setCellValue('C57', '=C37+C49+C56');
$objPHPExcel->getActiveSheet()->setCellValue('D57', '=D37+D49+D56');
$objPHPExcel->getActiveSheet()->setCellValue('E57', '=E37+E49+E56');
$objPHPExcel->getActiveSheet()->setCellValue('F57', '=F37+F49+F56');
$objPHPExcel->getActiveSheet()->setCellValue('G57', '=G37+G49+G56');
$objPHPExcel->getActiveSheet()->setCellValue('H57', '=H37+H49+H56');
$objPHPExcel->getActiveSheet()->setCellValue('I57', '=I37+I49+I56');
$objPHPExcel->getActiveSheet()->setCellValue('J57', '=J37+J49+J56');
-->
<tr class="warning">
                                                <td width="20%">TOTAL LIABILITIES & EQUITY  </td>
                                                <td width="10%"><?=nice_number($dash_tle)?></td>
                                                <td width="10%"><?=nice_number($dash_tle3)?></td>
                                                <td width="10%"><?=nice_number($dash_tle5)?></td>
                                                <td width="10%"><?=nice_number($dash_tle4)?></td>
                                                <td width="10%"><?=nice_number($dash_tle6)?></td>
                                                <td width="10%"><?=nice_number($dash_tle)?></td>
                                                <td width="10%"><?=numberFormat($budget_dash_tle)?></td>
                                                <td width="10%"><?=nice_number($dash_tle7)?></td>
                                                </tr>
                                                </tbody>
                                            </table>


                                        </div>
                                        </p>
                                    </div>
                                    <div class="tab-pane" id="tab_15_2">
                                        <!--
										<p>
                                            <h5><b>IS Flash Report</b></h5>
                                        </p>
										-->
                                        <div class="table-scrollable">
                                        <table class="table table-bordered table-hover" width="100%">
                                                <thead>
                                                
                                                </thead>
                                                <tbody>
                                                <tr class="active">
                                                <td width="10%" rowspan="2" align="center"> Description </td>
                                                <td width="45%" colspan="10" align="center">MTD</td>
                                                <td width="30%" colspan="5" align="center"> YTD </td>
                                                <td width="10%" colspan="2" align="center"> Whole Year  Budget </td>
                                                <td width="5%"  align="center"> YTD </td>
                                                </tr>

       
                                                <tr class="active">
                                                <td width="5%"> <?php echo $prev_date;?> </td>
                                                <td width="5%"> Var </td>
                                                <td width="5%"> % </td>
                                                <td width="5%"> <?php echo $label_tgl;?> </td>
                                                <td width="5%"> <?php echo $label_tgl_min1;?> </td>
                                                <td width="5%"> var </td>
                                                <td width="5%"> MTD Forecast <?php echo $label_tgl;?></td>
                                                <td width="5%"> Var MTD v.s Budget </td>
                                                <td width="5%"> % </td>
                                                <td width="5%"> Budget </td>

                                                <td width="5%"> <?php echo $label_tgl;?> </td>
                                                <td width="5%"> Var YTD v.s Budget </td>
                                                <td width="5%"> % </td>
                                                <td width="5%"> YTD Forecast <?php echo $label_tgl;?> </td>
                                                <td width="5%"> Budget </td>

                                                <td width="5%"> Rp </td>
                                                <td width="5%"> %</td>
                                                <td width="5%"> <?php echo $prev_date;?> </td>
                                                </tr>
 <?php
 $objPHPExcel = PHPExcel_IOFactory::load("download/Flash_Report_".$label_tgl."_".$file_eksport.".xls");
 $objWorksheet = $objPHPExcel->getActiveSheet('2');


 ?>                                              
                                                <tr class="warning">
                                                <td width="5%"> <b>Interest Income</b> </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T8')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Loan </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T9')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Treasury bills </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T10')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Interbank placements </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T11')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Placement with BI</td>
<td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T12')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Others </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T13')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

                                                <tr class="warning">
                                                <td width="5%"> <b>Interest Expense Funding </b></td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T14')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Current accounts </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T15')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Saving accounts </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T16')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Time deposits </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T17')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Bank deposits </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T18')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

                                                <tr class="warning">
                                                <td width="5%"> <b>Other Interest Expense </b></td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T19')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Borrowings (MCB) </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T20')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Guaranteed premium </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T21')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Others </td>
                                               <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T22')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Net Interest Income</b> </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T23')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>









 


                                                <tr class="warning">
                                                <td width="5%"> Other Income: </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Forex gain/(loss) on transactions</td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T26')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%">Gain/(Loss) on sale of securities/bonds: </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T27')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Remittance fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T28')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Trade Finance fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T29')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Processing fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T30')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Credit Card fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T31')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Insurance Fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T32')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%">Service Charges </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T33')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Other Commission & Fee </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T34')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Penalty </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T35')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>




                                                 <tr class="warning">
                                                <td width="5%"> Write Offs Recovered </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T36')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                 <tr class="warning">
                                                <td width="5%"> Other Income </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T37')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Total Other Income </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T38')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Operating Revenue </b></td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T39')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>







                                                <tr class="warning">
                                                <td width="5%"> Operating Expenses:</td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Staff Cost </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T42')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> General & Administrative Expenses </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T43')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Depreciation </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T44')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Total Operating Expenses</b> </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T45')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>





  

                                                <tr class="warning">
                                                <td width="5%"> Other Operating Expense/Income </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T46')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Operating Profit</b> </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T47')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> General Provision </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T50')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Specific Provision Charged </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T51')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Specific Provision Recovery </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T52')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                




                                                <tr class="warning">
                                                <td width="5%"> Write Offs Charged </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T53')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Write Offs Recovered </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T54')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Foreclose Properties Provision </td>
                                                <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T55')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Others </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T56')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>





                                                <tr class="warning">
                                                <td width="5%"> Total Provision </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T57')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Extraordinary item </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%">  </td>
                                                <td width="5%"></td>
                                                <td width="5%">  </td>
                                                <td width="5%"> </td>
                                                <td width="5%"> </td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                <td width="5%"></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Profit (Loss) Before Tax</b> </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T59')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> Corporate Tax </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T60')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <tr class="warning">
                                                <td width="5%"> <b>Profit (Loss) After Tax </b> </td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('C61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('D61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('E61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('F61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('G61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('H61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('I61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('J61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('K61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('L61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('M61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('N61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('O61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('P61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('Q61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('R61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('S61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                 <td width="5%"><?php echo $objPHPExcel->getActiveSheet()->getCell('T61')->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                </tbody>
                                                </table>
                                        </div>
                                        
                                    </div>
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

