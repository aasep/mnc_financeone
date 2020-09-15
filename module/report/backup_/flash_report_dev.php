<?php
//require_once 'config/config.php';
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

//error_reporting(0);
$tanggal=$_POST['tanggal']; 

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

$curr_tgl=date('Y-m-d',strtotime($tanggal));
$curr_tgl_min1=date('Y-m-d',strtotime($label_tgl_min1));
$curr_mon_min1=date('Y-m-d',strtotime($prev_date));

//$curr_day_budget=
//$curr_mon_budget=
//query budget
$query_budget=" select budget from Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget' ";

//=======================================
$budget=0;
//==============hardcode

$var_curr_tgl="  c.DataDate='".$curr_tgl."' ";//var tgl terpilih
$var_curr_tgl_min1="  c.DataDate='".$curr_tgl_min1."' ";//var tgl terpilih minus 1
$var_curr_mon_min1="  c.DataDate='".$curr_mon_min1."' ";//var tgl terakhir bulan sebelumnya
//and a.FLASH_Level_3_Description='Cash'
$query_currentDate="select sum (total) as jml_nominal from (
select DISTINCT(c.KodeGL),c.Nominal as total
from DM_Journal c
JOIN Referensi_GL_02_New b ON c.KodeGL = b.GLNO 
JOIN Referensi_Flash_Report a ON a.FLASH_LEVEL_3 = b.FLASH_LEVEL_3
WHERE  ";


$var_flash_add=" ) as tabel1 ";



//$query_currentDate="select sum (c.Nominal) as jml_nominal from Referensi_Flash_Report a, GL_02_Baru b, DM_Journal c where a.FLASH_Level_3=b.FLASH_LEVEL_3 and b.GLNO=c.KodeGL  ";



//========================================
//$query="select distinct FLASH_Level_3_Description from Referensi_Flash_Report  order by FLASH_Level_3_Description asc";
//$result=odbc_exec($connection, $query);
//$jsonData = array();
//$i=1;
//while ($row = odbc_fetch_array($result)) {
    //$jsonData[] = $array;

//switch ($row['FLASH_Level_3_Description']) {
 //   case "Cash":
    //current_date
        $var_flash=" and a.FLASH_Level_3_Description='Cash' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $cash=$row2['jml_nominal'];
//echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
//die();
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

       // echo $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
       // die();
        //echo $query_budget.$var_budget."<br>";
        //echo  $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add;
        //die();
 //       break;
 //   case "Current account - Bank Indonesia":
        $var_flash=" and a.FLASH_Level_3_Description='Current account - Bank Indonesia' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account_bi=$row2['jml_nominal'];

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
        $var_flash=" and a.FLASH_Level_3_Description='Certificate of bank Indonesia (SBI & BI call money)' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Interbank Placement' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Securities' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $scurities=$row2['jml_nominal'];

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
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Securities' ";
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
   //     break;
  //  case "Loans":

/*        
        $var_flash=" and a.FLASH_Level_3_Description='Loans' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $loans=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $loans3=$row3['jml_nominal'];

        $loans5=$loans-$loans3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $loans4=$row4['jml_nominal'];
        $loans6=$loans-$loans4;

        $var_budget=" and FLASH_Level_3='FLASH101000007' ";
        $result5=odbc_exec($connection2, $query_budget.$var_budget);
        $row5=odbc_fetch_array($result5);
        $budget_loans=$row5['budget'];

        $loans7=$loans-$budget_loans;
    //    break;
   // case "Performing Loan":
        $var_flash=" and a.FLASH_Level_3_Description='Performing Loan' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $performing_loan=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $performing_loan3=$row3['jml_nominal'];

        $performing_loan5=$performing_loan-$performing_loan3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $performing_loan4=$row4['jml_nominal'];
        $performing_loan6=$performing_loan-$performing_loan4;

        $performing_loan7=$performing_loan-$budget;
   //     break;
  //  case "Non Performing Loan*)":
        $var_flash=" and a.FLASH_Level_3_Description='Non Performing Loan*)' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $non_performing_loan=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $non_performing_loan3=$row3['jml_nominal'];

        $non_performing_loan5=$non_performing_loan-$non_performing_loan3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $non_performing_loan4=$row4['jml_nominal'];
        $non_performing_loan6=$non_performing_loan-$non_performing_loan4;


        $non_performing_loan7=$non_performing_loan-$budget;
   //     break;
   // case "Allowance For Loan":
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Loan' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $allowence_for_loan=$row2['jml_nominal'];

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
        $var_flash=" and a.FLASH_Level_3_Description='Acceptance receivables' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Derivative receivables' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Fixed assets (Property, Plant Equipment)' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Deferred taxes' ";
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

        $deferred_taxes7=$deferred_taxes-$budget_deferred_taxes;
  //      break;
  //  case "Others - Assets":
        $var_flash=" and a.FLASH_Level_3_Description='Others - Assets' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_assets=$row2['jml_nominal'];

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $others_assets3=$row3['jml_nominal'];

        $others_assets5=$others_assets-$others_assets3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $others_assets4=$row4['jml_nominal']; 
        $others_assets6=$others_assets-$others_assets4;

        $others_assets7=$others_assets-$budget;
  //      break;
  //  case "Foreclosed properties":
        $var_flash=" and a.FLASH_Level_3_Description='Foreclosed properties' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Foreclosed properties' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Account receivable' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Others - Other assets' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_assets=$row2['jml_nominal'];

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
        $var_flash=" and a.FLASH_Level_3_Description='Allowances For Suspence Account' ";
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
  //  case "Current Account":
        $var_flash=" and a.FLASH_Level_3_Description='Current Account' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $current_account=$row2['jml_nominal']; 

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
   //     break;
   // case "Saving Deposits":
        $var_flash=" and a.FLASH_Level_3_Description='Saving Deposits' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $saving_deposits=$row2['jml_nominal']; 

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
        $var_flash=" and a.FLASH_Level_3_Description='Time deposits' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Interbank' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Call Money' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Bank deposits' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Current Account' ";
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

        $current_account_interbank7=$current_account_interbank-$budget;
  //      break;
//	case "Saving accounts":
        $var_flash=" and a.FLASH_Level_3_Description='Saving accounts' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $saving_account=$row2['jml_nominal']; 

        $result3=odbc_exec($connection2, $query_currentDate.$var_curr_tgl_min1.$var_flash.$var_flash_add);
        $row3=odbc_fetch_array($result3);
        $saving_account3=$row3['jml_nominal']; 

        $saving_account5=$saving_account-$saving_account3;

        $result4=odbc_exec($connection2, $query_currentDate.$var_curr_mon_min1.$var_flash.$var_flash_add);
        $row4=odbc_fetch_array($result4);
        $saving_account4=$row4['jml_nominal'];
        $saving_account6=$saving_account-$saving_account4;

        $saving_account7=$saving_account-$budget;
 //       break;
 //   case "Derivative payable":
        $var_flash=" and a.FLASH_Level_3_Description='Derivative payable' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Acceptance payable' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='KLBI Payable' ";
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

        $klbi_payable7=$klbi_payable-$budget;
//        break;
//    case "Mandatory Convertible Bonds":
        $var_flash=" and a.FLASH_Level_3_Description='Mandatory Convertible Bonds' ";
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

        $mandatory_convertible_bonds7=$mandatory_convertible_bonds-$budget;
 //       break;
 //   case "Securities sold with agreement to repurchase":
        $var_flash=" and a.FLASH_Level_3_Description='Securities sold with agreement to repurchase' ";
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

        $scurities_sold_watr7=$scurities_sold_watr-$budget;
//        break;
//    case "Others Liabilities":
        $var_flash=" and a.FLASH_Level_3_Description='Others Liabilities' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $others_liabilities=$row2['jml_nominal'];

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
        $var_flash=" and a.FLASH_Level_3_Description='Paid in capital' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Agio ( disagio)' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='General reserve' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Available for sale securities - net' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Retained earnings' ";
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
        $var_flash=" and a.FLASH_Level_3_Description='Profit/loss current year' ";
        $result2=odbc_exec($connection2, $query_currentDate.$var_curr_tgl.$var_flash.$var_flash_add);
        $row2=odbc_fetch_array($result2);
        $profit_los=$row2['jml_nominal']; 

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

        $profit_los6=$profit_los-$budget_profit_los;
//        break;

//}



//$i++;
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


*/

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
//BOLD
$objPHPExcel->getActiveSheet()->getStyle('B1:B3')->applyFromArray($styleArrayFontBold);
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
//Bakgroud
//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArraybackgroundRed);
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B60:C60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B63:C63')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');
$objPHPExcel->getActiveSheet()->getStyle('B71:C71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');

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


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:A1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('K1:Z1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B73:C1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:J1000');
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('D60:Z1000');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B1:J1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B2:J2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B3:J3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B4:J4');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B5:B7');//Account of Assets
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C5:G6');//
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H5:J6');//
//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A73:B73');//


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
$objPHPExcel->getActiveSheet()->setCellValue('C8', Milion_format($cash));
$objPHPExcel->getActiveSheet()->setCellValue('D8', Milion_format($cash3));
$objPHPExcel->getActiveSheet()->setCellValue('E8', Milion_format($cash5));
$objPHPExcel->getActiveSheet()->setCellValue('F8', Milion_format($cash4));
$objPHPExcel->getActiveSheet()->setCellValue('G8', Milion_format($cash6));
$objPHPExcel->getActiveSheet()->setCellValue('H8', Milion_format($cash));
$objPHPExcel->getActiveSheet()->setCellValue('I8', Milion_format( $budget_cash));
$objPHPExcel->getActiveSheet()->setCellValue('J8', Milion_format($cash7));
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Current account - Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('C9', Milion_format($current_account_bi));
$objPHPExcel->getActiveSheet()->setCellValue('D9', Milion_format($current_account_bi3));
$objPHPExcel->getActiveSheet()->setCellValue('E9', Milion_format($current_account_bi5));
$objPHPExcel->getActiveSheet()->setCellValue('F9', Milion_format($current_account_bi4));
$objPHPExcel->getActiveSheet()->setCellValue('G9', Milion_format($current_account_bi6));
$objPHPExcel->getActiveSheet()->setCellValue('H9', Milion_format($current_account_bi));
$objPHPExcel->getActiveSheet()->setCellValue('I9', Milion_format($budget_current_account_bi));
$objPHPExcel->getActiveSheet()->setCellValue('J9', Milion_format($current_account_bi7));
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Certificate of bank Indonesia (SBI & BI call money)	');
$objPHPExcel->getActiveSheet()->setCellValue('C10', Milion_format($certificate_bi));
$objPHPExcel->getActiveSheet()->setCellValue('D10', Milion_format($certificate_bi3));
$objPHPExcel->getActiveSheet()->setCellValue('E10', Milion_format($certificate_bi5));
$objPHPExcel->getActiveSheet()->setCellValue('F10', Milion_format($certificate_bi4));
$objPHPExcel->getActiveSheet()->setCellValue('G10', Milion_format($certificate_bi6));
$objPHPExcel->getActiveSheet()->setCellValue('H10', Milion_format($certificate_bi));
$objPHPExcel->getActiveSheet()->setCellValue('I10', Milion_format($budget_certificate_bi));
$objPHPExcel->getActiveSheet()->setCellValue('J10', Milion_format($certificate_bi7));
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Interbank Placement');
$objPHPExcel->getActiveSheet()->setCellValue('C11', Milion_format($interbank_placement));
$objPHPExcel->getActiveSheet()->setCellValue('D11', Milion_format($interbank_placement3));
$objPHPExcel->getActiveSheet()->setCellValue('E11', Milion_format($interbank_placement5));
$objPHPExcel->getActiveSheet()->setCellValue('F11', Milion_format($interbank_placement4));
$objPHPExcel->getActiveSheet()->setCellValue('G11', Milion_format($interbank_placement6));
$objPHPExcel->getActiveSheet()->setCellValue('H11', Milion_format($interbank_placement));
$objPHPExcel->getActiveSheet()->setCellValue('I11', Milion_format($budget_interbank_placement));
$objPHPExcel->getActiveSheet()->setCellValue('J11', Milion_format($interbank_placement7));

$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Securities	');
$objPHPExcel->getActiveSheet()->setCellValue('C12', Milion_format($scurities));
$objPHPExcel->getActiveSheet()->setCellValue('D12', Milion_format($scurities3));
$objPHPExcel->getActiveSheet()->setCellValue('E12', Milion_format($scurities5));
$objPHPExcel->getActiveSheet()->setCellValue('F12', Milion_format($scurities4));
$objPHPExcel->getActiveSheet()->setCellValue('G12', Milion_format($scurities6));
$objPHPExcel->getActiveSheet()->setCellValue('H12', Milion_format($scurities));
$objPHPExcel->getActiveSheet()->setCellValue('I12', Milion_format($budget_scurities));
$objPHPExcel->getActiveSheet()->setCellValue('J12', Milion_format($scurities7));

$objPHPExcel->getActiveSheet()->setCellValue('B13', '-	Allowance For Securities');
$objPHPExcel->getActiveSheet()->setCellValue('C13', Milion_format($allowence_fs));
$objPHPExcel->getActiveSheet()->setCellValue('D13', Milion_format($allowence_fs3));
$objPHPExcel->getActiveSheet()->setCellValue('E13', Milion_format($allowence_fs5));
$objPHPExcel->getActiveSheet()->setCellValue('F13', Milion_format($allowence_fs4));
$objPHPExcel->getActiveSheet()->setCellValue('G13', Milion_format($allowence_fs6));
$objPHPExcel->getActiveSheet()->setCellValue('H13', Milion_format($allowence_fs));
$objPHPExcel->getActiveSheet()->setCellValue('J13', Milion_format($allowence_fs7));

$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Loans');
$objPHPExcel->getActiveSheet()->setCellValue('C14', Milion_format($loans));
$objPHPExcel->getActiveSheet()->setCellValue('D14', Milion_format($loans3));
$objPHPExcel->getActiveSheet()->setCellValue('E14', Milion_format($loans5));
$objPHPExcel->getActiveSheet()->setCellValue('F14', Milion_format($loans4));
$objPHPExcel->getActiveSheet()->setCellValue('G14', Milion_format($loans6));
$objPHPExcel->getActiveSheet()->setCellValue('H14', Milion_format($loans));
$objPHPExcel->getActiveSheet()->setCellValue('I14', Milion_format($budget_loans));
$objPHPExcel->getActiveSheet()->setCellValue('J14', Milion_format($loans7));

$objPHPExcel->getActiveSheet()->setCellValue('B15', '-	Performing Loan');
$objPHPExcel->getActiveSheet()->setCellValue('C15', Milion_format($performing_loan));
$objPHPExcel->getActiveSheet()->setCellValue('D15', Milion_format($performing_loan3));
$objPHPExcel->getActiveSheet()->setCellValue('E15', Milion_format($performing_loan5));
$objPHPExcel->getActiveSheet()->setCellValue('F15', Milion_format($performing_loan4));
$objPHPExcel->getActiveSheet()->setCellValue('G15', Milion_format($performing_loan6));
$objPHPExcel->getActiveSheet()->setCellValue('H15', Milion_format($performing_loan));
$objPHPExcel->getActiveSheet()->setCellValue('J15', Milion_format($performing_loan7));

$objPHPExcel->getActiveSheet()->setCellValue('B16', '-	Non Performing Loan*)	');
$objPHPExcel->getActiveSheet()->setCellValue('C16', Milion_format($non_performing_loan));
$objPHPExcel->getActiveSheet()->setCellValue('D16', Milion_format($non_performing_loan3));
$objPHPExcel->getActiveSheet()->setCellValue('E16', Milion_format($non_performing_loan5));
$objPHPExcel->getActiveSheet()->setCellValue('F16', Milion_format($non_performing_loan4));
$objPHPExcel->getActiveSheet()->setCellValue('G16', Milion_format($non_performing_loan6));
$objPHPExcel->getActiveSheet()->setCellValue('H16', Milion_format($non_performing_loan));
$objPHPExcel->getActiveSheet()->setCellValue('J16', Milion_format($non_performing_loan7));

$objPHPExcel->getActiveSheet()->setCellValue('B17', '-	Allowance For Loan	');
$objPHPExcel->getActiveSheet()->setCellValue('C17', Milion_format($allowence_for_loan));
$objPHPExcel->getActiveSheet()->setCellValue('D17', Milion_format($allowence_for_loan3));
$objPHPExcel->getActiveSheet()->setCellValue('E17', Milion_format($allowence_for_loan5));
$objPHPExcel->getActiveSheet()->setCellValue('F17', Milion_format($allowence_for_loan4));
$objPHPExcel->getActiveSheet()->setCellValue('G17', Milion_format($allowence_for_loan6));
$objPHPExcel->getActiveSheet()->setCellValue('H17', Milion_format($allowence_for_loan));
$objPHPExcel->getActiveSheet()->setCellValue('I17', Milion_format($budget_allowence_for_loan));
$objPHPExcel->getActiveSheet()->setCellValue('J17', Milion_format($allowence_for_loan7));

$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Acceptance receivables	');
$objPHPExcel->getActiveSheet()->setCellValue('C18', Milion_format($acceptance_recevables));
$objPHPExcel->getActiveSheet()->setCellValue('D18', Milion_format($acceptance_recevables3));
$objPHPExcel->getActiveSheet()->setCellValue('E18', Milion_format($acceptance_recevables5));
$objPHPExcel->getActiveSheet()->setCellValue('F18', Milion_format($acceptance_recevables4));
$objPHPExcel->getActiveSheet()->setCellValue('G18', Milion_format($acceptance_recevables6));
$objPHPExcel->getActiveSheet()->setCellValue('H18', Milion_format($acceptance_recevables));
$objPHPExcel->getActiveSheet()->setCellValue('I18', Milion_format($budget_acceptance_recevables));
$objPHPExcel->getActiveSheet()->setCellValue('J18', Milion_format($acceptance_recevables7));

$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Derivative receivables');
$objPHPExcel->getActiveSheet()->setCellValue('C19', Milion_format($deferred_receivables));
$objPHPExcel->getActiveSheet()->setCellValue('D19', Milion_format($deferred_receivables3));
$objPHPExcel->getActiveSheet()->setCellValue('E19', Milion_format($deferred_receivables5));
$objPHPExcel->getActiveSheet()->setCellValue('F19', Milion_format($deferred_receivables4));
$objPHPExcel->getActiveSheet()->setCellValue('G19', Milion_format($deferred_receivables6));
$objPHPExcel->getActiveSheet()->setCellValue('H19', Milion_format($deferred_receivables));
$objPHPExcel->getActiveSheet()->setCellValue('I19', Milion_format($budget_deferred_receivables));
$objPHPExcel->getActiveSheet()->setCellValue('J19', Milion_format($deferred_receivables7));

$objPHPExcel->getActiveSheet()->setCellValue('B20','Fixed assets (Property, Plant Equipment)');
$objPHPExcel->getActiveSheet()->setCellValue('C20',Milion_format($fixed_assets));
$objPHPExcel->getActiveSheet()->setCellValue('D20',Milion_format($fixed_assets3));
$objPHPExcel->getActiveSheet()->setCellValue('E20',Milion_format($fixed_assets5));
$objPHPExcel->getActiveSheet()->setCellValue('F20',Milion_format($fixed_assets4));
$objPHPExcel->getActiveSheet()->setCellValue('G20',Milion_format($fixed_assets6));
$objPHPExcel->getActiveSheet()->setCellValue('H20',Milion_format($fixed_assets));
$objPHPExcel->getActiveSheet()->setCellValue('I20',Milion_format($budget_fixed_assets));
$objPHPExcel->getActiveSheet()->setCellValue('J20',Milion_format($fixed_assets7));

$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Deferred taxes	');
$objPHPExcel->getActiveSheet()->setCellValue('C21', Milion_format($deferred_taxes));
$objPHPExcel->getActiveSheet()->setCellValue('D21', Milion_format($deferred_taxes3));
$objPHPExcel->getActiveSheet()->setCellValue('E21', Milion_format($deferred_taxes5));
$objPHPExcel->getActiveSheet()->setCellValue('F21', Milion_format($deferred_taxes4));
$objPHPExcel->getActiveSheet()->setCellValue('G21', Milion_format($deferred_taxes6));
$objPHPExcel->getActiveSheet()->setCellValue('H21', Milion_format($deferred_taxes));
$objPHPExcel->getActiveSheet()->setCellValue('I21', Milion_format($budget_deferred_taxes));
$objPHPExcel->getActiveSheet()->setCellValue('J21', Milion_format($deferred_taxes7));

$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Other assets');
$objPHPExcel->getActiveSheet()->setCellValue('C22', Milion_format($others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('D22', Milion_format($others_assets3));
$objPHPExcel->getActiveSheet()->setCellValue('E22', Milion_format($others_assets5));
$objPHPExcel->getActiveSheet()->setCellValue('F22', Milion_format($others_assets4));
$objPHPExcel->getActiveSheet()->setCellValue('G22', Milion_format($others_assets6));
$objPHPExcel->getActiveSheet()->setCellValue('H22', Milion_format($others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('I22', Milion_format($budget_others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('J22', Milion_format($others_assets7));


$objPHPExcel->getActiveSheet()->setCellValue('B23', '-	Foreclosed properties');
$objPHPExcel->getActiveSheet()->setCellValue('C23', Milion_format($foreclose_properties));
$objPHPExcel->getActiveSheet()->setCellValue('D23', Milion_format($foreclose_properties3));
$objPHPExcel->getActiveSheet()->setCellValue('E23', Milion_format($foreclose_properties5));
$objPHPExcel->getActiveSheet()->setCellValue('F23', Milion_format($foreclose_properties4));
$objPHPExcel->getActiveSheet()->setCellValue('G23', Milion_format($foreclose_properties6));
$objPHPExcel->getActiveSheet()->setCellValue('H23', Milion_format($foreclose_properties));
$objPHPExcel->getActiveSheet()->setCellValue('I23', Milion_format($budget_foreclose_properties));
$objPHPExcel->getActiveSheet()->setCellValue('J23', Milion_format($foreclose_properties7));

$objPHPExcel->getActiveSheet()->setCellValue('B24', '- 	Allowance For Foreclosed properties	');
$objPHPExcel->getActiveSheet()->setCellValue('C24', Milion_format($allowence_for_fp));
$objPHPExcel->getActiveSheet()->setCellValue('D24', Milion_format($allowence_for_fp3));
$objPHPExcel->getActiveSheet()->setCellValue('E24', Milion_format($allowence_for_fp5));
$objPHPExcel->getActiveSheet()->setCellValue('F24', Milion_format($allowence_for_fp4));
$objPHPExcel->getActiveSheet()->setCellValue('G24', Milion_format($allowence_for_fp6));
$objPHPExcel->getActiveSheet()->setCellValue('H24', Milion_format($allowence_for_fp));
$objPHPExcel->getActiveSheet()->setCellValue('I24', Milion_format($budget_allowence_for_fp));
$objPHPExcel->getActiveSheet()->setCellValue('J24', Milion_format($allowence_for_fp7));

$objPHPExcel->getActiveSheet()->setCellValue('B25', '-	Account receivable	');
$objPHPExcel->getActiveSheet()->setCellValue('C25', Milion_format($account_receivable));
$objPHPExcel->getActiveSheet()->setCellValue('D25', Milion_format($account_receivable3));
$objPHPExcel->getActiveSheet()->setCellValue('E25', Milion_format($account_receivable5));
$objPHPExcel->getActiveSheet()->setCellValue('F25', Milion_format($account_receivable4));
$objPHPExcel->getActiveSheet()->setCellValue('G25', Milion_format($account_receivable6));
$objPHPExcel->getActiveSheet()->setCellValue('H25', Milion_format($account_receivable));
$objPHPExcel->getActiveSheet()->setCellValue('I25', Milion_format($budget_account_receivable));
$objPHPExcel->getActiveSheet()->setCellValue('J25', Milion_format($account_receivable7));

$objPHPExcel->getActiveSheet()->setCellValue('B26', '-	Others');
$objPHPExcel->getActiveSheet()->setCellValue('C26', Milion_format($others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('D26', Milion_format($others_assets3));
$objPHPExcel->getActiveSheet()->setCellValue('E26', Milion_format($others_assets5));
$objPHPExcel->getActiveSheet()->setCellValue('F26', Milion_format($others_assets4));
$objPHPExcel->getActiveSheet()->setCellValue('G26', Milion_format($others_assets6));
$objPHPExcel->getActiveSheet()->setCellValue('H26', Milion_format($others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('I26', Milion_format($budget_others_assets));
$objPHPExcel->getActiveSheet()->setCellValue('J26', Milion_format($others_assets7));

$objPHPExcel->getActiveSheet()->setCellValue('B27', '-	Allowances For Suspence Account	');
$objPHPExcel->getActiveSheet()->setCellValue('C27', Milion_format($allowence_fsa));
$objPHPExcel->getActiveSheet()->setCellValue('D27', Milion_format($allowence_fsa3));
$objPHPExcel->getActiveSheet()->setCellValue('E27', Milion_format($allowence_fsa5));
$objPHPExcel->getActiveSheet()->setCellValue('F27', Milion_format($allowence_fsa4));
$objPHPExcel->getActiveSheet()->setCellValue('G27', Milion_format($allowence_fsa6));
$objPHPExcel->getActiveSheet()->setCellValue('H27', Milion_format($allowence_fsa));
$objPHPExcel->getActiveSheet()->setCellValue('I27', Milion_format($budget_allowence_fsa));
$objPHPExcel->getActiveSheet()->setCellValue('J27', Milion_format($allowence_fsa7));

$objPHPExcel->getActiveSheet()->setCellValue('B28', 'TOTAL ASSETS');
$objPHPExcel->getActiveSheet()->setCellValue('C28', Milion_format($total_assets_curr));
$objPHPExcel->getActiveSheet()->setCellValue('D28', Milion_format($total_assets_curr_min1));
$objPHPExcel->getActiveSheet()->setCellValue('E28', Milion_format($total_assets_var));
$objPHPExcel->getActiveSheet()->setCellValue('F28', Milion_format($total_assets_curr_mon_min1));

$objPHPExcel->getActiveSheet()->setCellValue('H28', Milion_format($total_assets_curr));


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
$objPHPExcel->getActiveSheet()->setCellValue('C34', Milion_format($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('D34', Milion_format($current_account3));
$objPHPExcel->getActiveSheet()->setCellValue('E34', Milion_format($current_account5));
$objPHPExcel->getActiveSheet()->setCellValue('F34', Milion_format($current_account4));
$objPHPExcel->getActiveSheet()->setCellValue('G34', Milion_format($current_account6));
$objPHPExcel->getActiveSheet()->setCellValue('H34', Milion_format($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('I34', Milion_format($budget_current_account));
$objPHPExcel->getActiveSheet()->setCellValue('J34', Milion_format($current_account7));

$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Saving Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C35', Milion_format($saving_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('D35', Milion_format($saving_deposits3));
$objPHPExcel->getActiveSheet()->setCellValue('E35', Milion_format($saving_deposits5));
$objPHPExcel->getActiveSheet()->setCellValue('F35', Milion_format($saving_deposits4));
$objPHPExcel->getActiveSheet()->setCellValue('G35', Milion_format($saving_deposits6));
$objPHPExcel->getActiveSheet()->setCellValue('H35', Milion_format($saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('I35', Milion_format($budget_saving_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('J35', Milion_format($saving_deposits7));

$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Time Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C36', Milion_format($time_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('D36', Milion_format($time_deposits3));
$objPHPExcel->getActiveSheet()->setCellValue('E36', Milion_format($time_deposits5));
$objPHPExcel->getActiveSheet()->setCellValue('F36', Milion_format($time_deposits4));
$objPHPExcel->getActiveSheet()->setCellValue('G36', Milion_format($time_deposits6));
$objPHPExcel->getActiveSheet()->setCellValue('H36', Milion_format($time_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('I36', Milion_format($budget_time_deposits));
$objPHPExcel->getActiveSheet()->setCellValue('J36', Milion_format($time_deposits7));

$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Total deposits	');
$objPHPExcel->getActiveSheet()->setCellValue('C37', Milion_format($total_deposit_curr));
$objPHPExcel->getActiveSheet()->setCellValue('D37', Milion_format($total_deposit_curr_min1));
$objPHPExcel->getActiveSheet()->setCellValue('E37', Milion_format($total_deposit_var));
$objPHPExcel->getActiveSheet()->setCellValue('F37', Milion_format($total_deposit_curr_mon_min1));

$objPHPExcel->getActiveSheet()->setCellValue('H37', Milion_format($total_deposit_curr));

$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Interbank');
$objPHPExcel->getActiveSheet()->setCellValue('C38', Milion_format($interbank));
$objPHPExcel->getActiveSheet()->setCellValue('D38', Milion_format($interbank3));
$objPHPExcel->getActiveSheet()->setCellValue('E38', Milion_format($interbank5));
$objPHPExcel->getActiveSheet()->setCellValue('F38', Milion_format($interbank4));
$objPHPExcel->getActiveSheet()->setCellValue('G38', Milion_format($interbank6));
$objPHPExcel->getActiveSheet()->setCellValue('H38', Milion_format($interbank));
$objPHPExcel->getActiveSheet()->setCellValue('I38', Milion_format($budget_interbank));
$objPHPExcel->getActiveSheet()->setCellValue('J38', Milion_format($interbank7));

$objPHPExcel->getActiveSheet()->setCellValue('B39', '-	Call Money');
$objPHPExcel->getActiveSheet()->setCellValue('C39', Milion_format($call_money));
$objPHPExcel->getActiveSheet()->setCellValue('D39', Milion_format($call_money3));
$objPHPExcel->getActiveSheet()->setCellValue('E39', Milion_format($call_money5));
$objPHPExcel->getActiveSheet()->setCellValue('F39', Milion_format($call_money4));
$objPHPExcel->getActiveSheet()->setCellValue('G39', Milion_format($call_money6));
$objPHPExcel->getActiveSheet()->setCellValue('H39', Milion_format($call_money));
$objPHPExcel->getActiveSheet()->setCellValue('I39', Milion_format($budget_call_money));
$objPHPExcel->getActiveSheet()->setCellValue('J39', Milion_format($call_money7));

$objPHPExcel->getActiveSheet()->setCellValue('B40', '-	Bank Deposits');
$objPHPExcel->getActiveSheet()->setCellValue('C40', Milion_format($bank_deposit));
$objPHPExcel->getActiveSheet()->setCellValue('D40', Milion_format($bank_deposit3));
$objPHPExcel->getActiveSheet()->setCellValue('E40', Milion_format($bank_deposit5));
$objPHPExcel->getActiveSheet()->setCellValue('F40', Milion_format($bank_deposit4));
$objPHPExcel->getActiveSheet()->setCellValue('G40', Milion_format($bank_deposit6));
$objPHPExcel->getActiveSheet()->setCellValue('H40', Milion_format($bank_deposit));
$objPHPExcel->getActiveSheet()->setCellValue('I40', Milion_format($budget_bank_deposit));
$objPHPExcel->getActiveSheet()->setCellValue('J40', Milion_format($bank_deposit7));

$objPHPExcel->getActiveSheet()->setCellValue('B41', '-	Current account	');
$objPHPExcel->getActiveSheet()->setCellValue('C41', Milion_format($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('D41', Milion_format($current_account3));
$objPHPExcel->getActiveSheet()->setCellValue('E41', Milion_format($current_account5));
$objPHPExcel->getActiveSheet()->setCellValue('F41', Milion_format($current_account4));
$objPHPExcel->getActiveSheet()->setCellValue('G41', Milion_format($current_account6));
$objPHPExcel->getActiveSheet()->setCellValue('H41', Milion_format($current_account));
$objPHPExcel->getActiveSheet()->setCellValue('I41', Milion_format($budget_current_account));
$objPHPExcel->getActiveSheet()->setCellValue('J41', Milion_format($current_account7));

$objPHPExcel->getActiveSheet()->setCellValue('B42', '-	Saving Account	');
$objPHPExcel->getActiveSheet()->setCellValue('C42', Milion_format($saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('D42', Milion_format($saving_account3));
$objPHPExcel->getActiveSheet()->setCellValue('E42', Milion_format($saving_account5));
$objPHPExcel->getActiveSheet()->setCellValue('F42', Milion_format($saving_account4));
$objPHPExcel->getActiveSheet()->setCellValue('G42', Milion_format($saving_account6));
$objPHPExcel->getActiveSheet()->setCellValue('H42', Milion_format($saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('I42', Milion_format($budget_saving_account));
$objPHPExcel->getActiveSheet()->setCellValue('J42', Milion_format($saving_account7));

$objPHPExcel->getActiveSheet()->setCellValue('B43', 'Derivative payable	');
$objPHPExcel->getActiveSheet()->setCellValue('C43', Milion_format($derivative_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D43', Milion_format($derivative_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E43', Milion_format($derivative_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F43', Milion_format($derivative_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G43', Milion_format($derivative_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H43', Milion_format($derivative_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I43', Milion_format($budget_derivative_payable));
$objPHPExcel->getActiveSheet()->setCellValue('J43', Milion_format($derivative_payable7));

$objPHPExcel->getActiveSheet()->setCellValue('B44', 'Acceptance payable	');
$objPHPExcel->getActiveSheet()->setCellValue('C44', Milion_format($acceptance_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D44', Milion_format($acceptance_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E44', Milion_format($acceptance_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F44', Milion_format($acceptance_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G44', Milion_format($acceptance_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H44', Milion_format($acceptance_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I44', Milion_format($budget_acceptance_payable));
$objPHPExcel->getActiveSheet()->setCellValue('J44', Milion_format($acceptance_payable7));

$objPHPExcel->getActiveSheet()->setCellValue('B45', 'KLBI Payable');
$objPHPExcel->getActiveSheet()->setCellValue('C45', Milion_format($klbi_payable));
$objPHPExcel->getActiveSheet()->setCellValue('D45', Milion_format($klbi_payable3));
$objPHPExcel->getActiveSheet()->setCellValue('E45', Milion_format($klbi_payable5));
$objPHPExcel->getActiveSheet()->setCellValue('F45', Milion_format($klbi_payable4));
$objPHPExcel->getActiveSheet()->setCellValue('G45', Milion_format($klbi_payable6));
$objPHPExcel->getActiveSheet()->setCellValue('H45', Milion_format($klbi_payable));
$objPHPExcel->getActiveSheet()->setCellValue('I45', Milion_format($budget_klbi_payable));
$objPHPExcel->getActiveSheet()->setCellValue('J45', Milion_format($klbi_payable7));

$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Mandatory Convertible Bonds');
$objPHPExcel->getActiveSheet()->setCellValue('C46', Milion_format($mandatory_convertible_bonds));
$objPHPExcel->getActiveSheet()->setCellValue('D46', Milion_format($mandatory_convertible_bonds3));
$objPHPExcel->getActiveSheet()->setCellValue('E46', Milion_format($mandatory_convertible_bonds5));
$objPHPExcel->getActiveSheet()->setCellValue('F46', Milion_format($mandatory_convertible_bonds4));
$objPHPExcel->getActiveSheet()->setCellValue('G46', Milion_format($mandatory_convertible_bonds6));
$objPHPExcel->getActiveSheet()->setCellValue('H46', Milion_format($mandatory_convertible_bonds));
$objPHPExcel->getActiveSheet()->setCellValue('I46', Milion_format($budget_mandatory_convertible_bonds));
$objPHPExcel->getActiveSheet()->setCellValue('J46', Milion_format($mandatory_convertible_bonds7));

$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Securities sold with agreement to repurchase');
$objPHPExcel->getActiveSheet()->setCellValue('C47', Milion_format($scurities_sold_watr));
$objPHPExcel->getActiveSheet()->setCellValue('D47', Milion_format($scurities_sold_watr3));
$objPHPExcel->getActiveSheet()->setCellValue('E47', Milion_format($scurities_sold_watr5));
$objPHPExcel->getActiveSheet()->setCellValue('F47', Milion_format($scurities_sold_watr4));
$objPHPExcel->getActiveSheet()->setCellValue('G47', Milion_format($scurities_sold_watr6));
$objPHPExcel->getActiveSheet()->setCellValue('H47', Milion_format($scurities_sold_watr));
$objPHPExcel->getActiveSheet()->setCellValue('I47', Milion_format($budget_scurities_sold_watr));
$objPHPExcel->getActiveSheet()->setCellValue('J47', Milion_format($scurities_sold_watr7));

$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('C48', Milion_format($others_liabilities));
$objPHPExcel->getActiveSheet()->setCellValue('D48', Milion_format($others_liabilities3));
$objPHPExcel->getActiveSheet()->setCellValue('E48', Milion_format($others_liabilities5));
$objPHPExcel->getActiveSheet()->setCellValue('F48', Milion_format($others_liabilities4));
$objPHPExcel->getActiveSheet()->setCellValue('G48', Milion_format($others_liabilities6));
$objPHPExcel->getActiveSheet()->setCellValue('H48', Milion_format($others_liabilities));
$objPHPExcel->getActiveSheet()->setCellValue('I48', Milion_format($budget_others_liabilities));
$objPHPExcel->getActiveSheet()->setCellValue('J48', Milion_format($others_liabilities7));

$objPHPExcel->getActiveSheet()->setCellValue('B49', 'Total Other Liabilities');
$objPHPExcel->getActiveSheet()->setCellValue('C49', Milion_format($total_ol_curr));
$objPHPExcel->getActiveSheet()->setCellValue('D49', Milion_format($total_ol_curr_min1));
$objPHPExcel->getActiveSheet()->setCellValue('E49', Milion_format($total_ol_var));
$objPHPExcel->getActiveSheet()->setCellValue('F49', Milion_format($total_ol_curr_mon_min1));

$objPHPExcel->getActiveSheet()->setCellValue('H49', Milion_format($total_ol_curr));

$objPHPExcel->getActiveSheet()->setCellValue('B50', 'Paid in capital');
$objPHPExcel->getActiveSheet()->setCellValue('C50', Milion_format($paid_in_capital));
$objPHPExcel->getActiveSheet()->setCellValue('D50', Milion_format($paid_in_capital3));
$objPHPExcel->getActiveSheet()->setCellValue('E50', Milion_format($paid_in_capital5));
$objPHPExcel->getActiveSheet()->setCellValue('F50', Milion_format($paid_in_capital4));
$objPHPExcel->getActiveSheet()->setCellValue('G50', Milion_format($paid_in_capital6));
$objPHPExcel->getActiveSheet()->setCellValue('H50', Milion_format($paid_in_capital));
$objPHPExcel->getActiveSheet()->setCellValue('I50', Milion_format($budget_paid_in_capital));
$objPHPExcel->getActiveSheet()->setCellValue('J50', Milion_format($paid_in_capital7));

$objPHPExcel->getActiveSheet()->setCellValue('B51', 'Agio ( disagio)');
$objPHPExcel->getActiveSheet()->setCellValue('C51', Milion_format($agio_disagio));
$objPHPExcel->getActiveSheet()->setCellValue('D51', Milion_format($agio_disagio3));
$objPHPExcel->getActiveSheet()->setCellValue('E51', Milion_format($agio_disagio5));
$objPHPExcel->getActiveSheet()->setCellValue('F51', Milion_format($agio_disagio4));
$objPHPExcel->getActiveSheet()->setCellValue('G51', Milion_format($agio_disagio6));
$objPHPExcel->getActiveSheet()->setCellValue('H51', Milion_format($agio_disagio));
$objPHPExcel->getActiveSheet()->setCellValue('I51', Milion_format($budget_agio_disagio));
$objPHPExcel->getActiveSheet()->setCellValue('J51', Milion_format($agio_disagio7));

$objPHPExcel->getActiveSheet()->setCellValue('B52', 'General reserve');
$objPHPExcel->getActiveSheet()->setCellValue('C52', Milion_format($general_reserve));
$objPHPExcel->getActiveSheet()->setCellValue('D52', Milion_format($general_reserve3));
$objPHPExcel->getActiveSheet()->setCellValue('E52', Milion_format($general_reserve5));
$objPHPExcel->getActiveSheet()->setCellValue('F52', Milion_format($general_reserve4));
$objPHPExcel->getActiveSheet()->setCellValue('G52', Milion_format($general_reserve6));
$objPHPExcel->getActiveSheet()->setCellValue('H52', Milion_format($general_reserve));
$objPHPExcel->getActiveSheet()->setCellValue('I52', Milion_format($budget_general_reserve));
$objPHPExcel->getActiveSheet()->setCellValue('J52', Milion_format($general_reserve7));

$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Available for sale securities - net');
$objPHPExcel->getActiveSheet()->setCellValue('C53', Milion_format($available_fss_net));
$objPHPExcel->getActiveSheet()->setCellValue('D53', Milion_format($available_fss_net3));
$objPHPExcel->getActiveSheet()->setCellValue('E53', Milion_format($available_fss_net5));
$objPHPExcel->getActiveSheet()->setCellValue('F53', Milion_format($available_fss_net4));
$objPHPExcel->getActiveSheet()->setCellValue('G53', Milion_format($available_fss_net6));
$objPHPExcel->getActiveSheet()->setCellValue('H53', Milion_format($available_fss_net));
$objPHPExcel->getActiveSheet()->setCellValue('I53', Milion_format($budget_available_fss_net));
$objPHPExcel->getActiveSheet()->setCellValue('J53', Milion_format($available_fss_net7));

$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Retained earnings');
$objPHPExcel->getActiveSheet()->setCellValue('C54', Milion_format($retained_earning));
$objPHPExcel->getActiveSheet()->setCellValue('D54', Milion_format($retained_earning3));
$objPHPExcel->getActiveSheet()->setCellValue('E54', Milion_format($retained_earning5));
$objPHPExcel->getActiveSheet()->setCellValue('F54', Milion_format($retained_earning4));
$objPHPExcel->getActiveSheet()->setCellValue('G54', Milion_format($retained_earning6));
$objPHPExcel->getActiveSheet()->setCellValue('H54', Milion_format($retained_earning));
$objPHPExcel->getActiveSheet()->setCellValue('I54', Milion_format($budget_retained_earning));
$objPHPExcel->getActiveSheet()->setCellValue('J54', Milion_format($retained_earning7));

$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Profit/loss current year');
$objPHPExcel->getActiveSheet()->setCellValue('C55', Milion_format($profit_los));
$objPHPExcel->getActiveSheet()->setCellValue('D55', Milion_format($profit_los3));
$objPHPExcel->getActiveSheet()->setCellValue('E55', Milion_format($profit_los5));
$objPHPExcel->getActiveSheet()->setCellValue('F55', Milion_format($profit_los4));
$objPHPExcel->getActiveSheet()->setCellValue('G55', Milion_format($profit_los6));
$objPHPExcel->getActiveSheet()->setCellValue('H55', Milion_format($profit_los));
$objPHPExcel->getActiveSheet()->setCellValue('I55', Milion_format($budget_profit_los));
$objPHPExcel->getActiveSheet()->setCellValue('J55', Milion_format($profit_los7));

$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Total Equity');
$objPHPExcel->getActiveSheet()->setCellValue('C56', Milion_format($total_equity_curr));
$objPHPExcel->getActiveSheet()->setCellValue('D56', Milion_format($total_equity_curr_min1));
$objPHPExcel->getActiveSheet()->setCellValue('E56', Milion_format($total_equity_var));
$objPHPExcel->getActiveSheet()->setCellValue('F56', Milion_format($total_equity_curr_mon_min1));

$objPHPExcel->getActiveSheet()->setCellValue('H56', Milion_format($total_equity_curr));

$objPHPExcel->getActiveSheet()->setCellValue('B57', 'TOTAL LIABILITIES & EQUITY');
$objPHPExcel->getActiveSheet()->setCellValue('C57', Milion_format($total_deposit_curr+$total_ol_curr+$total_equity_curr));
$objPHPExcel->getActiveSheet()->setCellValue('D57', Milion_format($total_deposit_curr_min1+$total_ol_curr_min1+$total_equity_curr_min1));
$objPHPExcel->getActiveSheet()->setCellValue('E57', Milion_format($total_deposit_var+$total_ol_var+$total_equity_var));
$objPHPExcel->getActiveSheet()->setCellValue('F57', Milion_format($total_deposit_curr_mon_min1+$total_ol_curr_mon_min1+$total_equity_curr_mon_min1));

$objPHPExcel->getActiveSheet()->setCellValue('H57', Milion_format($total_deposit_curr+$total_ol_curr+$total_equity_curr));



	
$objPHPExcel->getActiveSheet()->setCellValue('B60', $label_tgl_min1);
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'New NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Penambah_OS_NPL');
$objPHPExcel->getActiveSheet()->setCellValue('B63', 'Total New NPL');
$objPHPExcel->getActiveSheet()->setCellValue('B64', '');	
$objPHPExcel->getActiveSheet()->setCellValue('B65', 'NPL to PL (Reklass) ');
$objPHPExcel->getActiveSheet()->setCellValue('B66', 'NPL Paid Off');
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'Reverse Saldo NPL ');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'NPL Payment');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'NPL Exchange Rate');
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'NPL Credit Card');
$objPHPExcel->getActiveSheet()->setCellValue('B71', $label_tgl);

// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle('BS_Flash Report');

//=======BORDER
$styleArray = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$objPHPExcel->getActiveSheet()->getStyle('B5:J28')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B31:J57')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B60:C71')->applyFromArray($styleArray);
//=======END BORDER

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');