<?php
//require_once 'config/config.php';

require_once '../../config/config.php';


//$tanggal=$_POST['tanggal']; 

//error_reporting(0);
$tanggal="2015-06-30";
$day=date('d',strtotime($tanggal));
$day_min1=date('j',strtotime($tanggal))-1;

//if (strlen($day_min1)==1){
//$day_min1="0".$day_min1;
//}

$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$prev_date=date('t-M-y', strtotime(date('Y-m',strtotime($tanggal))." -1 month")); // tanggal terakhir pada bulan sebelumnya 

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih

$label_tgl_min1=date('d-M-y', strtotime(date('Y-m-d',strtotime($tanggal))." -1 day")); // tanggal terpilih minus (-) 1



$curr_tgl=date('Y-m-d',strtotime($tanggal));
$curr_tgl_min1=date('Y-m-d',strtotime($label_tgl_min1));
$curr_mon_min1=date('Y-m-d',strtotime($prev_date
	));

echo "$curr_tgl";
echo "<br>";
echo "$curr_tgl_min1";
echo "<br>";
echo "$curr_mon_min1";
echo "<br>";

//=======================================

$var_curr_tgl=" and c.DataDate='".$curr_tgl."' ";//var tgl terpilih
$var_curr_tgl_min1=" and c.DataDate='".$curr_tgl_min1."' ";//var tgl terpilih minus 1
$var_curr_mon_min1=" and c.DataDate='".$curr_mon_min1."' ";//var tgl terakhir bulan sebelumnya
//and a.FLASH_Level_3_Description='Cash'

$query_currentDate="select sum (c.Nominal) as jml_nominal from Referensi_Flash_Report a, GL_02_Baru b, DM_Journal c where a.FLASH_Level_3=b.FLASH_LEVEL_3 and b.GLNO=c.KodeGL  ";



//========================================
$query="select distinct FLASH_Level_3_Description from Referensi_Flash_Report  order by FLASH_Level_3_Description asc";
$result=odbc_exec($connection, $query);
//$jsonData = array();
$i=1;
while ($row = odbc_fetch_array($result)) {
    //$jsonData[] = $array;

switch ($row['FLASH_Level_3_Description']) {
    case "Cash":
    //current_date
        $var_flash=" and a.FLASH_Level_3_Description='Cash' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $cash=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $cash3=$row3['jml_nominal'];

        $cash5=$cash-$cash3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $cash4=$row4['jml_nominal'];

        break;
    case "Current account - Bank Indonesia":
        $var_flash=" and a.FLASH_Level_3_Description='Current account - Bank Indonesia' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $current_account_bi=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $current_account_bi3=$row3['jml_nominal'];

        $current_account_bi5=$current_account_bi-$current_account_bi3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $current_account_bi4=$row4['jml_nominal'];
        break;
    case "Certificate of bank Indonesia (SBI & BI call money)":
        $var_flash=" and a.FLASH_Level_3_Description='Certificate of bank Indonesia (SBI & BI call money)' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $certificate_bi=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $certificate_bi3=$row3['jml_nominal'];

        $certificate_bi5=$certificate_bi-$certificate_bi3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $certificate_bi4=$row4['jml_nominal'];
        break;
    case "Interbank Placement":
        $var_flash=" and a.FLASH_Level_3_Description='Interbank Placement' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $interbank_placement=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $interbank_placement3=$row3['jml_nominal'];

        $interbank_placement5=$interbank_placement-$interbank_placement3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $interbank_placement4=$row4['jml_nominal'];
        break;
    case "Securities":
        $var_flash=" and a.FLASH_Level_3_Description='Securities' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $scurities=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $scurities3=$row3['jml_nominal'];

        $scurities5=$scurities-$scurities3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $scurities4=$row4['jml_nominal'];
        break;
    case "Allowance For Securities":
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Securities' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $allowence_fs=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $allowence_fs3=$row3['jml_nominal'];

        $allowence_fs5=$allowence_fs-$allowence_fs3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $allowence_fs4=$row4['jml_nominal'];
        break;
    case "Loans":
        $var_flash=" and a.FLASH_Level_3_Description='Loans' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $loans=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $loans3=$row3['jml_nominal'];

        $loans5=$loans-$loans3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $loans4=$row4['jml_nominal'];
        break;
    case "Performing Loan":
        $var_flash=" and a.FLASH_Level_3_Description='Performing Loan' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $performing_loan=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $performing_loan3=$row3['jml_nominal'];

        $performing_loan5=$performing_loan-$performing_loan3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $performing_loan4=$row4['jml_nominal'];
        break;
    case "Non Performing Loan*)":
        $var_flash=" and a.FLASH_Level_3_Description='Non Performing Loan*)' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $non_performing_loan=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $non_performing_loan3=$row3['jml_nominal'];

        $non_performing_loan5=$non_performing_loan-$non_performing_loan3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $non_performing_loan4=$row4['jml_nominal'];
        break;
    case "Allowance For Loan":
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Loan' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $allowence_for_loan=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $allowence_for_loan3=$row3['jml_nominal']; 

        $allowence_for_loan5=$allowence_for_loan-$allowence_for_loan3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $allowence_for_loan4=$row4['jml_nominal'];
        break;
    case "Acceptance receivables":
        $var_flash=" and a.FLASH_Level_3_Description='Acceptance receivables' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $acceptance_recevables=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $acceptance_recevables3=$row3['jml_nominal'];

        $acceptance_recevables5=$acceptance_recevables-$acceptance_recevables3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $acceptance_recevables4=$row4['jml_nominal'];
        break; //==================================================================================================
    case "Derivative receivables":
        $var_flash=" and a.FLASH_Level_3_Description='Derivative receivables' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $deferred_receivables=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $deferred_receivables3=$row3['jml_nominal'];

        $deferred_receivables5=$deferred_receivables-$deferred_receivables3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $deferred_receivables4=$row4['jml_nominal'];
        break;
    case "Fixed assets (Property, Plant Equipment)":
        $var_flash=" and a.FLASH_Level_3_Description='Fixed assets (Property, Plant Equipment)' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $fixed_assets=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $fixed_assets3=$row3['jml_nominal']; 

        $fixed_assets5=$fixed_assets-$fixed_assets3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $fixed_assets4=$row4['jml_nominal'];
        break;
    case "Deferred taxes":
        $var_flash=" and a.FLASH_Level_3_Description='Deferred taxes' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $deferred_taxes=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $deferred_taxes3=$row3['jml_nominal']; 

        $deferred_taxes5=$deferred_taxes-$deferred_taxes3;


        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $deferred_taxes4=$row4['jml_nominal'];
        break;
    case "Other assets":
        $var_flash=" and a.FLASH_Level_3_Description='Other assets' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $others_assets=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $others_assets3=$row3['jml_nominal'];

        $others_assets5=$others_assets-$others_assets3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $others_assets4=$row4['jml_nominal']; 
        break;
    case "Foreclosed properties":
        $var_flash=" and a.FLASH_Level_3_Description='Foreclosed properties' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $foreclose_properties=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $foreclose_properties3=$row3['jml_nominal']; 

        $foreclose_properties5=$foreclose_properties-$foreclose_properties3;


        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $foreclose_properties4=$row4['jml_nominal'];
        break;
    case "Allowance For Foreclosed properties":
        $var_flash=" and a.FLASH_Level_3_Description='Allowance For Foreclosed properties' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $allowence_for_fp=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $allowence_for_fp3=$row3['jml_nominal']; 

        $allowence_for_fp5=$allowence_for_fp-$allowence_for_fp3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $allowence_for_fp4=$row4['jml_nominal'];
        break;
    case "Account receivable":
        $var_flash=" and a.FLASH_Level_3_Description='Account receivable' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $account_receivable=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $account_receivable3=$row3['jml_nominal']; 

        $account_receivable5=$account_receivable-$account_receivable3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $account_receivable4=$row4['jml_nominal'];
        break;
    case "Others - Other assets":
        $var_flash=" and a.FLASH_Level_3_Description='Others - Other assets' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $others_assets=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $others_assets3=$row3['jml_nominal']; 


        $others_assets5=$others_assets-$others_assets3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $others_assets4=$row4['jml_nominal'];
        break;
    case "Allowances For Suspence Account":
        $var_flash=" and a.FLASH_Level_3_Description='Allowances For Suspence Account' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $allowence_fsa=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $allowence_fsa3=$row3['jml_nominal']; 

        $allowence_fsa5=$allowence_fsa-$allowence_fsa3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $allowence_fsa4=$row4['jml_nominal'];
        break;
    case "Current Account":
        $var_flash=" and a.FLASH_Level_3_Description='Current Account' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $current_account=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $current_account3=$row3['jml_nominal'];

        $current_account5=$current_account-$current_account3;


        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $current_account4=$row4['jml_nominal'];
        break;
    case "Saving Deposits":
        $var_flash=" and a.FLASH_Level_3_Description='Saving Deposits' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $saving_deposits=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $saving_deposits3=$row3['jml_nominal'];

        $saving_deposits5=$saving_deposits-$saving_deposits3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $saving_deposits4=$row4['jml_nominal'];
        break;
    case "Time deposits":
        $var_flash=" and a.FLASH_Level_3_Description='Time deposits' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $time_deposits=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $time_deposits3=$row3['jml_nominal']; 

        $time_deposits5=$time_deposits-$time_deposits3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $time_deposits4=$row4['jml_nominal'];
        break;
    case "Interbank":
        $var_flash=" and a.FLASH_Level_3_Description='Interbank' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $interbank=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $interbank3=$row3['jml_nominal']; 

        $interbank5=$interbank-$interbank3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $interbank4=$row4['jml_nominal'];
        break;
    case "Call Money":
        $var_flash=" and a.FLASH_Level_3_Description='Call Money' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $call_money=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $call_money3=$row3['jml_nominal'];

        $call_money5=$call_money-$call_money3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $call_money4=$row4['jml_nominal'];
        break;
    case "Bank deposits":
        $var_flash=" and a.FLASH_Level_3_Description='Bank deposits' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $bank_deposit=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $bank_deposit3=$row3['jml_nominal']; 

        $bank_deposit5=$bank_deposit-$bank_deposit3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $bank_deposit4=$row4['jml_nominal'];
        break;
    case "Interbank Current Account":
        $var_flash=" and a.FLASH_Level_3_Description='Current Account' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $current_account_interbank=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $current_account_interbank3=$row3['jml_nominal']; 

        $current_account_interbank5=$current_account_interbank-$current_account_interbank3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $current_account_interbank4=$row4['jml_nominal'];
        break;
	case "Saving accounts":
        $var_flash=" and a.FLASH_Level_3_Description='Saving accounts' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $saving_account=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $saving_account3=$row3['jml_nominal']; 

        $saving_account5=$saving_account-$saving_account3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $saving_account4=$row4['jml_nominal'];
        break;
    case "Derivative payable":
        $var_flash=" and a.FLASH_Level_3_Description='Derivative payable' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $derivative_payable=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $derivative_payable3=$row3['jml_nominal']; 

        $derivative_payable=$derivative_payable-$derivative_payable3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $derivative_payable4=$row4['jml_nominal'];
        break;
    case "Acceptance payable":
        $var_flash=" and a.FLASH_Level_3_Description='Acceptance payable' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $acceptance_payable=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $acceptance_payable3=$row3['jml_nominal']; 

        $acceptance_payable5=$acceptance_payable-$acceptance_payable3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $acceptance_payable4=$row4['jml_nominal'];
        break;
    case "KLBI Payable":
        $var_flash=" and a.FLASH_Level_3_Description='KLBI Payable' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $klbi_payable=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $klbi_payable3=$row3['jml_nominal']; 

        $klbi_payable5=$klbi_payable-$klbi_payable3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $klbi_payable4=$row4['jml_nominal'];
        break;
    case "Mandatory Convertible Bonds":
        $var_flash=" and a.FLASH_Level_3_Description='Mandatory Convertible Bonds' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $mandatory_convertible_bonds=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $mandatory_convertible_bonds3=$row3['jml_nominal'];

        $mandatory_convertible_bonds5=$mandatory_convertible_bonds-$mandatory_convertible_bonds3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $mandatory_convertible_bond4=$row4['jml_nominal'];
        break;
    case "Securities sold with agreement to repurchase":
        $var_flash=" and a.FLASH_Level_3_Description='Securities sold with agreement to repurchase' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $scurities_sold_watr=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $scurities_sold_watr3=$row3['jml_nominal'];

        $scurities_sold_watr5=$scurities_sold_watr-$scurities_sold_watr3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $scurities_sold_watr4=$row4['jml_nominal']; 
        break;
    case "Others Liabilities":
        $var_flash=" and a.FLASH_Level_3_Description='Others Liabilities' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $others_liabilities=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $others_liabilities3=$row3['jml_nominal']; 

        $others_liabilities5=$others_liabilities-$others_liabilities3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $others_liabilities4=$row4['jml_nominal']; 
        break;
    case "Paid in capital":
        $var_flash=" and a.FLASH_Level_3_Description='Paid in capital' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $paid_in_capital=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $paid_in_capital3=$row3['jml_nominal']; 


        $paid_in_capital5=$paid_in_capital-$paid_in_capital3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $paid_in_capital4=$row4['jml_nominal'];
        break;
    case "Agio ( disagio)":
        $var_flash=" and a.FLASH_Level_3_Description='Agio ( disagio)' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $agio_disagio=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $agio_disagio3=$row3['jml_nominal']; 


        $agio_disagio5=$agio_disagio-$agio_disagio3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $agio_disagio4=$row4['jml_nominal'];
        break;
    case "General reserve":
        $var_flash=" and a.FLASH_Level_3_Description='General reserve' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $general_reserve=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $general_reserve3=$row3['jml_nominal']; 


        $general_reserve5=$general_reserve-$general_reserve3;


        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $general_reserve4=$row4['jml_nominal'];
        break;
    case "Available for sale securities - net":
        $var_flash=" and a.FLASH_Level_3_Description='Available for sale securities - net' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $available_fss_net=$row2['jml_nominal'];

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $available_fss_net3=$row3['jml_nominal']; 


        $available_fss_net5=$available_fss_net-$available_fss_net3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $available_fss_net4=$row4['jml_nominal'];
        break;
    case "Retained earnings":
        $var_flash=" and a.FLASH_Level_3_Description='Retained earnings' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $retained_earning=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $retained_earning3=$row3['jml_nominal']; 

        $retained_earning5=$retained_earning-$retained_earning3;

        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $retained_earning4=$row4['jml_nominal'];
        break;
    case "Profit/loss current year":
        $var_flash=" and a.FLASH_Level_3_Description='Profit/loss current year' ";
        $result2=odbc_exec($connection, $query_currentDate.$var_curr_tgl.$var_flash);
        $row2=odbc_fetch_array($result2);
        $profit_los=$row2['jml_nominal']; 

        $result3=odbc_exec($connection, $query_currentDate.$var_curr_tgl_min1.$var_flash);
        $row3=odbc_fetch_array($result3);
        $profit_los3=$row3['jml_nominal']; 

        $profit_los5=$profit_los-$profit_los3;


        $result4=odbc_exec($connection, $query_currentDate.$var_curr_mon_min1.$var_flash);
        $row4=odbc_fetch_array($result4);
        $profit_los4=$row4['jml_nominal'];
        break;

}



$i++;
}


/*
echo "<pre>";
$data1=json_encode($jsonData);

//print_r($jsonData);
$obj = json_decode($data1,true);

print_r($obj);
echo "</pre>";
*/




echo "Cash :$cash";
echo "<br>";
echo "Current Account : $current_account_bi";
echo "<br>";
echo "certificate BI :$certificate_bi";
echo "<br>";
echo "Interbank : $interbank_placement";
echo "<br>";
echo "Securitie : $scurities";
echo "<br>";



?>


