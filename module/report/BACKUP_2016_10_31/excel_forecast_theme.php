<?php
//session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
//require_once '../../session_login.php';
//require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';
date_default_timezone_set("Asia/Bangkok");






$file_eksport=date('Y_m_d_H_i_s');

error_reporting(1);
logActivity("generate theme forecast",date('Y_m_d_H_i_s'));
$tanggal=$_POST['tanggal']; 

if (isset($tanggal) && $tanggal !="" )
{
$curr_tgl=date('Y/m/d',strtotime($tanggal));
$q_tgl=date('Y-m-d',strtotime($tanggal));
} else {
$curr_tgl=date('Y/m/d');
$tanggal=date('Y/m/d');
$q_tgl=date('Y-m-d');
}
$tgl_terakhir=date('Y/m/t', strtotime(date('Y-m',strtotime($tanggal))." 0 month")); // tanggal terakhir pada bulan sebelumnya 
$tgl_terakhir_min1=date('Y/m/t', strtotime(date('Y-m',strtotime($tanggal))." -1 month")); // tanggal 
$q_tgl_min1=date('Y-m-d',strtotime($tgl_terakhir_min1));

/*
SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM (SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) 
left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct 
left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 
WHERE a.DataDate='2016-05-31' and a. Flag_M='Y' and b.FLASH_LEVEL_3='FLASH201000004' GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang)AS tabel1
*/


/* cld version 09-08-2016
$query =" SELECT SUM(Nilai) /1000000 AS jml_nominal FROM( SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a "; 
$query.=" JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
$query.=" JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 WHERE  ";

$var_tgl= " a.DataDate='$q_tgl' ";

$var_add="  GROUP BY a.kodegl ,b.FLASH_LEVEL_3 ) AS tabel1 ";
*/
$query =" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM (SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) "; 
$query.=" left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct  ";
$query.=" left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 WHERE  ";

$var_tgl= " a.DataDate='$q_tgl' and a. Flag_M='Y'  ";

$var_add="  GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang) AS tabel1 ";



#--------------- FLASH101000007  Loan
        $var_flash=" and b.FLASH_Level_3='FLASH201000002' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $loan=$row2['jml_nominal'];
#--------------- FLASH201000003  Treasury bills
        $var_flash=" and b.FLASH_Level_3='FLASH201000003' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $treasury_bill=$row2['jml_nominal'];
#--------------- FLASH101000004  Interbank placements
        $var_flash=" and b.FLASH_Level_3='FLASH201000004' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $interbank_placements=$row2['jml_nominal'];


        //echo $query.$var_tgl.$var_flash.$var_add;
        //die();
#--------------- FLASH201000005  Placement with BI
        $var_flash=" and b.FLASH_Level_3='FLASH201000005' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $placement_wbi=$row2['jml_nominal'];
#--------------- FLASH101000019   Others
        $var_flash=" and b.FLASH_Level_3='FLASH201000006' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others1=$row2['jml_nominal'];
#--------------- FLASH102000001  Current accounts
        $var_flash=" and b.FLASH_Level_3='FLASH202000002' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $current_account=$row2['jml_nominal'];
#--------------- FLASH102000002  Saving accounts
        $var_flash=" and b.FLASH_Level_3='FLASH202000003' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $saving_account=$row2['jml_nominal'];
#--------------- FLASH102000003  Time deposits
        $var_flash=" and b.FLASH_Level_3='FLASH202000004' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $time_deposits=$row2['jml_nominal'];

#--------------- FLASH102000006  Bank deposits
        $var_flash=" and b.FLASH_Level_3='FLASH202000005' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $bank_deposits=$row2['jml_nominal'];

#--------------- FLASH202000007  Borrowings (MCB)
        $var_flash=" and b.FLASH_Level_3='FLASH202000007' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $borrowing_mcb=$row2['jml_nominal'];

#--------------- FLASH202000008  Guaranteed premium
        $var_flash=" and b.FLASH_Level_3='FLASH202000008' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $guaranted_premium=$row2['jml_nominal'];
#--------------- FLASH202000006  Others
        $var_flash=" and b.FLASH_Level_3='FLASH202000009' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others2=$row2['jml_nominal'];
#--------------- FLASH201000008  Forex gain/(loss) on transactions
        $var_flash=" and b.FLASH_Level_3='FLASH201000008' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $forex_gain=$row2['jml_nominal'];
#--------------- FLASH201000009  Gain/(Loss) on sale of securities/bonds
        $var_flash=" and b.FLASH_Level_3='FLASH201000009' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $gain_loss=$row2['jml_nominal'];
#--------------- FLASH201000010  Remittance fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000010' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $remittance_fee=$row2['jml_nominal'];
#--------------- FLASH201000011  Trade Finance fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000011' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $trade_finance_fee=$row2['jml_nominal'];
#--------------- FLASH201000012  Processing fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000012' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $processing_fee=$row2['jml_nominal'];
#--------------- FLASH201000013  Credit Card fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000013' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $credit_card_fee=$row2['jml_nominal'];
#--------------- FLASH201000014  Insurance Fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000014' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $insurance_fee=$row2['jml_nominal'];
#--------------- FLASH201000015  Service Charges
        $var_flash=" and b.FLASH_Level_3='FLASH201000015' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $service_charges=$row2['jml_nominal'];
#--------------- FLASH201000016  Other Commission & Fee 
        $var_flash=" and b.FLASH_Level_3='FLASH201000016' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_commission=$row2['jml_nominal'];
#--------------- FLASH201000017  Penalty
        $var_flash=" and b.FLASH_Level_3='FLASH201000017' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $penalty=$row2['jml_nominal'];

#--------------- FLASH201000018  Write Offs Recovered
        $var_flash=" and b.FLASH_Level_3='FLASH201000018' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_r=$row2['jml_nominal'];

#--------------- FLASH201000019  Other Income
        $var_flash=" and b.FLASH_Level_3='FLASH201000019' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_income=$row2['jml_nominal'];

#--------------- FLASH202000010  Staff Cost
        $var_flash=" and b.FLASH_Level_3='FLASH202000010' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $staff_cost=$row2['jml_nominal'];

#--------------- FLASH202000011  General & Administrative Expenses
        $var_flash=" and b.FLASH_Level_3='FLASH202000011' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $general_administrative=$row2['jml_nominal'];

#--------------- FLASH202000012  Depreciation
        $var_flash=" and b.FLASH_Level_3='FLASH202000012' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $depreciation=$row2['jml_nominal'];

#--------------- FLASH202000014  Other Operating Expense/income
        $var_flash=" and b.FLASH_Level_3='FLASH202000014' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_operating=$row2['jml_nominal'];

#--------------- FLASH202000015  General Provision
        $var_flash=" and b.FLASH_Level_3='FLASH202000015' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $general_provision=$row2['jml_nominal'];

#--------------- FLASH202000016  Specific Provision Charged
        $var_flash=" and b.FLASH_Level_3='FLASH202000016' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $specific_provision_charge=$row2['jml_nominal'];

#--------------- FLASH202000017  Specific Provision Recovery  
        $var_flash=" and b.FLASH_Level_3='FLASH202000017' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $specific_provision_recovery=$row2['jml_nominal'];

#--------------- FLASH202000018  Write Offs Charged
        $var_flash=" and b.FLASH_Level_3='FLASH202000018' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_c=$row2['jml_nominal'];

#--------------- FLASH202000019  Write Offs Recovered
        $var_flash=" and b.FLASH_Level_3='FLASH202000019' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_r2=$row2['jml_nominal'];

#--------------- FLASH202000020  Foreclose Properties Provision
        $var_flash=" and b.FLASH_Level_3='FLASH202000020' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $foreclose_pp=$row2['jml_nominal'];

#--------------- FLASH202000021  Others
        $var_flash=" and b.FLASH_Level_3='FLASH202000021' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others3=$row2['jml_nominal'];

#--------------- FLASH202000023  Corporate Tax
        $var_flash=" and b.FLASH_Level_3='FLASH202000023' ";
        $result2=odbc_exec($connection2, $query.$var_tgl.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $corporate_tax=$row2['jml_nominal'];


// BUlan sebelumnya=====================
        /*
$query =" SELECT SUM(Nilai)/1000000 AS jml_nominal FROM( SELECT a.kodegl,SUM(a.nominal) AS Nilai FROM DM_Journal a "; 
$query.=" JOIN Referensi_GL_02_New b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
$query.=" JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 WHERE  ";

$var_tgl_1= " a.DataDate='$q_tgl_min1' ";

$var_add="  GROUP BY a.kodegl ,b.FLASH_LEVEL_3 ) AS tabel1 ";
*/
$query =" SELECT SUM(Nilai)*(-1)/1000000 AS jml_nominal FROM (SELECT a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) "; 
$query.=" left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
$query.=" left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 WHERE  ";

$var_tgl_1= " a.DataDate='$q_tgl_min1' and a. Flag_M='Y' ";

$var_add="  GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang) AS tabel1 ";





#--------------- FLASH101000007  Loan
        $var_flash=" and b.FLASH_Level_3='FLASH201000002' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $loan_1=$row2['jml_nominal'];

       // echo $query.$var_tgl_1.$var_flash.$var_add;
        //die();
#--------------- FLASH201000003  Treasury bills
        $var_flash=" and b.FLASH_Level_3='FLASH201000003' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $treasury_bill_1=$row2['jml_nominal'];
#--------------- FLASH101000004  Interbank placements
        $var_flash=" and b.FLASH_Level_3='FLASH201000004' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $interbank_placements_1=$row2['jml_nominal'];
#--------------- FLASH201000005  Placement with BI
        $var_flash=" and b.FLASH_Level_3='FLASH201000005' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $placement_wbi_1=$row2['jml_nominal'];
#--------------- FLASH101000019   Others
        $var_flash=" and b.FLASH_Level_3='FLASH201000006' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others1_1=$row2['jml_nominal'];
#--------------- FLASH102000001  Current accounts
        $var_flash=" and b.FLASH_Level_3='FLASH202000002' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $current_account_1=$row2['jml_nominal'];
#--------------- FLASH102000002  Saving accounts
        $var_flash=" and b.FLASH_Level_3='FLASH202000003' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $saving_account_1=$row2['jml_nominal'];
#--------------- FLASH102000003  Time deposits
        $var_flash=" and b.FLASH_Level_3='FLASH202000004' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $time_deposits_1=$row2['jml_nominal'];

#--------------- FLASH102000006  Bank deposits
        $var_flash=" and b.FLASH_Level_3='FLASH202000005' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $bank_deposits_1=$row2['jml_nominal'];

#--------------- FLASH202000007  Borrowings (MCB)
        $var_flash=" and b.FLASH_Level_3='FLASH202000007' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $borrowing_mcb_1=$row2['jml_nominal'];

#--------------- FLASH202000008  Guaranteed premium
        $var_flash=" and b.FLASH_Level_3='FLASH202000008' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $guaranted_premium_1=$row2['jml_nominal'];
#--------------- FLASH202000006  Others
        $var_flash=" and b.FLASH_Level_3='FLASH202000009' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others2_1=$row2['jml_nominal'];
#--------------- FLASH201000008  Forex gain/(loss) on transactions
        $var_flash=" and b.FLASH_Level_3='FLASH201000008' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $forex_gain_1=$row2['jml_nominal'];
#--------------- FLASH201000009  Gain/(Loss) on sale of securities/bonds
        $var_flash=" and b.FLASH_Level_3='FLASH201000009' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $gain_loss_1=$row2['jml_nominal'];
#--------------- FLASH201000010  Remittance fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000010' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $remittance_fee_1=$row2['jml_nominal'];
#--------------- FLASH201000011  Trade Finance fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000011' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $trade_finance_fee_1=$row2['jml_nominal'];
#--------------- FLASH201000012  Processing fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000012' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $processing_fee_1=$row2['jml_nominal'];
#--------------- FLASH201000013  Credit Card fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000013' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $credit_card_fee_1=$row2['jml_nominal'];
#--------------- FLASH201000014  Insurance Fee
        $var_flash=" and b.FLASH_Level_3='FLASH201000014' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $insurance_fee_1=$row2['jml_nominal'];
#--------------- FLASH201000015  Service Charges
        $var_flash=" and b.FLASH_Level_3='FLASH201000015' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $service_charges_1=$row2['jml_nominal'];
#--------------- FLASH201000016  Other Commission & Fee 
        $var_flash=" and b.FLASH_Level_3='FLASH201000016' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_commission_1=$row2['jml_nominal'];
#--------------- FLASH201000017  Penalty
        $var_flash=" and b.FLASH_Level_3='FLASH201000017' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $penalty_1=$row2['jml_nominal'];

#--------------- FLASH201000018  Write Offs Recovered
        $var_flash=" and b.FLASH_Level_3='FLASH201000018' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_r_1=$row2['jml_nominal'];

#--------------- FLASH201000019  Other Income
        $var_flash=" and b.FLASH_Level_3='FLASH201000019' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_income_1=$row2['jml_nominal'];

#--------------- FLASH202000010  Staff Cost
        $var_flash=" and b.FLASH_Level_3='FLASH202000010' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $staff_cost_1=$row2['jml_nominal'];

#--------------- FLASH202000011  General & Administrative Expenses
        $var_flash=" and b.FLASH_Level_3='FLASH202000011' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $general_administrative_1=$row2['jml_nominal'];

#--------------- FLASH202000012  Depreciation
        $var_flash=" and b.FLASH_Level_3='FLASH202000012' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $depreciation_1=$row2['jml_nominal'];

#--------------- FLASH202000014  Other Operating Expense/income
        $var_flash=" and b.FLASH_Level_3='FLASH202000014' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $other_operating_1=$row2['jml_nominal'];

#--------------- FLASH202000015  General Provision
        $var_flash=" and b.FLASH_Level_3='FLASH202000015' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $general_provision_1=$row2['jml_nominal'];

#--------------- FLASH202000016  Specific Provision Charged
        $var_flash=" and b.FLASH_Level_3='FLASH202000016' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $specific_provision_charge_1=$row2['jml_nominal'];

#--------------- FLASH202000017  Specific Provision Recovery  
        $var_flash=" and b.FLASH_Level_3='FLASH202000017' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $specific_provision_recovery_1=$row2['jml_nominal'];

#--------------- FLASH202000018  Write Offs Charged
        $var_flash=" and b.FLASH_Level_3='FLASH202000018' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_c_1=$row2['jml_nominal'];

#--------------- FLASH202000019  Write Offs Recovered
        $var_flash=" and b.FLASH_Level_3='FLASH202000019' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $write_off_r2_1=$row2['jml_nominal'];

#--------------- FLASH202000020  Foreclose Properties Provision
        $var_flash=" and b.FLASH_Level_3='FLASH202000020' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $foreclose_pp_1=$row2['jml_nominal'];

#--------------- FLASH202000021  Others
        $var_flash=" and b.FLASH_Level_3='FLASH202000021' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $others3_1=$row2['jml_nominal'];

#--------------- FLASH202000023  Corporate Tax
        $var_flash=" and b.FLASH_Level_3='FLASH202000023' ";
        $result2=odbc_exec($connection2, $query.$var_tgl_1.$var_flash.$var_add);
        $row2=odbc_fetch_array($result2);
        $corporate_tax_1=$row2['jml_nominal'];




// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFont = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
//$styleArraybackgroundRed = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignmentCenter = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$objPHPExcel->getActiveSheet()->getStyle('A1:U1')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A1:U1')->applyFromArray($styleArrayAlignment);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(15);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'FLASH_Level_3');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Description');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Data_Date');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Number of Days Actual');
$objPHPExcel->getActiveSheet()->setCellValue('E1', 'End_Date');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Number of Days End_Date');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'Rest of Days');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'Last Actual');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'Actual YTD');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'Actual MTD');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'Average');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'Interest');
$objPHPExcel->getActiveSheet()->setCellValue('M1', 'LDR');
$objPHPExcel->getActiveSheet()->setCellValue('N1', 'Loan');
$objPHPExcel->getActiveSheet()->setCellValue('O1', 'Asumsi_Loan'); //tambahan
$objPHPExcel->getActiveSheet()->setCellValue('P1', 'DPK');
$objPHPExcel->getActiveSheet()->setCellValue('Q1', 'AddLoan');
$objPHPExcel->getActiveSheet()->setCellValue('R1', 'Add_Interest');
$objPHPExcel->getActiveSheet()->setCellValue('S1', 'Asumsi_EIR');
$objPHPExcel->getActiveSheet()->setCellValue('T1', 'Employee_Loan_Benefit');
$objPHPExcel->getActiveSheet()->setCellValue('U1', 'Add');
$objPHPExcel->getActiveSheet()->setCellValue('V1', 'Proyeksi_MTD');
$objPHPExcel->getActiveSheet()->setCellValue('W1', 'YTD Actual Last Month');
$objPHPExcel->getActiveSheet()->setCellValue('X1', 'Proyeksi_YTD'); 

# FLASH_LEVEL_3

$objPHPExcel->getActiveSheet()->setCellValue('A2', 'FLASH201000002');
$objPHPExcel->getActiveSheet()->setCellValue('A3', 'FLASH201000003');
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'FLASH201000004');
$objPHPExcel->getActiveSheet()->setCellValue('A5', 'FLASH201000005');
$objPHPExcel->getActiveSheet()->setCellValue('A6', 'FLASH201000006');
$objPHPExcel->getActiveSheet()->setCellValue('A7', 'FLASH202000002');
$objPHPExcel->getActiveSheet()->setCellValue('A8', 'FLASH202000003');
$objPHPExcel->getActiveSheet()->setCellValue('A9', 'FLASH202000004');
$objPHPExcel->getActiveSheet()->setCellValue('A10', 'FLASH202000005');
$objPHPExcel->getActiveSheet()->setCellValue('A11', 'FLASH202000007');
$objPHPExcel->getActiveSheet()->setCellValue('A12', 'FLASH202000008');
$objPHPExcel->getActiveSheet()->setCellValue('A13', 'FLASH202000009');
$objPHPExcel->getActiveSheet()->setCellValue('A14', 'FLASH201000008');
$objPHPExcel->getActiveSheet()->setCellValue('A15', 'FLASH201000009');
$objPHPExcel->getActiveSheet()->setCellValue('A16', 'FLASH201000010');
$objPHPExcel->getActiveSheet()->setCellValue('A17', 'FLASH201000011');
$objPHPExcel->getActiveSheet()->setCellValue('A18', 'FLASH201000012');
$objPHPExcel->getActiveSheet()->setCellValue('A19', 'FLASH201000013');
$objPHPExcel->getActiveSheet()->setCellValue('A20', 'FLASH201000014');
$objPHPExcel->getActiveSheet()->setCellValue('A21', 'FLASH201000015');
$objPHPExcel->getActiveSheet()->setCellValue('A22', 'FLASH201000016');
$objPHPExcel->getActiveSheet()->setCellValue('A23', 'FLASH201000017');
$objPHPExcel->getActiveSheet()->setCellValue('A24', 'FLASH201000018');
$objPHPExcel->getActiveSheet()->setCellValue('A25', 'FLASH201000019');
$objPHPExcel->getActiveSheet()->setCellValue('A26', 'FLASH202000010');
$objPHPExcel->getActiveSheet()->setCellValue('A27', 'FLASH202000011');
$objPHPExcel->getActiveSheet()->setCellValue('A28', 'FLASH202000012');
$objPHPExcel->getActiveSheet()->setCellValue('A29', 'FLASH202000014');
$objPHPExcel->getActiveSheet()->setCellValue('A30', 'FLASH202000015');
$objPHPExcel->getActiveSheet()->setCellValue('A31', 'FLASH202000016');
$objPHPExcel->getActiveSheet()->setCellValue('A32', 'FLASH202000017');
$objPHPExcel->getActiveSheet()->setCellValue('A33', 'FLASH202000018');
$objPHPExcel->getActiveSheet()->setCellValue('A34', 'FLASH202000019');
$objPHPExcel->getActiveSheet()->setCellValue('A35', 'FLASH202000020');
$objPHPExcel->getActiveSheet()->setCellValue('A36', 'FLASH202000021');
$objPHPExcel->getActiveSheet()->setCellValue('A37', 'FLASH202000023');

# DESCRIPTION
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Loan');
$objPHPExcel->getActiveSheet()->setCellValue('B3', 'Treasury bills');
$objPHPExcel->getActiveSheet()->setCellValue('B4', 'Interbank placements');
$objPHPExcel->getActiveSheet()->setCellValue('B5', 'Placement with BI');
$objPHPExcel->getActiveSheet()->setCellValue('B6', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'Current accounts');
$objPHPExcel->getActiveSheet()->setCellValue('B8', 'Saving accounts');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Time deposits');
$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Bank deposits');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Borrowings (MCB)');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Guaranteed premium');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Forex gain/(loss) on transactions');
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Gain/(Loss) on sale of securities/bonds');
$objPHPExcel->getActiveSheet()->setCellValue('B16', 'Remittance fee');
$objPHPExcel->getActiveSheet()->setCellValue('B17', 'Trade Finance fee');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Processing fee');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Credit Card fee');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Insurance Fee');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Service Charges');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Other Commission & Fee ');
$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Penalty');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'Write Offs Recovered');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'Other Income');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'Staff Cost');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'General & Administrative Expenses');
$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Depreciation');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Other Operating Expense/income');
$objPHPExcel->getActiveSheet()->setCellValue('B30', 'General Provision');
$objPHPExcel->getActiveSheet()->setCellValue('B31', 'Specific Provision Charged');
$objPHPExcel->getActiveSheet()->setCellValue('B32', 'Specific Provision Recovery ');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'Write Offs Charged');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Write Offs Recovered');
$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Foreclose Properties Provision');
$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Others');
$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Corporate Tax');

# DATA DATE
$objPHPExcel->getActiveSheet()->setCellValue('C2', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C3', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C4', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C5', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C6', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C7', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C8', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C9', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C10', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C11', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C12', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C13', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C14', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C15', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C16', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C17', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C18', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C19', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C20', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C21', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C22', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C23', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C24', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C25', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C26', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C27', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C28', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C29', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C30', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C31', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C32', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C33', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C34', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C35', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C36', $curr_tgl);
$objPHPExcel->getActiveSheet()->setCellValue('C37', $curr_tgl);

# NUMBER OF DAY ACTUAL
$objPHPExcel->getActiveSheet()->setCellValue('D2', '=DAY(C2)');
$objPHPExcel->getActiveSheet()->setCellValue('D3', '=DAY(C3)');
$objPHPExcel->getActiveSheet()->setCellValue('D4', '=DAY(C4)');
$objPHPExcel->getActiveSheet()->setCellValue('D5', '=DAY(C5)');
$objPHPExcel->getActiveSheet()->setCellValue('D6', '=DAY(C6)');
$objPHPExcel->getActiveSheet()->setCellValue('D7', '=DAY(C7)');
$objPHPExcel->getActiveSheet()->setCellValue('D8', '=DAY(C8)');
$objPHPExcel->getActiveSheet()->setCellValue('D9', '=DAY(C9)');
$objPHPExcel->getActiveSheet()->setCellValue('D10', '=DAY(C10)');
$objPHPExcel->getActiveSheet()->setCellValue('D11', '=DAY(C11)');
$objPHPExcel->getActiveSheet()->setCellValue('D12', '=DAY(C12)');
$objPHPExcel->getActiveSheet()->setCellValue('D13', '=DAY(C13)');
$objPHPExcel->getActiveSheet()->setCellValue('D14', '=DAY(C14)');
$objPHPExcel->getActiveSheet()->setCellValue('D15', '=DAY(C15)');
$objPHPExcel->getActiveSheet()->setCellValue('D16', '=DAY(C16)');
$objPHPExcel->getActiveSheet()->setCellValue('D17', '=DAY(C17)');
$objPHPExcel->getActiveSheet()->setCellValue('D18', '=DAY(C18)');
$objPHPExcel->getActiveSheet()->setCellValue('D19', '=DAY(C19)');
$objPHPExcel->getActiveSheet()->setCellValue('D20', '=DAY(C20)');
$objPHPExcel->getActiveSheet()->setCellValue('D21', '=DAY(C21)');
$objPHPExcel->getActiveSheet()->setCellValue('D22', '=DAY(C22)');
$objPHPExcel->getActiveSheet()->setCellValue('D23', '=DAY(C23)');
$objPHPExcel->getActiveSheet()->setCellValue('D24', '=DAY(C24)');
$objPHPExcel->getActiveSheet()->setCellValue('D25', '=DAY(C25)');
$objPHPExcel->getActiveSheet()->setCellValue('D26', '=DAY(C26)');
$objPHPExcel->getActiveSheet()->setCellValue('D27', '=DAY(C27)');
$objPHPExcel->getActiveSheet()->setCellValue('D28', '=DAY(C28)');
$objPHPExcel->getActiveSheet()->setCellValue('D29', '=DAY(C29)');
$objPHPExcel->getActiveSheet()->setCellValue('D30', '=DAY(C30)');
$objPHPExcel->getActiveSheet()->setCellValue('D31', '=DAY(C31)');
$objPHPExcel->getActiveSheet()->setCellValue('D32', '=DAY(C32)');
$objPHPExcel->getActiveSheet()->setCellValue('D33', '=DAY(C33)');
$objPHPExcel->getActiveSheet()->setCellValue('D34', '=DAY(C34)');
$objPHPExcel->getActiveSheet()->setCellValue('D35', '=DAY(C35)');
$objPHPExcel->getActiveSheet()->setCellValue('D36', '=DAY(C36)');
$objPHPExcel->getActiveSheet()->setCellValue('D37', '=DAY(C37)');
# END DATE 
$objPHPExcel->getActiveSheet()->setCellValue('E2', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E3', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E4', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E5', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E6', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E7', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E8', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E9', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E10', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E11', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E12', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E13', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E14', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E15', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E16', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E17', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E18', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E19', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E20', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E21', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E22', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E23', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E24', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E25', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E26', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E27', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E28', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E29', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E30', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E31', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E32', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E33', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E34', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E35', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E36', $tgl_terakhir);
$objPHPExcel->getActiveSheet()->setCellValue('E37', $tgl_terakhir);
# NUMBER OF DAY END DATE
$objPHPExcel->getActiveSheet()->setCellValue('F2', '=DAY(E2)');
$objPHPExcel->getActiveSheet()->setCellValue('F3', '=DAY(E3)');
$objPHPExcel->getActiveSheet()->setCellValue('F4', '=DAY(E4)');
$objPHPExcel->getActiveSheet()->setCellValue('F5', '=DAY(E5)');
$objPHPExcel->getActiveSheet()->setCellValue('F6', '=DAY(E6)');
$objPHPExcel->getActiveSheet()->setCellValue('F7', '=DAY(E7)');
$objPHPExcel->getActiveSheet()->setCellValue('F8', '=DAY(E8)');
$objPHPExcel->getActiveSheet()->setCellValue('F9', '=DAY(E9)');
$objPHPExcel->getActiveSheet()->setCellValue('F10', '=DAY(E10)');
$objPHPExcel->getActiveSheet()->setCellValue('F11', '=DAY(E11)');
$objPHPExcel->getActiveSheet()->setCellValue('F12', '=DAY(E12)');
$objPHPExcel->getActiveSheet()->setCellValue('F13', '=DAY(E13)');
$objPHPExcel->getActiveSheet()->setCellValue('F14', '=DAY(E14)');
$objPHPExcel->getActiveSheet()->setCellValue('F15', '=DAY(E15)');
$objPHPExcel->getActiveSheet()->setCellValue('F16', '=DAY(E16)');
$objPHPExcel->getActiveSheet()->setCellValue('F17', '=DAY(E17)');
$objPHPExcel->getActiveSheet()->setCellValue('F18', '=DAY(E18)');
$objPHPExcel->getActiveSheet()->setCellValue('F19', '=DAY(E19)');
$objPHPExcel->getActiveSheet()->setCellValue('F20', '=DAY(E20)');
$objPHPExcel->getActiveSheet()->setCellValue('F21', '=DAY(E21)');
$objPHPExcel->getActiveSheet()->setCellValue('F22', '=DAY(E22)');
$objPHPExcel->getActiveSheet()->setCellValue('F23', '=DAY(E23)');
$objPHPExcel->getActiveSheet()->setCellValue('F24', '=DAY(E24)');
$objPHPExcel->getActiveSheet()->setCellValue('F25', '=DAY(E25)');
$objPHPExcel->getActiveSheet()->setCellValue('F26', '=DAY(E26)');
$objPHPExcel->getActiveSheet()->setCellValue('F27', '=DAY(E27)');
$objPHPExcel->getActiveSheet()->setCellValue('F28', '=DAY(E28)');
$objPHPExcel->getActiveSheet()->setCellValue('F29', '=DAY(E29)');
$objPHPExcel->getActiveSheet()->setCellValue('F30', '=DAY(E30)');
$objPHPExcel->getActiveSheet()->setCellValue('F31', '=DAY(E31)');
$objPHPExcel->getActiveSheet()->setCellValue('F32', '=DAY(E32)');
$objPHPExcel->getActiveSheet()->setCellValue('F33', '=DAY(E33)');
$objPHPExcel->getActiveSheet()->setCellValue('F34', '=DAY(E34)');
$objPHPExcel->getActiveSheet()->setCellValue('F35', '=DAY(E35)');
$objPHPExcel->getActiveSheet()->setCellValue('F36', '=DAY(E36)');
$objPHPExcel->getActiveSheet()->setCellValue('F37', '=DAY(E37)');
# REST OF DAYS
$objPHPExcel->getActiveSheet()->setCellValue('G2', '=F2-D2');
$objPHPExcel->getActiveSheet()->setCellValue('G3', '=F3-D3');
$objPHPExcel->getActiveSheet()->setCellValue('G4', '=F4-D4');
$objPHPExcel->getActiveSheet()->setCellValue('G5', '=F5-D5');
$objPHPExcel->getActiveSheet()->setCellValue('G6', '=F6-D6');
$objPHPExcel->getActiveSheet()->setCellValue('G7', '=F7-D7');
$objPHPExcel->getActiveSheet()->setCellValue('G8', '=F8-D8');
$objPHPExcel->getActiveSheet()->setCellValue('G9', '=F9-D9');
$objPHPExcel->getActiveSheet()->setCellValue('G10', '=F10-D10');
$objPHPExcel->getActiveSheet()->setCellValue('G11', '=F11-D11');
$objPHPExcel->getActiveSheet()->setCellValue('G12', '=F12-D12');
$objPHPExcel->getActiveSheet()->setCellValue('G13', '=F13-D13');
$objPHPExcel->getActiveSheet()->setCellValue('G14', '=F14-D14');
$objPHPExcel->getActiveSheet()->setCellValue('G15', '=F15-D15');
$objPHPExcel->getActiveSheet()->setCellValue('G16', '=F16-D16');
$objPHPExcel->getActiveSheet()->setCellValue('G17', '=F17-D17');
$objPHPExcel->getActiveSheet()->setCellValue('G18', '=F18-D18');
$objPHPExcel->getActiveSheet()->setCellValue('G19', '=F19-D19');
$objPHPExcel->getActiveSheet()->setCellValue('G20', '=F20-D20');
$objPHPExcel->getActiveSheet()->setCellValue('G21', '=F21-D21');
$objPHPExcel->getActiveSheet()->setCellValue('G22', '=F22-D22');
$objPHPExcel->getActiveSheet()->setCellValue('G23', '=F23-D23');
$objPHPExcel->getActiveSheet()->setCellValue('G24', '=F24-D24');
$objPHPExcel->getActiveSheet()->setCellValue('G25', '=F25-D25');
$objPHPExcel->getActiveSheet()->setCellValue('G26', '=F26-D26');
$objPHPExcel->getActiveSheet()->setCellValue('G27', '=F27-D27');
$objPHPExcel->getActiveSheet()->setCellValue('G28', '=F28-D28');
$objPHPExcel->getActiveSheet()->setCellValue('G29', '=F29-D29');
$objPHPExcel->getActiveSheet()->setCellValue('G30', '=F30-D30');
$objPHPExcel->getActiveSheet()->setCellValue('G31', '=F31-D31');
$objPHPExcel->getActiveSheet()->setCellValue('G32', '=F32-D32');
$objPHPExcel->getActiveSheet()->setCellValue('G33', '=F33-D33');
$objPHPExcel->getActiveSheet()->setCellValue('G34', '=F34-D34');
$objPHPExcel->getActiveSheet()->setCellValue('G35', '=F35-D35');
$objPHPExcel->getActiveSheet()->setCellValue('G36', '=F36-D36');
$objPHPExcel->getActiveSheet()->setCellValue('G37', '=F37-D37');
#LAST ACTUAL YTD'
$objPHPExcel->getActiveSheet()->setCellValue('H2', $loan_1);
$objPHPExcel->getActiveSheet()->setCellValue('H3', $treasury_bill_1);
$objPHPExcel->getActiveSheet()->setCellValue('H4', $interbank_placements_1);
$objPHPExcel->getActiveSheet()->setCellValue('H5', $placement_wbi_1);
$objPHPExcel->getActiveSheet()->setCellValue('H6', $others1_1);
$objPHPExcel->getActiveSheet()->setCellValue('H7', $current_account_1);
$objPHPExcel->getActiveSheet()->setCellValue('H8', $saving_account_1);
$objPHPExcel->getActiveSheet()->setCellValue('H9', $time_deposits_1);
$objPHPExcel->getActiveSheet()->setCellValue('H10', $bank_deposits_1);
$objPHPExcel->getActiveSheet()->setCellValue('H11', $borrowing_mcb_1);
$objPHPExcel->getActiveSheet()->setCellValue('H12', $guaranted_premium_1);
$objPHPExcel->getActiveSheet()->setCellValue('H13', $others2_1);
$objPHPExcel->getActiveSheet()->setCellValue('H14', $forex_gain_1);
$objPHPExcel->getActiveSheet()->setCellValue('H15', $gain_loss_1);
$objPHPExcel->getActiveSheet()->setCellValue('H16', $remittance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('H17', $trade_finance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('H18', $processing_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('H19', $credit_card_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('H20', $insurance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('H21', $service_charges_1);
$objPHPExcel->getActiveSheet()->setCellValue('H22', $other_commission_1);
$objPHPExcel->getActiveSheet()->setCellValue('H23', $penalty_1);
$objPHPExcel->getActiveSheet()->setCellValue('H24', $write_off_r_1);
$objPHPExcel->getActiveSheet()->setCellValue('H25', $other_income_1);
$objPHPExcel->getActiveSheet()->setCellValue('H26', $staff_cost_1);
$objPHPExcel->getActiveSheet()->setCellValue('H27', $general_administrative_1);
$objPHPExcel->getActiveSheet()->setCellValue('H28', $depreciation_1);
$objPHPExcel->getActiveSheet()->setCellValue('H29', $other_operating_1);
$objPHPExcel->getActiveSheet()->setCellValue('H30', $general_provision_1);
$objPHPExcel->getActiveSheet()->setCellValue('H31', $specific_provision_charge_1);
$objPHPExcel->getActiveSheet()->setCellValue('H32', $specific_provision_recovery_1);
$objPHPExcel->getActiveSheet()->setCellValue('H33', $write_off_c_1);
$objPHPExcel->getActiveSheet()->setCellValue('H34', $write_off_r2_1);
$objPHPExcel->getActiveSheet()->setCellValue('H35', $foreclose_pp_1);
$objPHPExcel->getActiveSheet()->setCellValue('H36', $others3_1);
$objPHPExcel->getActiveSheet()->setCellValue('H37', $corporate_tax_1);
#ACTUAL YTD
$objPHPExcel->getActiveSheet()->setCellValue('I2', $loan);
$objPHPExcel->getActiveSheet()->setCellValue('I3', $treasury_bill);
$objPHPExcel->getActiveSheet()->setCellValue('I4', $interbank_placements);
$objPHPExcel->getActiveSheet()->setCellValue('I5', $placement_wbi);
$objPHPExcel->getActiveSheet()->setCellValue('I6', $others1);
$objPHPExcel->getActiveSheet()->setCellValue('I7', $current_account);
$objPHPExcel->getActiveSheet()->setCellValue('I8', $saving_account);
$objPHPExcel->getActiveSheet()->setCellValue('I9', $time_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('I10', $bank_deposits);
$objPHPExcel->getActiveSheet()->setCellValue('I11', $borrowing_mcb);
$objPHPExcel->getActiveSheet()->setCellValue('I12', $guaranted_premium);
$objPHPExcel->getActiveSheet()->setCellValue('I13', $others2);
$objPHPExcel->getActiveSheet()->setCellValue('I14', $forex_gain);
$objPHPExcel->getActiveSheet()->setCellValue('I15', $gain_loss);
$objPHPExcel->getActiveSheet()->setCellValue('I16', $remittance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('I17', $trade_finance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $processing_fee);
$objPHPExcel->getActiveSheet()->setCellValue('I19', $credit_card_fee);
$objPHPExcel->getActiveSheet()->setCellValue('I20', $insurance_fee);
$objPHPExcel->getActiveSheet()->setCellValue('I21', $service_charges);
$objPHPExcel->getActiveSheet()->setCellValue('I22', $other_commission);
$objPHPExcel->getActiveSheet()->setCellValue('I23', $penalty);
$objPHPExcel->getActiveSheet()->setCellValue('I24', $write_off_r);
$objPHPExcel->getActiveSheet()->setCellValue('I25', $other_income);
$objPHPExcel->getActiveSheet()->setCellValue('I26', $staff_cost);
$objPHPExcel->getActiveSheet()->setCellValue('I27', $general_administrative);
$objPHPExcel->getActiveSheet()->setCellValue('I28', $depreciation);
$objPHPExcel->getActiveSheet()->setCellValue('I29', $other_operating);
$objPHPExcel->getActiveSheet()->setCellValue('I30', $general_provision);
$objPHPExcel->getActiveSheet()->setCellValue('I31', $specific_provision_charge);
$objPHPExcel->getActiveSheet()->setCellValue('I32', $specific_provision_recovery);
$objPHPExcel->getActiveSheet()->setCellValue('I33', $write_off_c);
$objPHPExcel->getActiveSheet()->setCellValue('I34', $write_off_r2);
$objPHPExcel->getActiveSheet()->setCellValue('I35', $foreclose_pp);
$objPHPExcel->getActiveSheet()->setCellValue('I36', $others3);
$objPHPExcel->getActiveSheet()->setCellValue('I37', $corporate_tax);
#ACTUAL MTD
$objPHPExcel->getActiveSheet()->setCellValue('J2', "=I2-H2");
$objPHPExcel->getActiveSheet()->setCellValue('J3', "=I3-H3");
$objPHPExcel->getActiveSheet()->setCellValue('J4', "=I4-H4");
$objPHPExcel->getActiveSheet()->setCellValue('J5', "=I5-H5");
$objPHPExcel->getActiveSheet()->setCellValue('J6', "=I6-H6");
$objPHPExcel->getActiveSheet()->setCellValue('J7', "=I7-H7");
$objPHPExcel->getActiveSheet()->setCellValue('J8', "=I8-H8");
$objPHPExcel->getActiveSheet()->setCellValue('J9', "=I9-H9");
$objPHPExcel->getActiveSheet()->setCellValue('J10', "=I10-H10");
$objPHPExcel->getActiveSheet()->setCellValue('J11', "=I11-H11");
$objPHPExcel->getActiveSheet()->setCellValue('J12', "=I12-H12");
$objPHPExcel->getActiveSheet()->setCellValue('J13', "=I13-H13");
$objPHPExcel->getActiveSheet()->setCellValue('J14', "=I14-H14");
$objPHPExcel->getActiveSheet()->setCellValue('J15', "=I15-H15");
$objPHPExcel->getActiveSheet()->setCellValue('J16', "=I16-H16");
$objPHPExcel->getActiveSheet()->setCellValue('J17', "=I17-H17");
$objPHPExcel->getActiveSheet()->setCellValue('J18', "=I18-H18");
$objPHPExcel->getActiveSheet()->setCellValue('J19', "=I19-H19");
$objPHPExcel->getActiveSheet()->setCellValue('J20', "=I20-H20");
$objPHPExcel->getActiveSheet()->setCellValue('J21', "=I21-H21");
$objPHPExcel->getActiveSheet()->setCellValue('J22', "=I22-H22");
$objPHPExcel->getActiveSheet()->setCellValue('J23', "=I23-H23");
$objPHPExcel->getActiveSheet()->setCellValue('J24', "=I24-H24");
$objPHPExcel->getActiveSheet()->setCellValue('J25', "=I25-H25");
$objPHPExcel->getActiveSheet()->setCellValue('J26', "=I26-H26");
$objPHPExcel->getActiveSheet()->setCellValue('J27', "=I27-H27");
$objPHPExcel->getActiveSheet()->setCellValue('J28', "=I28-H28");
$objPHPExcel->getActiveSheet()->setCellValue('J29', "=I29-H29");
$objPHPExcel->getActiveSheet()->setCellValue('J30', "=I30-H30");
$objPHPExcel->getActiveSheet()->setCellValue('J31', "=I31-H31");
$objPHPExcel->getActiveSheet()->setCellValue('J32', "=I32-H32");
$objPHPExcel->getActiveSheet()->setCellValue('J33', "=I33-H33");
$objPHPExcel->getActiveSheet()->setCellValue('J34', "=I34-H34");
$objPHPExcel->getActiveSheet()->setCellValue('J35', "=I35-H35");
$objPHPExcel->getActiveSheet()->setCellValue('J36', "=I36-H36");
$objPHPExcel->getActiveSheet()->setCellValue('J37', "=I37-H37");
# I (AVERAGE)
$objPHPExcel->getActiveSheet()->setCellValue('K2', '=J2/D2');
$objPHPExcel->getActiveSheet()->setCellValue('K3', '=J3/D3');
$objPHPExcel->getActiveSheet()->setCellValue('K4', '=J4/D4');
$objPHPExcel->getActiveSheet()->setCellValue('K5', '=J5/D5');
$objPHPExcel->getActiveSheet()->setCellValue('K6', '=J6/D6');
$objPHPExcel->getActiveSheet()->setCellValue('K7', '=J7/D7');
$objPHPExcel->getActiveSheet()->setCellValue('K8', '=J8/D8');
$objPHPExcel->getActiveSheet()->setCellValue('K9', '=J9/D9');
$objPHPExcel->getActiveSheet()->setCellValue('K10', '=J10/D10');
$objPHPExcel->getActiveSheet()->setCellValue('K11', '=J11/D11');
$objPHPExcel->getActiveSheet()->setCellValue('K12', '=J12/D12');
$objPHPExcel->getActiveSheet()->setCellValue('K13', '=J13/D13');
$objPHPExcel->getActiveSheet()->setCellValue('K14', '=J14/D14');
$objPHPExcel->getActiveSheet()->setCellValue('K15', '=J15/D15');
$objPHPExcel->getActiveSheet()->setCellValue('K16', '=J16/D16');
$objPHPExcel->getActiveSheet()->setCellValue('K17', '=J17/D17');
$objPHPExcel->getActiveSheet()->setCellValue('K18', '=J18/D18');
$objPHPExcel->getActiveSheet()->setCellValue('K19', '=J19/D19');
$objPHPExcel->getActiveSheet()->setCellValue('K20', '=J20/D20');
$objPHPExcel->getActiveSheet()->setCellValue('K21', '=J21/D21');
$objPHPExcel->getActiveSheet()->setCellValue('K22', '=J22/D22');
$objPHPExcel->getActiveSheet()->setCellValue('K23', '=J23/D23');
$objPHPExcel->getActiveSheet()->setCellValue('K24', '=J24/D24');
$objPHPExcel->getActiveSheet()->setCellValue('K25', '=J25/D25');
$objPHPExcel->getActiveSheet()->setCellValue('K26', '=J26/D26');
$objPHPExcel->getActiveSheet()->setCellValue('K27', '=J27/D27');
$objPHPExcel->getActiveSheet()->setCellValue('K28', '=J28/D28');
$objPHPExcel->getActiveSheet()->setCellValue('K29', '=J29/D29');
$objPHPExcel->getActiveSheet()->setCellValue('K30', '=J30/D30');
$objPHPExcel->getActiveSheet()->setCellValue('K31', '=J31/D31');
$objPHPExcel->getActiveSheet()->setCellValue('K32', '=J32/D32');
$objPHPExcel->getActiveSheet()->setCellValue('K33', '=J33/D33');
$objPHPExcel->getActiveSheet()->setCellValue('K34', '=J34/D34');
$objPHPExcel->getActiveSheet()->setCellValue('K35', '=J35/D35');
$objPHPExcel->getActiveSheet()->setCellValue('K36', '=J36/D36');
$objPHPExcel->getActiveSheet()->setCellValue('K37', '=J37/D37');

# L --> Loan
/*
$objPHPExcel->getActiveSheet()->setCellValue('L2', '=H2-(K2*H2)');
$objPHPExcel->getActiveSheet()->setCellValue('L3', '=H3-(K3*H3)');
$objPHPExcel->getActiveSheet()->setCellValue('L4', '=H4-(K4*H4)');
$objPHPExcel->getActiveSheet()->setCellValue('L5', '=H5-(K5*H5)');
$objPHPExcel->getActiveSheet()->setCellValue('L6', '=H6-(K6*H6)');
$objPHPExcel->getActiveSheet()->setCellValue('L7', '=H7-(K7*H7)');
$objPHPExcel->getActiveSheet()->setCellValue('L8', '=H8-(K8*H8)');
$objPHPExcel->getActiveSheet()->setCellValue('L9', '=H9-(K9*H9)');
$objPHPExcel->getActiveSheet()->setCellValue('L10', '=H10-(K10*H10)');
$objPHPExcel->getActiveSheet()->setCellValue('L11', '=H11-(K11*H11)');
$objPHPExcel->getActiveSheet()->setCellValue('L12', '=H12-(K12*H12)');
$objPHPExcel->getActiveSheet()->setCellValue('L13', '=H13-(K13*H13)');
$objPHPExcel->getActiveSheet()->setCellValue('L14', '=H14-(K14*H14)');
$objPHPExcel->getActiveSheet()->setCellValue('L15', '=H15-(K15*H15)');
$objPHPExcel->getActiveSheet()->setCellValue('L16', '=H16-(K16*H16)');
$objPHPExcel->getActiveSheet()->setCellValue('L17', '=H17-(K17*H17)');
$objPHPExcel->getActiveSheet()->setCellValue('L18', '=H18-(K18*H18)');
$objPHPExcel->getActiveSheet()->setCellValue('L19', '=H19-(K19*H19)');
$objPHPExcel->getActiveSheet()->setCellValue('L20', '=H20-(K20*H20)');
$objPHPExcel->getActiveSheet()->setCellValue('L21', '=H21-(K21*H21)');
$objPHPExcel->getActiveSheet()->setCellValue('L22', '=H22-(K22*H22)');
$objPHPExcel->getActiveSheet()->setCellValue('L23', '=H23-(K23*H23)');
$objPHPExcel->getActiveSheet()->setCellValue('L24', '=H24-(K24*H24)');
$objPHPExcel->getActiveSheet()->setCellValue('L25', '=H25-(K25*H25)');
$objPHPExcel->getActiveSheet()->setCellValue('L26', '=H26-(K26*H26)');
$objPHPExcel->getActiveSheet()->setCellValue('L27', '=H27-(K27*H27)');
$objPHPExcel->getActiveSheet()->setCellValue('L28', '=H28-(K28*H28)');
$objPHPExcel->getActiveSheet()->setCellValue('L29', '=H29-(K29*H29)');
$objPHPExcel->getActiveSheet()->setCellValue('L30', '=H30-(K30*H30)');
$objPHPExcel->getActiveSheet()->setCellValue('L31', '=H31-(K31*H31)');
$objPHPExcel->getActiveSheet()->setCellValue('L32', '=H32-(K32*H32)');
$objPHPExcel->getActiveSheet()->setCellValue('L33', '=H33-(K33*H33)');
$objPHPExcel->getActiveSheet()->setCellValue('L34', '=H34-(K34*H34)');
$objPHPExcel->getActiveSheet()->setCellValue('L35', '=H35-(K35*H35)');
$objPHPExcel->getActiveSheet()->setCellValue('L36', '=H36-(K36*H36)');
$objPHPExcel->getActiveSheet()->setCellValue('L37', '=H37-(K37*H37)');
*/

# M -->DPK
/*
$objPHPExcel->getActiveSheet()->setCellValue('M2', '=H2-(K2*H2)');
$objPHPExcel->getActiveSheet()->setCellValue('M3', '=H3-(K3*H3)');
$objPHPExcel->getActiveSheet()->setCellValue('M4', '=H4-(K4*H4)');
$objPHPExcel->getActiveSheet()->setCellValue('M5', '=H5-(K5*H5)');
$objPHPExcel->getActiveSheet()->setCellValue('M6', '=H6-(K6*H6)');
$objPHPExcel->getActiveSheet()->setCellValue('M7', '=H7-(K7*H7)');
$objPHPExcel->getActiveSheet()->setCellValue('M8', '=H8-(K8*H8)');
$objPHPExcel->getActiveSheet()->setCellValue('M9', '=H9-(K9*H9)');
$objPHPExcel->getActiveSheet()->setCellValue('M10', '=H10-(K10*H10)');
$objPHPExcel->getActiveSheet()->setCellValue('M11', '=H11-(K11*H11)');
$objPHPExcel->getActiveSheet()->setCellValue('M12', '=H12-(K12*H12)');
$objPHPExcel->getActiveSheet()->setCellValue('M13', '=H13-(K13*H13)');
$objPHPExcel->getActiveSheet()->setCellValue('M14', '=H14-(K14*H14)');
$objPHPExcel->getActiveSheet()->setCellValue('M15', '=H15-(K15*H15)');
$objPHPExcel->getActiveSheet()->setCellValue('M16', '=H16-(K16*H16)');
$objPHPExcel->getActiveSheet()->setCellValue('M17', '=H17-(K17*H17)');
$objPHPExcel->getActiveSheet()->setCellValue('M18', '=H18-(K18*H18)');
$objPHPExcel->getActiveSheet()->setCellValue('M19', '=H19-(K19*H19)');
$objPHPExcel->getActiveSheet()->setCellValue('M20', '=H20-(K20*H20)');
$objPHPExcel->getActiveSheet()->setCellValue('M21', '=H21-(K21*H21)');
$objPHPExcel->getActiveSheet()->setCellValue('M22', '=H22-(K22*H22)');
$objPHPExcel->getActiveSheet()->setCellValue('M23', '=H23-(K23*H23)');
$objPHPExcel->getActiveSheet()->setCellValue('M24', '=H24-(K24*H24)');
$objPHPExcel->getActiveSheet()->setCellValue('M25', '=H25-(K25*H25)');
$objPHPExcel->getActiveSheet()->setCellValue('M26', '=H26-(K26*H26)');
$objPHPExcel->getActiveSheet()->setCellValue('M27', '=H27-(K27*H27)');
$objPHPExcel->getActiveSheet()->setCellValue('M28', '=H28-(K28*H28)');
$objPHPExcel->getActiveSheet()->setCellValue('M29', '=H29-(K29*H29)');
$objPHPExcel->getActiveSheet()->setCellValue('M30', '=H30-(K30*H30)');
$objPHPExcel->getActiveSheet()->setCellValue('M31', '=H31-(K31*H31)');
$objPHPExcel->getActiveSheet()->setCellValue('M32', '=H32-(K32*H32)');
$objPHPExcel->getActiveSheet()->setCellValue('M33', '=H33-(K33*H33)');
$objPHPExcel->getActiveSheet()->setCellValue('M34', '=H34-(K34*H34)');
$objPHPExcel->getActiveSheet()->setCellValue('M35', '=H35-(K35*H35)');
$objPHPExcel->getActiveSheet()->setCellValue('M36', '=H36-(K36*H36)');
$objPHPExcel->getActiveSheet()->setCellValue('M37', '=H37-(K37*H37)');
*/

# O --> Asumsi Loan
$objPHPExcel->getActiveSheet()->setCellValue('O2', '=M2*P2');
$objPHPExcel->getActiveSheet()->setCellValue('O3', '=M3*P3');
$objPHPExcel->getActiveSheet()->setCellValue('O4', '=M4*P4');
$objPHPExcel->getActiveSheet()->setCellValue('O5', '=M5*P5');
$objPHPExcel->getActiveSheet()->setCellValue('O6', '=M6*P6');
$objPHPExcel->getActiveSheet()->setCellValue('O7', '=M7*P7');
$objPHPExcel->getActiveSheet()->setCellValue('O8', '=M8*P8');
$objPHPExcel->getActiveSheet()->setCellValue('O9', '=M9*P9');
$objPHPExcel->getActiveSheet()->setCellValue('O10', '=M10*P10');
$objPHPExcel->getActiveSheet()->setCellValue('O11', '=M11*P11');
$objPHPExcel->getActiveSheet()->setCellValue('O12', '=M12*P12');
$objPHPExcel->getActiveSheet()->setCellValue('O13', '=M13*P13');
$objPHPExcel->getActiveSheet()->setCellValue('O14', '=M14*P14');
$objPHPExcel->getActiveSheet()->setCellValue('O15', '=M15*P15');
$objPHPExcel->getActiveSheet()->setCellValue('O16', '=M16*P16');
$objPHPExcel->getActiveSheet()->setCellValue('O17', '=M17*P17');
$objPHPExcel->getActiveSheet()->setCellValue('O18', '=M18*P18');
$objPHPExcel->getActiveSheet()->setCellValue('O19', '=M19*P19');
$objPHPExcel->getActiveSheet()->setCellValue('O20', '=M20*P20');
$objPHPExcel->getActiveSheet()->setCellValue('O21', '=M21*P21');
$objPHPExcel->getActiveSheet()->setCellValue('O22', '=M22*P22');
$objPHPExcel->getActiveSheet()->setCellValue('O23', '=M23*P23');
$objPHPExcel->getActiveSheet()->setCellValue('O24', '=M24*P24');
$objPHPExcel->getActiveSheet()->setCellValue('O25', '=M25*P25');
$objPHPExcel->getActiveSheet()->setCellValue('O26', '=M26*P26');
$objPHPExcel->getActiveSheet()->setCellValue('O27', '=M27*P27');
$objPHPExcel->getActiveSheet()->setCellValue('O28', '=M28*P28');
$objPHPExcel->getActiveSheet()->setCellValue('O29', '=M29*P29');
$objPHPExcel->getActiveSheet()->setCellValue('O30', '=M30*P30');
$objPHPExcel->getActiveSheet()->setCellValue('O31', '=M31*P31');
$objPHPExcel->getActiveSheet()->setCellValue('O32', '=M32*P32');
$objPHPExcel->getActiveSheet()->setCellValue('O33', '=M33*P33');
$objPHPExcel->getActiveSheet()->setCellValue('O34', '=M34*P34');
$objPHPExcel->getActiveSheet()->setCellValue('O35', '=M35*P35');
$objPHPExcel->getActiveSheet()->setCellValue('O36', '=M36*P36');
$objPHPExcel->getActiveSheet()->setCellValue('O37', '=M37*P37');
# Q --> AddLoan
$objPHPExcel->getActiveSheet()->setCellValue('Q2', '=O2-N2');
$objPHPExcel->getActiveSheet()->setCellValue('Q3', '=O3-N3');
$objPHPExcel->getActiveSheet()->setCellValue('Q4', '=O4-N4');
$objPHPExcel->getActiveSheet()->setCellValue('Q5', '=O5-N5');
$objPHPExcel->getActiveSheet()->setCellValue('Q6', '=O6-N6');
$objPHPExcel->getActiveSheet()->setCellValue('Q7', '=O7-N7');
$objPHPExcel->getActiveSheet()->setCellValue('Q8', '=O8-N8');
$objPHPExcel->getActiveSheet()->setCellValue('Q9', '=O9-N9');
$objPHPExcel->getActiveSheet()->setCellValue('Q10', '=O10-N10');
$objPHPExcel->getActiveSheet()->setCellValue('Q11', '=O11-N11');
$objPHPExcel->getActiveSheet()->setCellValue('Q12', '=O12-N12');
$objPHPExcel->getActiveSheet()->setCellValue('Q13', '=O13-N13');
$objPHPExcel->getActiveSheet()->setCellValue('Q14', '=O14-N14');
$objPHPExcel->getActiveSheet()->setCellValue('Q15', '=O15-N15');
$objPHPExcel->getActiveSheet()->setCellValue('Q16', '=O16-N16');
$objPHPExcel->getActiveSheet()->setCellValue('Q17', '=O17-N17');
$objPHPExcel->getActiveSheet()->setCellValue('Q18', '=O18-N18');
$objPHPExcel->getActiveSheet()->setCellValue('Q19', '=O19-N19');
$objPHPExcel->getActiveSheet()->setCellValue('Q20', '=O20-N20');
$objPHPExcel->getActiveSheet()->setCellValue('Q21', '=O21-N21');
$objPHPExcel->getActiveSheet()->setCellValue('Q22', '=O22-N22');
$objPHPExcel->getActiveSheet()->setCellValue('Q23', '=O23-N23');
$objPHPExcel->getActiveSheet()->setCellValue('Q24', '=O24-N24');
$objPHPExcel->getActiveSheet()->setCellValue('Q25', '=O25-N25');
$objPHPExcel->getActiveSheet()->setCellValue('Q26', '=O26-N26');
$objPHPExcel->getActiveSheet()->setCellValue('Q27', '=O27-N27');
$objPHPExcel->getActiveSheet()->setCellValue('Q28', '=O28-N28');
$objPHPExcel->getActiveSheet()->setCellValue('Q29', '=O29-N29');
$objPHPExcel->getActiveSheet()->setCellValue('Q30', '=O30-N30');
$objPHPExcel->getActiveSheet()->setCellValue('Q31', '=O31-N31');
$objPHPExcel->getActiveSheet()->setCellValue('Q32', '=O32-N32');
$objPHPExcel->getActiveSheet()->setCellValue('Q33', '=O33-N33');
$objPHPExcel->getActiveSheet()->setCellValue('Q34', '=O34-N34');
$objPHPExcel->getActiveSheet()->setCellValue('Q35', '=O35-N35');
$objPHPExcel->getActiveSheet()->setCellValue('Q36', '=O36-N36');
$objPHPExcel->getActiveSheet()->setCellValue('Q37', '=O37-N37');
# R -->Add_Interest
$objPHPExcel->getActiveSheet()->setCellValue('R2', '=((Q2*L2)/360)*G2');
$objPHPExcel->getActiveSheet()->setCellValue('R3', '=((Q3*L3)/360)*G3');
$objPHPExcel->getActiveSheet()->setCellValue('R4', '=((Q4*L4)/360)*G4');
$objPHPExcel->getActiveSheet()->setCellValue('R5', '=((Q5*L5)/360)*G5');
$objPHPExcel->getActiveSheet()->setCellValue('R6', '=((Q6*L6)/360)*G6');
$objPHPExcel->getActiveSheet()->setCellValue('R7', '=((Q7*L7)/360)*G7');
$objPHPExcel->getActiveSheet()->setCellValue('R8', '=((Q8*L8)/360)*G8');
$objPHPExcel->getActiveSheet()->setCellValue('R9', '=((Q9*L9)/360)*G9');
$objPHPExcel->getActiveSheet()->setCellValue('R10', '=((Q10*L10)/360)*G10');
$objPHPExcel->getActiveSheet()->setCellValue('R11', '=((Q11*L11)/360)*G11');
$objPHPExcel->getActiveSheet()->setCellValue('R12', '=((Q12*L12)/360)*G12');
$objPHPExcel->getActiveSheet()->setCellValue('R13', '=((Q13*L13)/360)*G13');
$objPHPExcel->getActiveSheet()->setCellValue('R14', '=((Q14*L14)/360)*G14');
$objPHPExcel->getActiveSheet()->setCellValue('R15', '=((Q15*L15)/360)*G15');
$objPHPExcel->getActiveSheet()->setCellValue('R16', '=((Q16*L16)/360)*G16');
$objPHPExcel->getActiveSheet()->setCellValue('R17', '=((Q17*L17)/360)*G17');
$objPHPExcel->getActiveSheet()->setCellValue('R18', '=((Q18*L18)/360)*G18');
$objPHPExcel->getActiveSheet()->setCellValue('R19', '=((Q19*L19)/360)*G19');
$objPHPExcel->getActiveSheet()->setCellValue('R20', '=((Q20*L20)/360)*G20');
$objPHPExcel->getActiveSheet()->setCellValue('R21', '=((Q21*L21)/360)*G21');
$objPHPExcel->getActiveSheet()->setCellValue('R22', '=((Q22*L22)/360)*G22');
$objPHPExcel->getActiveSheet()->setCellValue('R23', '=((Q23*L23)/360)*G23');
$objPHPExcel->getActiveSheet()->setCellValue('R24', '=((Q24*L24)/360)*G24');
$objPHPExcel->getActiveSheet()->setCellValue('R25', '=((Q25*L25)/360)*G25');
$objPHPExcel->getActiveSheet()->setCellValue('R26', '=((Q26*L26)/360)*G26');
$objPHPExcel->getActiveSheet()->setCellValue('R27', '=((Q27*L27)/360)*G27');
$objPHPExcel->getActiveSheet()->setCellValue('R28', '=((Q28*L28)/360)*G28');
$objPHPExcel->getActiveSheet()->setCellValue('R29', '=((Q29*L29)/360)*G29');
$objPHPExcel->getActiveSheet()->setCellValue('R30', '=((Q30*L30)/360)*G30');
$objPHPExcel->getActiveSheet()->setCellValue('R31', '=((Q31*L31)/360)*G31');
$objPHPExcel->getActiveSheet()->setCellValue('R32', '=((Q32*L32)/360)*G32');
$objPHPExcel->getActiveSheet()->setCellValue('R33', '=((Q33*L33)/360)*G33');
$objPHPExcel->getActiveSheet()->setCellValue('R34', '=((Q34*L34)/360)*G34');
$objPHPExcel->getActiveSheet()->setCellValue('R35', '=((Q35*L35)/360)*G35');
$objPHPExcel->getActiveSheet()->setCellValue('R36', '=((Q36*L36)/360)*G36');
$objPHPExcel->getActiveSheet()->setCellValue('R37', '=((Q37*L37)/360)*G37');
# U --> Add
$objPHPExcel->getActiveSheet()->setCellValue('U2', '=+R2+S2-T2');
$objPHPExcel->getActiveSheet()->setCellValue('U3', '=+R3+S3-T3');
$objPHPExcel->getActiveSheet()->setCellValue('U4', '=+R4+S4-T4');
$objPHPExcel->getActiveSheet()->setCellValue('U5', '=+R5+S5-T5');
$objPHPExcel->getActiveSheet()->setCellValue('U6', '=+R6+S6-T6');
$objPHPExcel->getActiveSheet()->setCellValue('U7', '=+R7+S7-T7');
$objPHPExcel->getActiveSheet()->setCellValue('U8', '=+R8+S8-T8');
$objPHPExcel->getActiveSheet()->setCellValue('U9', '=+R9+S9-T9');
$objPHPExcel->getActiveSheet()->setCellValue('U10', '=+R10+S10-T10');
$objPHPExcel->getActiveSheet()->setCellValue('U11', '=+R11+S11-T11');
$objPHPExcel->getActiveSheet()->setCellValue('U12', '=+R12+S12-T12');
$objPHPExcel->getActiveSheet()->setCellValue('U13', '=+R13+S13-T13');
$objPHPExcel->getActiveSheet()->setCellValue('U14', '=+R14+S14-T14');
$objPHPExcel->getActiveSheet()->setCellValue('U15', '=+R15+S15-T15');
$objPHPExcel->getActiveSheet()->setCellValue('U16', '=+R16+S16-T16');
$objPHPExcel->getActiveSheet()->setCellValue('U17', '=+R17+S17-T17');
$objPHPExcel->getActiveSheet()->setCellValue('U18', '=+R18+S18-T18');
$objPHPExcel->getActiveSheet()->setCellValue('U19', '=+R19+S19-T19');
$objPHPExcel->getActiveSheet()->setCellValue('U20', '=+R20+S20-T20');
$objPHPExcel->getActiveSheet()->setCellValue('U21', '=+R21+S21-T21');
$objPHPExcel->getActiveSheet()->setCellValue('U22', '=+R22+S22-T22');
$objPHPExcel->getActiveSheet()->setCellValue('U23', '=+R23+S23-T23');
$objPHPExcel->getActiveSheet()->setCellValue('U24', '=+R24+S24-T24');
$objPHPExcel->getActiveSheet()->setCellValue('U25', '=+R25+S25-T25');
$objPHPExcel->getActiveSheet()->setCellValue('U26', '=+R26+S26-T26');
$objPHPExcel->getActiveSheet()->setCellValue('U27', '=+R27+S27-T27');
$objPHPExcel->getActiveSheet()->setCellValue('U28', '=+R28+S28-T28');
$objPHPExcel->getActiveSheet()->setCellValue('U29', '=+R29+S29-T29');
$objPHPExcel->getActiveSheet()->setCellValue('U30', '=+R30+S30-T30');
$objPHPExcel->getActiveSheet()->setCellValue('U31', '=+R31+S31-T31');
$objPHPExcel->getActiveSheet()->setCellValue('U32', '=+R32+S32-T32');
$objPHPExcel->getActiveSheet()->setCellValue('U33', '=+R33+S33-T33');
$objPHPExcel->getActiveSheet()->setCellValue('U34', '=+R34+S34-T34');
$objPHPExcel->getActiveSheet()->setCellValue('U35', '=+R35+S35-T35');
$objPHPExcel->getActiveSheet()->setCellValue('U36', '=+R36+S36-T36');
$objPHPExcel->getActiveSheet()->setCellValue('U37', '=+R37+S37-T37');

# V --> Proyeksi_MTD
$objPHPExcel->getActiveSheet()->setCellValue('V2', '=(K2*F2)+U2');
$objPHPExcel->getActiveSheet()->setCellValue('V3', '=(K3*F3)+U3');
$objPHPExcel->getActiveSheet()->setCellValue('V4', '=(K4*F4)+U4');
$objPHPExcel->getActiveSheet()->setCellValue('V5', '=(K5*F5)+U5');
$objPHPExcel->getActiveSheet()->setCellValue('V6', '=(K6*F6)+U6');
$objPHPExcel->getActiveSheet()->setCellValue('V7', '=(K7*F7)+U7');
$objPHPExcel->getActiveSheet()->setCellValue('V8', '=(K8*F8)+U8');
$objPHPExcel->getActiveSheet()->setCellValue('V9', '=(K9*F9)+U9');
$objPHPExcel->getActiveSheet()->setCellValue('V10', '=(K10*F10)+U10');
$objPHPExcel->getActiveSheet()->setCellValue('V11', '=(K11*F11)+U11');
$objPHPExcel->getActiveSheet()->setCellValue('V12', '=(K12*F12)+U12');
$objPHPExcel->getActiveSheet()->setCellValue('V13', '=(K13*F13)+U13');
$objPHPExcel->getActiveSheet()->setCellValue('V14', '=(K14*F14)+U14');
$objPHPExcel->getActiveSheet()->setCellValue('V15', '=(K15*F15)+U15');
$objPHPExcel->getActiveSheet()->setCellValue('V16', '=(K16*F16)+U16');
$objPHPExcel->getActiveSheet()->setCellValue('V17', '=(K17*F17)+U17');
$objPHPExcel->getActiveSheet()->setCellValue('V18', '=(K18*F18)+U18');
$objPHPExcel->getActiveSheet()->setCellValue('V19', '=(K19*F19)+U19');
$objPHPExcel->getActiveSheet()->setCellValue('V20', '=(K20*F20)+U20');
$objPHPExcel->getActiveSheet()->setCellValue('V21', '=(K21*F21)+U21');
$objPHPExcel->getActiveSheet()->setCellValue('V22', '=(K22*F22)+U22');
$objPHPExcel->getActiveSheet()->setCellValue('V23', '=(K23*F23)+U23');
$objPHPExcel->getActiveSheet()->setCellValue('V24', '=(K24*F24)+U24');
$objPHPExcel->getActiveSheet()->setCellValue('V25', '=(K25*F25)+U25');
$objPHPExcel->getActiveSheet()->setCellValue('V26', '=(K26*F26)+U26');
$objPHPExcel->getActiveSheet()->setCellValue('V27', '=(K27*F27)+U27');
$objPHPExcel->getActiveSheet()->setCellValue('V28', '=(K28*F28)+U28');
$objPHPExcel->getActiveSheet()->setCellValue('V29', '=(K29*F29)+U29');
$objPHPExcel->getActiveSheet()->setCellValue('V30', '=(K30*F30)+U30');
$objPHPExcel->getActiveSheet()->setCellValue('V31', '=(K31*F31)+U31');
$objPHPExcel->getActiveSheet()->setCellValue('V32', '=(K32*F32)+U32');
$objPHPExcel->getActiveSheet()->setCellValue('V33', '=(K33*F33)+U33');
$objPHPExcel->getActiveSheet()->setCellValue('V34', '=(K34*F34)+U34');
$objPHPExcel->getActiveSheet()->setCellValue('V35', '=(K35*F35)+U35');
$objPHPExcel->getActiveSheet()->setCellValue('V36', '=(K36*F36)+U36');
$objPHPExcel->getActiveSheet()->setCellValue('V37', '=(K37*F37)+U37');

#W
$objPHPExcel->getActiveSheet()->setCellValue('W2', $loan_1);
$objPHPExcel->getActiveSheet()->setCellValue('W3', $treasury_bill_1);
$objPHPExcel->getActiveSheet()->setCellValue('W4', $interbank_placements_1);
$objPHPExcel->getActiveSheet()->setCellValue('W5', $placement_wbi_1);
$objPHPExcel->getActiveSheet()->setCellValue('W6', $others1_1);
$objPHPExcel->getActiveSheet()->setCellValue('W7', $current_account_1);
$objPHPExcel->getActiveSheet()->setCellValue('W8', $saving_account_1);
$objPHPExcel->getActiveSheet()->setCellValue('W9', $time_deposits_1);
$objPHPExcel->getActiveSheet()->setCellValue('W10', $bank_deposits_1);
$objPHPExcel->getActiveSheet()->setCellValue('W11', $borrowing_mcb_1);
$objPHPExcel->getActiveSheet()->setCellValue('W12', $guaranted_premium_1);
$objPHPExcel->getActiveSheet()->setCellValue('W13', $others2_1);
$objPHPExcel->getActiveSheet()->setCellValue('W14', $forex_gain_1);
$objPHPExcel->getActiveSheet()->setCellValue('W15', $gain_loss_1);
$objPHPExcel->getActiveSheet()->setCellValue('W16', $remittance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('W17', $trade_finance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('W18', $processing_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('W19', $credit_card_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('W20', $insurance_fee_1);
$objPHPExcel->getActiveSheet()->setCellValue('W21', $service_charges_1);
$objPHPExcel->getActiveSheet()->setCellValue('W22', $other_commission_1);
$objPHPExcel->getActiveSheet()->setCellValue('W23', $penalty_1);
$objPHPExcel->getActiveSheet()->setCellValue('W24', $write_off_r_1);
$objPHPExcel->getActiveSheet()->setCellValue('W25', $other_income_1);
$objPHPExcel->getActiveSheet()->setCellValue('W26', $staff_cost_1);
$objPHPExcel->getActiveSheet()->setCellValue('W27', $general_administrative_1);
$objPHPExcel->getActiveSheet()->setCellValue('W28', $depreciation_1);
$objPHPExcel->getActiveSheet()->setCellValue('W29', $other_operating_1);
$objPHPExcel->getActiveSheet()->setCellValue('W30', $general_provision_1);
$objPHPExcel->getActiveSheet()->setCellValue('W31', $specific_provision_charge_1);
$objPHPExcel->getActiveSheet()->setCellValue('W32', $specific_provision_recovery_1);
$objPHPExcel->getActiveSheet()->setCellValue('W33', $write_off_c_1);
$objPHPExcel->getActiveSheet()->setCellValue('W34', $write_off_r2_1);
$objPHPExcel->getActiveSheet()->setCellValue('W35', $foreclose_pp_1);
$objPHPExcel->getActiveSheet()->setCellValue('W36', $others3_1);
$objPHPExcel->getActiveSheet()->setCellValue('W37', $corporate_tax_1);
# X
$objPHPExcel->getActiveSheet()->setCellValue('X2', '=+W2+V2');
$objPHPExcel->getActiveSheet()->setCellValue('X3', '=+W3+V3');
$objPHPExcel->getActiveSheet()->setCellValue('X4', '=+W4+V4');
$objPHPExcel->getActiveSheet()->setCellValue('X5', '=+W5+V5');
$objPHPExcel->getActiveSheet()->setCellValue('X6', '=+W6+V6');
$objPHPExcel->getActiveSheet()->setCellValue('X7', '=+W7+V7');
$objPHPExcel->getActiveSheet()->setCellValue('X8', '=+W8+V8');
$objPHPExcel->getActiveSheet()->setCellValue('X9', '=+W9+V9');
$objPHPExcel->getActiveSheet()->setCellValue('X10', '=+W10+V10');
$objPHPExcel->getActiveSheet()->setCellValue('X11', '=+W11+V11');
$objPHPExcel->getActiveSheet()->setCellValue('X12', '=+W12+V12');
$objPHPExcel->getActiveSheet()->setCellValue('X13', '=+W13+V13');
$objPHPExcel->getActiveSheet()->setCellValue('X14', '=+W14+V14');
$objPHPExcel->getActiveSheet()->setCellValue('X15', '=+W15+V15');
$objPHPExcel->getActiveSheet()->setCellValue('X16', '=+W16+V16');
$objPHPExcel->getActiveSheet()->setCellValue('X17', '=+W17+V17');
$objPHPExcel->getActiveSheet()->setCellValue('X18', '=+W18+V18');
$objPHPExcel->getActiveSheet()->setCellValue('X19', '=+W19+V19');
$objPHPExcel->getActiveSheet()->setCellValue('X20', '=+W20+V20');
$objPHPExcel->getActiveSheet()->setCellValue('X21', '=+W21+V21');
$objPHPExcel->getActiveSheet()->setCellValue('X22', '=+W22+V22');
$objPHPExcel->getActiveSheet()->setCellValue('X23', '=+W23+V23');
$objPHPExcel->getActiveSheet()->setCellValue('X24', '=+W24+V24');
$objPHPExcel->getActiveSheet()->setCellValue('X25', '=+W25+V25');
$objPHPExcel->getActiveSheet()->setCellValue('X26', '=+W26+V26');
$objPHPExcel->getActiveSheet()->setCellValue('X27', '=+W27+V27');
$objPHPExcel->getActiveSheet()->setCellValue('X28', '=+W28+V28');
$objPHPExcel->getActiveSheet()->setCellValue('X29', '=+W29+V29');
$objPHPExcel->getActiveSheet()->setCellValue('X30', '=+W30+V30');
$objPHPExcel->getActiveSheet()->setCellValue('X31', '=+W31+V31');
$objPHPExcel->getActiveSheet()->setCellValue('X32', '=+W32+V32');
$objPHPExcel->getActiveSheet()->setCellValue('X33', '=+W33+V33');
$objPHPExcel->getActiveSheet()->setCellValue('X34', '=+W34+V34');
$objPHPExcel->getActiveSheet()->setCellValue('X35', '=+W35+V35');
$objPHPExcel->getActiveSheet()->setCellValue('X36', '=+W36+V36');
$objPHPExcel->getActiveSheet()->setCellValue('X37', '=+W37+V37');

$objPHPExcel->getActiveSheet()->getStyle('H2:K37')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('N2:R37')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle('U2:X37')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

//$objPHPExcel->getActiveSheet()->getStyle('J8:K23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('E20:K22')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
//$objPHPExcel->getActiveSheet()->getStyle('E25:K29')->getNumberFormat()->setFormatCode('#,##0,,;(#,##0,,);"-"');
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('H'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('I'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('I'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('J'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('K'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('L'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('M'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('N'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('N'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('O'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('P'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('P'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('Q'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('R'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('R'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('S'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('S'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('T'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('T'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('U'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('U'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('V'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('V'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('W'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('W'.$i, 0);
    }
}
for ($i=2;$i<38;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('X'.$i)->getValue();
    if ($colB == NULL || $colB == '' ||  $colB == '-') {
        //$objPHPExcel->getActiveSheet()->setCellValue('C'.$i,  $objPHPExcel->getActiveSheet()->getCell('C'.($i-1))->getValue());
        $objPHPExcel->getActiveSheet()->setCellValue('X'.$i, 0);
    }
}
// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$file_eksport.'.xls');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
//$objWriter->save("download/NOP_".$label_tgl."_".$file_eksport.".xls");
?>
