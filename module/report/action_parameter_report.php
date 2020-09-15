<?php
session_start();
error_reporting(0);
require_once '../../data/lib/excel_reader2.php';
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

	$current_date = date('YmHis');
	$current_date_time = date('Y/m/d H:i:s');
	

//if (($_GET['module'])==sha1('66') && ($_GET['act']=='report')){
if (isset($_GET['module']) && ($_GET['act']=='report')){

$report_type=$_POST['report_type'];
$nama_file=$_FILES["nama_file"]["name"];
$file=$_FILES['nama_file']['tmp_name'];




if (isset($report_type) && $report_type != "" && isset($file) && $file!= "" )
{
switch ($report_type) {
		 case "FLASH" : $directory= "../../data/referensi/flash_report/"; break;
		 case "GL_02" : $directory= "../../data/referensi/gl02/"; break;
		 case "NII" : $directory= "../../data/referensi/nii/"; break;
		 case "ADJUSTMENT" : $directory= "../../data/referensi/adjustment/"; break;
		 case "FORECAST" : $directory= "../../data/referensi/forecast/"; break;
		 case "BS-BUDGET" : $directory= "../../data/referensi/budget/bs/"; break;
		 case "PL-BUDGET" : $directory= "../../data/referensi/budget/pl/"; break;

		 }
		 
copy ($file,$directory.$current_date.$nama_file);		 

//echo $report_type."<br>".$nama_file."<br>".$file."<br>".$directory.$nama_file;

$data = new Spreadsheet_Excel_Reader($directory.$current_date.$nama_file,false);
$jmlrow = $data->rowcount(0);		 
		 
}
//========jika ADJUSTMENT
if ($report_type=="ADJUSTMENT"){
for($i=2; $i<=$jmlrow; $i++){

		$data_bulan=trim(stripslashes($data->val($i, 1, 0)));
		$data_tahun = trim(stripslashes($data->val($i, 2, 0)));
		$data_nogl = trim(stripslashes($data->val($i, 3, 0)));
		$data_flash_level_3 = trim(stripslashes($data->val($i,4, 0)));
		$data_nominal_debet = trim(stripslashes($data->val($i,5, 0)));
		$data_nominal_kredit = trim(stripslashes($data->val($i, 6, 0)));


//echo $data_bulan."_".$data_tahun."_".$data_nogl."_".$data_flash_level_3."_".$data_nominal_debet."_".$data_nominal_kredit."<br>";
//echo "<br> jumlah baris ;".$jmlrow;
//====cek data bulan,tahun,NOGL,dan Flash
$query_cek="select * from Adjustment_Ref where BulanData='$data_bulan' and TahunData='$data_tahun' and NOGL='$data_nogl' and FLASH_LEVEL_3='$data_flash_level_3'";
//echo $query_cek;
//	die();
$result_cek=odbc_exec($connection2, $query_cek);
$found=odbc_num_rows($result_cek);
if ($found >= 1 )
{
	//delete and insert
	$query_delete=" delete Adjustment_Ref where BulanData='$data_bulan' and TahunData='$data_tahun' and NOGL='$data_nogl'";
	odbc_exec($connection2, $query_delete);
	logActivity("DELETE ADJUSTMENT","BulanData=$data_bulan and TahunData=$data_tahun  NOGL=$data_nogl");
	$query_insert =" insert into Adjustment_Ref (BulanData,TahunData,NOGL,FLASH_LEVEL_3,NominalDebet,NominalKredit) ";
	$query_insert.=" VALUES ('$data_bulan','$data_tahun','$data_nogl','$data_flash_level_3',$data_nominal_debet,$data_nominal_kredit) ";
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT ADJUSTMENT","$data_tahun,$data_nogl,$data_flash_level_3,$data_nominal_debet,$data_nominal_kredit");
	} else {
	// only insert
		
	$query_insert =" insert into Adjustment_Ref (BulanData,TahunData,NOGL,FLASH_LEVEL_3,NominalDebet,NominalKredit) ";
	$query_insert.=" VALUES ('$data_bulan','$data_tahun','$data_nogl','$data_flash_level_3',$data_nominal_debet,$data_nominal_kredit) ";
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT ADJUSTMENT","$data_tahun,$data_nogl,$data_flash_level_3,$data_nominal_debet,$data_nominal_kredit");
		}

	
}
}
//====JIKA FLASH

if ($report_type=="FLASH"){
for($i=2; $i<=$jmlrow; $i++){

		$data_flash_level_1 =trim(stripslashes($data->val($i, 1, 0)));
		$data_flash_level_1_descr = trim(stripslashes($data->val($i, 2, 0)));
		$data_flash_level_2 = trim(stripslashes($data->val($i, 3, 0)));
		$data_flash_level_2_descr = trim(stripslashes($data->val($i,4, 0)));
		$data_flash_level_3 = trim(stripslashes($data->val($i,5, 0)));
		$data_flash_level_3_descr = stripslashes($data->val($i, 6, 0));





$query_cek ="select * from Referensi_Flash_Report where FLASH_Level_1='$data_flash_level_1' and FLASH_Level_1_Description='$data_flash_level_1_descr' and ";
$query_cek.=" FLASH_Level_2 ='$data_flash_level_2' and FLASH_Level_2_Description='$data_flash_level_2_descr' and ";
$query_cek.=" FLASH_Level_3 ='$data_flash_level_3' ";

//echo $query_cek;
//	die();
$result_cek=odbc_exec($connection2, $query_cek);
$found=odbc_num_rows($result_cek);
if ($found >= 1 )
{
	//delete and insert
	$query_delete=" delete Referensi_Flash_Report where FLASH_Level_1='$data_flash_level_1' and 		FLASH_Level_1_Description='$data_flash_level_1_descr' and ";
	$query_delete.=" FLASH_Level_2 ='$data_flash_level_2' and FLASH_Level_2_Description='$data_flash_level_2_descr' and ";
	$query_delete.=" FLASH_Level_3 ='$data_flash_level_3' ";
	odbc_exec($connection2, $query_delete);
	logActivity("INSERT Referensi_Flash_Report","$data_flash_level_1,$data_flash_level_1_descr,$data_flash_level_2,FLASH_Level_2_Description,$data_flash_level_3");
	$query_insert =" insert into Referensi_Flash_Report ( FLASH_Level_1,FLASH_Level_1_Description,FLASH_Level_2,FLASH_Level_2_Description, ";
	$query_insert.=" FLASH_Level_3,FLASH_Level_3_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_1','$data_flash_level_1_descr','$data_flash_level_2','$data_flash_level_2_descr', ";
	$query_insert.=" '$data_flash_level_3','$data_flash_level_3_descr' ) ";
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_Flash_Report","$data_flash_level_1,$data_flash_level_1_descr,$data_flash_level_2,FLASH_Level_2_Description,$data_flash_level_3,$data_flash_level_3_descr");
	} else {
	// only insert
		
	$query_insert =" insert into Referensi_Flash_Report ( FLASH_Level_1,FLASH_Level_1_Description,FLASH_Level_2,FLASH_Level_2_Description, ";
	$query_insert.=" FLASH_Level_3,FLASH_Level_3_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_1','$data_flash_level_1_descr','$data_flash_level_2','$data_flash_level_2_descr', ";
	$query_insert.=" '$data_flash_level_3','$data_flash_level_3_descr' ) ";
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_Flash_Report","$data_flash_level_1,$data_flash_level_1_descr,$data_flash_level_2,FLASH_Level_2_Description,$data_flash_level_3,$data_flash_level_3_descr");
		}
	
}
}


//=======JIKA NII
if ($report_type=="NII"){
for($i=2; $i<=$jmlrow; $i++){

		$data_flash_level_3_nii =trim(stripslashes($data->val($i, 1, 0)));
		$data_flash_level_3_nii_descr =stripslashes($data->val($i, 2, 0));
		

//echo $data_flash_level_1."_".$data_flash_level_1_descr.'_'.$data_flash_level_2.'_'.$data_flash_level_2_descr.''.$data_flash_level_3.'_'.$data_flash_level_3_descr;
//die();

$query_cek =" select * from Referensi_NII where FLASH_Level_3_NII='$data_flash_level_3_nii' ";

$result_cek=odbc_exec($connection2, $query_cek);
$found=odbc_num_rows($result_cek);
if ($found >= 1 )
{
	//delete and insert
	$query_delete=" delete Referensi_NII where FLASH_Level_3_NII='$data_flash_level_3_nii' ";
	odbc_exec($connection2, $query_delete);
	logActivity("DELETE Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii");
	$query_insert =" insert into Referensi_NII ( FLASH_Level_3_NII,FLASH_Level_3_NII_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_3_nii','$data_flash_level_3_nii_descr') ";
	//echo $query_insert;
	//die ();
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii,FLASH_Level_3_NII_description=$data_flash_level_3_nii_descr");
	} else {
	// only insert
		
	$query_insert =" insert into Referensi_NII ( FLASH_Level_3_NII,FLASH_Level_3_NII_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_3_nii','$data_flash_level_3_nii_descr') ";
	//echo $query_insert;
	//die ();
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii,FLASH_Level_3_NII_description=$data_flash_level_3_nii_descr");
		}
	
}
}


//=======JIKA FORECAST
if ($report_type=="FORECAST"){



    $dataFfile = $directory.$current_date.$nama_file;
    //echo $directory.$current_date.$nama_file;
    
    $objPHPExcel = PHPExcel_IOFactory::load($dataFfile);
    $sheet = $objPHPExcel->getActiveSheet();
    $data = $sheet->rangeToArray('A2:X37');
    //echo "<br>Rows available: " . count($data) . "<br>";
    foreach ($data as $row=>$value) {
		$query_insert="";
        //echo  $value['0'].", ".$value['1'].", ".$value['2'].", ".$value['3'].", ".$value['4'].", ".$value['5'].", ".$value['6'].", ".$value['7'].", ".$value['8'].", ".$value['9'].", ".$value['10'].", ".$value['11'].", ".$value['12'].", ".$value['13'].$value['14'].", ".$value['15'].", ".$value['16'].", ".$value['17'].", ".$value['18'].", ".$value['19'].", ".$value['20']."<br>";
        $date1=date('Y-m-d',strtotime($value['2']));

        $date2=date('Y-m-d',strtotime($value['4']));
       
        $value['7']=str_replace(array(',',')'),"",$value['7']);
        $value['7']=str_replace("(","-",$value['7']);
        $value['8']=str_replace(array(',',')'),"",$value['8']);
        $value['8']=str_replace("(","-",$value['8']);
        $value['9']=str_replace(array(',',')'),"",$value['9']);
        $value['9']=str_replace("(","-",$value['9']);
        $value['10']=str_replace(array(',',')'),"",$value['10']);
        $value['10']=str_replace("(","-",$value['10']);
        $value['11']=str_replace(array(',',')'),"",$value['11']);
        $value['11']=str_replace("(","-",$value['11']);
        $value['12']=str_replace(array(',',')'),"",$value['12']);
        $value['12']=str_replace("(","-",$value['12']);
        $value['13']=str_replace(array(',',')'),"",$value['13']);
        $value['13']=str_replace("(","-",$value['13']);
        $value['14']=str_replace(array(',',')'),"",$value['14']);
        $value['14']=str_replace("(","-",$value['14']);
        $value['15']=str_replace(array(',',')'),"",$value['15']);
        $value['15']=str_replace("(","-",$value['15']);
        $value['16']=str_replace(array(',',')'),"",$value['16']);
        $value['16']=str_replace("(","-",$value['16']);
        $value['17']=str_replace(array(',',')'),"",$value['17']);
        $value['17']=str_replace("(","-",$value['17']);
        $value['18']=str_replace(array(',',')'),"",$value['18']);
        $value['18']=str_replace("(","-",$value['18']);
        $value['19']=str_replace(array(',',')'),"",$value['19']);
        $value['19']=str_replace("(","-",$value['19']);
        $value['20']=str_replace(array(',',')'),"",$value['20']);
        $value['20']=str_replace("(","-",$value['20']);
        $value['21']=str_replace(array(',',')'),"",$value['21']);
        $value['21']=str_replace("(","-",$value['21']);
        $value['22']=str_replace(array(',',')'),"",$value['22']);
        $value['22']=str_replace("(","-",$value['22']);
        $value['23']=str_replace(array(',',')'),"",$value['23']);
        $value['23']=str_replace("(","-",$value['23']);
       //-----


        //-----
        $query_cek =" select * from master_forecast2 where  Data_Date='$date1' and FLASH_LEVEL_3='$value[0]' ";
        $result_cek=odbc_exec($connection2, $query_cek);
		$found=odbc_num_rows($result_cek);
		if ($found >=1){
			$query_delete=" delete master_forecast2 where Data_Date='$date1' and FLASH_LEVEL_3='$value[0]' ";
			odbc_exec($connection2, $query_delete);
		}
        $query_insert =" insert into master_forecast2 (FLASH_LEVEL_3,Description,Data_Date,Number_Of_Days_Actual,End_Date,Number_Of_Days_End,Rest_Of_Days, ";
        $query_insert.=" Last_Actual,Actual_YTD,Actual_MTD,Average,Interest,LDR,Loan,Asumsi_Loan,DPK,AddLoan,Add_Interest,Asumsi_EIR,Employee_Loan_Benefit,Add_forecast,";
        $query_insert.=" Proyeksi_MTD,YTD_Actual_Last_Month, Proyeksi_YTD ) ";
        $query_insert.=" values ('".$value['0']."','$value[1]','$date1',$value[3],'$date2',$value[5],$value[6],$value[7],$value[8],$value[9],$value[10], ";
        $query_insert.=" $value[11],$value[12],$value[13],$value[14],$value[15],$value[16],$value[17],$value[18],$value[19],$value[20], ";
        $query_insert.=" $value[21],$value[22],$value[23]"." ) ";
        //echo "<br>".$query_insert."<br>";
      	//die();
      	odbc_exec($connection2, $query_insert);

      	//echo "<br>".$query_insert."<br>";
      	//die();

}

//die();
}



############### budget #####################



if ($report_type=="BS-BUDGET"){



 $dataFfile = $directory.$current_date.$nama_file;
    //echo $directory.$current_date.$nama_file;
    
    $objPHPExcel = PHPExcel_IOFactory::load($dataFfile);
    $sheet = $objPHPExcel->getActiveSheet();
    $data = $sheet->rangeToArray("A2:X$jmlrow");
    //echo "<br>Rows available: " . count($data) . "<br>";
     //$objPHPExcel = PHPExcel_IOFactory::load("download/Flash_Report_".$label_tgl."_".$file_eksport.".xls");
 //$objWorksheet = $objPHPExcel->getActiveSheet('2');

    $i=2;
    foreach ($data as $row=>$value) {
		
		$year_budget=substr($value['2'],0,4);
		$mon_budget=substr($value['2'],5,2);

        $value['3']=str_replace(array(',',')'),"",$value['3']);
        $value['3']=str_replace("(","-",$value['3']);
        //echo $i.") ".$value['0']." ".$value['1']." ".$value['2']." ".$objPHPExcel->getActiveSheet()->getCell("D$i")."<br>" ;
       
        $budget=$objPHPExcel->getActiveSheet()->getCell("D$i");

        $query_cek=" select * from Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$value[0]' ";
   
        $result_cek=odbc_exec($connection2, $query_cek);
		$found=odbc_num_rows($result_cek);
		if ($found >=1){
			$query_delete=" delete Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$value[0]' ";
			odbc_exec($connection2, $query_delete);
			$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget) ";
			$query_insert.=" VALUES ('$value[0]','$value[1]','$value[2]','$budget') ";
			odbc_exec($connection2, $query_insert);
		} else {
			$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget) ";
			$query_insert.=" VALUES ('$value[0]','$value[1]','$value[2]','$budget') ";
			odbc_exec($connection2, $query_insert);
		}

 $i++;

}

}


if ($report_type=="PL-BUDGET"){
//echo "test";
//die();


 $dataFfile = $directory.$current_date.$nama_file;
    //echo $directory.$current_date.$nama_file;
    
    $objPHPExcel = PHPExcel_IOFactory::load($dataFfile);
    $sheet = $objPHPExcel->getActiveSheet();
    $data = $sheet->rangeToArray("A2:X$jmlrow");
    //echo "<br>Rows available: " . count($data) . "<br>";
     //$objPHPExcel = PHPExcel_IOFactory::load("download/Flash_Report_".$label_tgl."_".$file_eksport.".xls");
 //$objWorksheet = $objPHPExcel->getActiveSheet('2');

    $i=2;
    foreach ($data as $row=>$value) {
		
		$year_budget=substr($value['2'],0,4);
		$mon_budget=substr($value['2'],5,2);

        $value['3']=str_replace(array(',',')'),"",$value['3']);
        $value['3']=str_replace("(","-",$value['3']);
    //echo $i.") ".$value['0']." ".$value['1']." ".$value['2']." mtd :".$objPHPExcel->getActiveSheet()->getCell("D$i")." ytd :".$objPHPExcel->getActiveSheet()->getCell("E$i")."<br>";
       
        $budget_mtd=$objPHPExcel->getActiveSheet()->getCell("D$i");
        $budget_ytd=$objPHPExcel->getActiveSheet()->getCell("E$i");

        $query_cek=" select * from Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$value[0]' ";
   
        $result_cek=odbc_exec($connection2, $query_cek);
		$found=odbc_num_rows($result_cek);
		if ($found >=1){
			$query_delete=" delete Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$value[0]' ";
			odbc_exec($connection2, $query_delete);
			$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget_MTD,Budget_YTD) ";
			$query_insert.=" VALUES ('$value[0]','$value[1]','$value[2]','$budget') ";
			odbc_exec($connection2, $query_insert);
		} else {
			$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget_MTD,Budget_YTD) ";
			$query_insert.=" VALUES ('$value[0]','$value[1]','$value[2]','$budget_mtd','$budget_ytd') ";
			odbc_exec($connection2, $query_insert);
		}

 $i++;

}

}

/*

for($i=2; $i<=$jmlrow; $i++){

		$flash_3=trim(stripslashes($data->val($i, 1, 0)));
		$flash_descr =stripslashes($data->val($i, 2, 0));
		$datadate = trim(stripslashes($data->val($i, 3, 0)));
		$budget = trim(stripslashes($data->val($i, 4, 0)));
		$year_budget=substr("$datadate",0,4);
		$mon_budget=substr("$datadate",5,2);

//echo $budget."<br>";


$query_cek=" select * from Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$flash_3' ";

$result_cek=odbc_exec($connection2, $query_cek);
$found=odbc_num_rows($result_cek);
if ($found >= 1 )
{
	//delete and insert
	$query_delete=" delete Budget_BS where datepart(month,datadate) ='$mon_budget' and datepart(year,datadate) = '$year_budget'  and FLASH_LEVEL_3='$flash_3' ";
	odbc_exec($connection2, $query_delete);
	//logActivity("DELETE ADJUSTMENT","BulanData=$data_bulan and TahunData=$data_tahun  NOGL=$data_nogl");
	$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget) ";
	$query_insert.=" VALUES ('$flash_3','$flash_descr','$datadate','$budget') ";
	odbc_exec($connection2, $query_insert);
	//logActivity("INSERT ADJUSTMENT","$data_tahun,$data_nogl,$data_flash_level_3,$data_nominal_debet,$data_nominal_kredit");
	} else {
	// only insert
		
	$query_insert =" insert into Budget_BS (FLASH_Level_3,FLASH_Level_3_Description,DataDate,Budget) ";
	$query_insert.=" VALUES ('$flash_3','$flash_descr','$datadate','$budget') ";
	odbc_exec($connection2, $query_insert);
	//logActivity("INSERT ADJUSTMENT","$data_tahun,$data_nogl,$data_flash_level_3,$data_nominal_debet,$data_nominal_kredit");
		}


}
*/	



//=======JIKA GL_02

//GLNO,PRODNO,GLNAME,PM_COA_Level_4,NOP_Level_3,SBDK_Level_3,Lap_Keu_Level_3,LAPOSIM_LEVEL_2,SBELPS_Level_2,CRLPS_Level_2,CRLPS_Level_21,FLASH_LEVEL_3,UMKM_LEVEL_3,LTV_LEVEL_3,PBLKS_LEVEL_3,LKBLPS_LEVEL_3,KPMM_Level_3,RASIO_Level_3,[FLASH_LEVEL_3 _NII
if ($report_type=="GL_02"){
for($i=2; $i<=$jmlrow; $i++){

		$data_flash_level_3_nii =trim(stripslashes($data->val($i, 1, 0)));
		$data_flash_level_3_nii_descr =stripslashes($data->val($i, 2, 0));
		

//echo $data_flash_level_1."_".$data_flash_level_1_descr.'_'.$data_flash_level_2.'_'.$data_flash_level_2_descr.''.$data_flash_level_3.'_'.$data_flash_level_3_descr;
//die();

$query_cek =" select * from Referensi_NII where FLASH_Level_3_NII='$data_flash_level_3_nii' ";

$result_cek=odbc_exec($connection2, $query_cek);
$found=odbc_num_rows($result_cek);
if ($found >= 1 )
{
	//delete and insert
	$query_delete=" delete Referensi_NII where FLASH_Level_3_NII='$data_flash_level_3_nii' ";
	odbc_exec($connection2, $query_delete);
	logActivity("DELETE Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii");
	$query_insert =" insert into Referensi_NII ( FLASH_Level_3_NII,FLASH_Level_3_NII_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_3_nii','$data_flash_level_3_nii_descr') ";
	//echo $query_insert;
	//die ();
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii,FLASH_Level_3_NII_description=$data_flash_level_3_nii_descr");
	} else {
	// only insert
		
	$query_insert =" insert into Referensi_NII ( FLASH_Level_3_NII,FLASH_Level_3_NII_Description ) ";
	$query_insert.=" VALUES ('$data_flash_level_3_nii','$data_flash_level_3_nii_descr') ";
	//echo $query_insert;
	//die ();
	odbc_exec($connection2, $query_insert);
	logActivity("INSERT Referensi_NII","FLASH_Level_3_NII=$data_flash_level_3_nii,FLASH_Level_3_NII_description=$data_flash_level_3_nii_descr");
		}
	
}
}















//logActivity("test","keterangan----");

header("location: ../../index.php?module=$_GET[module]&type=$report_type&message=success");
//}
}
