<?php

/*######################################################################*\
#            				FUNCTION LIBRARY			                 #
# -----------------------------------------------------------------------#
# 																		 #
#  																		 #
#  	Developed by:	Asep Arifyan						    			 #
#	License:		Commercial											 #
#  	Copyright: 		2016. All Rights Reserved.		                     #
#                                                                        #
#  	Additional modules (embedded): 										 #
#	-- Metronic (Themes) 												 #
#																		 #
#																		 #
# -----------------------------------------------------------------------#
#	Designed and built with all the love and loyalty.					 #
\*######################################################################*/

function getIduser(){
    $iduser= $_SESSION['SESS_IDUSER'];
    return $iduser;
  }
function getImage(){
		$image= $_SESSION['SESS_IMAGE'];
		return $image;
	}
function getUsername(){
    $username = $_SESSION['SESS_USERNAME'];
    return $username;
  }
function getStatusAccount(){
		$status = $_SESSION['SESS_STATUS_ACCOUNT'];
		return $status;
	}
function getPassword(){
		$pass = $_SESSION['SESS_PASSWORD'];
		return $pass;
	}
function getGroupUser(){
		$group = $_SESSION['SESS_GROUP_USER'];
		return $group;
	}
function getGroupUserName(){
		global $connection;
		$query="SELECT nama_group FROM group_user  where id_group='$_SESSION[SESS_GROUP_USER]' ";
		$result=odbc_exec($connection,$query);
		$found = odbc_num_rows($result);
		if ($found >=1)
		{
			$row = odbc_fetch_array($result);
			$nama_group=$row['nama_group'];
			return $nama_group;
			}
	}

function getIp(){
		if (!empty($_SERVER['HTTP_CLIENT_IP'])) {
    		$ip = $_SERVER['HTTP_CLIENT_IP'];
		} elseif (!empty($_SERVER['HTTP_X_FORWARDED_FOR'])) {
    		$ip = $_SERVER['HTTP_X_FORWARDED_FOR'];
			} else {
   			 $ip = $_SERVER['REMOTE_ADDR'];
		}

return $ip;

	}

function getBrowser(){
		if(strpos($_SERVER['HTTP_USER_AGENT'], 'MSIE') !== FALSE)
   			$browser='Internet explorer';
 			elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Trident') !== FALSE) //For Supporting IE 11
    			$browser='Internet explorer';
				 elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Firefox') !== FALSE)
   					$browser='Mozilla Firefox';
 					elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Chrome') !== FALSE)
 					  $browser='Google Chrome';
						 elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Opera Mini') !== FALSE)
   							$browser="Opera Mini";
 								elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Opera') !== FALSE)
   									$browser="Opera";
 									elseif(strpos($_SERVER['HTTP_USER_AGENT'], 'Safari') !== FALSE)
   										$browser="Safari";
 											else
  											 $browser='Browser Lain';
return $browser;

	}

function logActivity($name,$info){
		global $connection;
		$query="insert into log_activity (username,time_activity,ip,browser,name_activity,info) values ('$_SESSION[SESS_USERNAME]',getdate(),'".getIp()."','".getBrowser()."','$name','$info')";
		$result=odbc_exec($connection,$query);

	}


function lastLogin(){
		global $connection;
		$query="select TOP 2 time_activity from log_activity where name_activity='login' order by time_activity desc";
		$result=odbc_exec($connection,$query);
		$found = odbc_num_rows($result);
		if ($found >=1)
		{
			$i=1;
			while ($row = odbc_fetch_array($result)){
			if ($i==2)
			$last_login=$row['time_activity'];
			$i++;
			}
			return $last_login;
			}

	}

function hashEncrypted($password){

		$encrypted = hash('sha256',$password);
		return $encrypted;
	}


function Milion_format($n) {
        // first strip any formatting;
        $n = (0+str_replace(",","",$n));
        
        // is this a number?
        if(!is_numeric($n)) return false;
        
       // $n=round(($n/1000000),9);
        // now filter it;
        /*
        if($n>1000000000000) return round(($n/1000000000000),1).' trillion';
        else if($n>1000000000) return round(($n/1000000000),1).' billion';
        else if($n>1000000) return round(($n/1000000),1).' million';
        else if($n>1000) return round(($n/1000),1).' thousand';
        */
        return  number_format($n,2,",",".");
        //return number_format($n);
    }

function getAccumulationMonth($parameter1,$parameter2){
		global $connection2;
 		$tgl_acc=date('Y-n-j',strtotime($parameter1));
        $bln_acc=date('n',strtotime($parameter1));
        $tot_acc=0;
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

                    //echo $query_acc;
                    //die();
                }
        } else {

                for( $i=1;$i<13;$i++){

                    $var_tgl_acc="a.Datadate='".date('Y-m-t', strtotime(date("Y-$i",strtotime($tgl_acc)). " -1 year"))."' ";

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
                }

            }

		
		return $tot_acc;
	}



function nice_number($n) {
        // first strip any formatting;
        $n = (0+str_replace(",", "", $n));

        // is this a number?
        if (!is_numeric($n)) return false;

        // now filter it;
        if ($n > 1000000000000 || $n < -1000000000000 ) {
            if ($n >0){
            return round(($n/1000000000), 3);
                } else {
            return "(".round(((-1)*$n/1000000000), 3).")";       
                }

        }
        elseif ($n > 1000000000 || $n < -1000000000) {

                if ($n >0){
           return round(($n/1000000000), 3);
                } else {
            return "(".round(((-1)*$n/1000000000), 3).")";       
                }   
        }
        elseif ($n > 1000000 || $n < -1000000 ) {
            if ($n >0){
           return round(($n/1000000), 3);
                } else {
            return "(".round(((-1)*$n/1000000), 3).")";       
                }  


            
        }
        elseif ($n > 1000 || $n < -1000) {
            if ($n >0){
           return round(($n/1000), 3);
                } else {
            return "(".round(((-1)*$n/1000), 3).")";       
                }  

            
        }
        return number_format($n);
    
    }

function numberFormat($example){
    if($example>0){
    return number_format($example, 0, ',', '.');
} else {
    return "(".number_format(-1*$example, 0, ',', '.').")";
}
}

function generateNominal($jml_record){

if($jml_record<10){
    $txt_record="00000000".$jml_record;
} else if ($jml_record<100){
    $txt_record="0000000".$jml_record;
} else if ($jml_record<1000){
    $txt_record="000000".$jml_record;
} else if($jml_record<10000){
    $txt_record="00000".$jml_record;
} else if($jml_record<100000){
    $txt_record="0000".$jml_record;
} else if($jml_record<1000000){
    $txt_record="000".$jml_record;
} else if($jml_record<10000000){
    $txt_record="00".$jml_record;
} else if($jml_record<100000000){
    $txt_record="0".$jml_record;
} else if($jml_record<1000000000){
    $txt_record=$jml_record;
} else {
    $txt_record=$jml_record;
}
return intval($txt_record);
}

function generateNominal2($jml_record){

if($jml_record<10){
    $txt_record="0000000".$jml_record;
} else if ($jml_record<100){
    $txt_record="000000".$jml_record;
} else if ($jml_record<1000){
    $txt_record="00000".$jml_record;
} else if($jml_record<10000){
    $txt_record="0000".$jml_record;
} else if($jml_record<100000){
    $txt_record="000".$jml_record;
} else if($jml_record<1000000){
    $txt_record="00".$jml_record;
} else if($jml_record<10000000){
    $txt_record="0".$jml_record;
} else if($jml_record<100000000){
    $txt_record=$jml_record;
} else if($jml_record<1000000000){
    $txt_record=$jml_record;
} else {
    $txt_record=$jml_record;
}
return $txt_record;
}




function getLengthInt($num){

$num_length = strlen((string)abs(intval($num)));
if ($num_length=='1')
{
     $txt_record="00000000".abs(intval($num));
}   else if ($num_length=='2'){
     $txt_record="0000000".abs(intval($num));
}   else if ($num_length=='3'){
     $txt_record="000000".abs(intval($num));
}   else if ($num_length=='4'){
     $txt_record="00000".abs(intval($num));
}   else if ($num_length=='5'){
     $txt_record="0000".abs(intval($num));
}   else if ($num_length=='6'){
     $txt_record="000".abs(intval($num));
}   else if ($num_length=='7'){
     $txt_record="00".abs(intval($num));
}   else if ($num_length=='8'){
     $txt_record="0".abs(intval($num));
}   else if ($num_length=='9'){
     $txt_record="".abs(intval($num));
}      else if ($num_length >= '10'){
     $txt_record=substr((string)abs($num), 0, 9);
}

 return $txt_record;
}

function getTextValue($num){
$nilai=str_replace(array('(', ')',','),'', (string)$num);
$num_length = strlen($nilai);
if ($num_length=='1')
{
     $txt_record="00000000".$nilai;
}   else if ($num_length=='2'){
     $txt_record="0000000".$nilai;
}   else if ($num_length=='3'){
     $txt_record="000000".$nilai;
}   else if ($num_length=='4'){
     $txt_record="00000".$nilai;
}   else if ($num_length=='5'){
     $txt_record="0000".$nilai;
}   else if ($num_length=='6'){
     $txt_record="000".$nilai;
}   else if ($num_length=='7'){
     $txt_record="00".$nilai;
}   else if ($num_length=='8'){
     $txt_record="0".$nilai;
}   else if ($num_length=='9'){
     $txt_record="".$nilai;
}     

 return $txt_record;
}




function getTahunBefore(){

date_default_timezone_set("Asia/Jakarta"); 
$start_date=date('Y-m-d H:i');

$array_tahun=array();

array_push($array_tahun, "<option value=''>Pilih Tahun </option>");
// Kebelakang
for ($i=4; $i>=0 ; $i--) { 
$tahun=date("Y", strtotime(date('Y-m-d H:i',strtotime($start_date))." -$i year")); //interval hari
array_push($array_tahun, "<option value='$tahun'>$tahun </option>");
}

foreach ($array_tahun as $key => $value) {
    echo  $value."<br>"; 
}



}












?>