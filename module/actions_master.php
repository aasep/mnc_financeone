<?php
	require_once '../config/config.php';
	require_once '../function/function.php';
	require_once 'session_login.php';
    //require_once '../session_group.php';


//==============add user=========
if (($_GET['module'])==sha1('2') && ($_GET['act']=='add_user')){
$username=strtolower($_POST['username']);
$password=hashEncrypted($_POST['password']);
$status_account=$_POST['status_account'];
$group_user=$_POST['group_user'];

$query="insert into user_account (username,password,status_account,id_group,addby) values ('$username','$password','$status_account','$group_user','admin')";
$result=odbc_exec($connection, $query);

	if ($result)
  		{
  			logActivity("add_user","username=$username;status_account=$status_account,group_user=$group_user,addby=$username");
  			header("location: ../index?module=$_GET[module]&message=success");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error");

  		}
//header("location: index.php");
}
//==============edit user=========
if (($_GET['module'])==sha1('2') && ($_GET['act']=='edit_user')){

$username=strtolower($_POST['username']);

$password=hashEncrypted($_POST['password']);
if (isset($_POST['password']) && $_POST['password'] !="" ){

	$var_password=" password='$password', ";

} else {
	$var_password="";

}

$status_account=$_POST['status_account'];
$group_user=$_POST['group_user'];

$query="update user_account set  $var_password status_account='$status_account', id_group='$group_user' where  username='$username'";
//echo $query;
//die();

$result=odbc_exec($connection, $query);
$found=odbc_num_rows($result);
	if ($found >=1 )
  		{
  			logActivity("edit_user","username=$username;status_account=$status_account,group_user=$group_user");
  			header("location: ../index?module=$_GET[module]&message=success3");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error3");

  		}


}
//==============delete user=========
if (($_GET['module'])==sha1('2') && ($_GET['act']=='delete_user')){

$id_user=strtolower($_POST['id_user']);
$query="delete user_account where username='$id_user'";
$result=odbc_exec($connection, $query);

if ($result)
  		{ 
  			logActivity("delete_user","username=$id_user");
  			header("location: ../index?module=$_GET[module]&message=success2");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error2");

  		}



}
//==============add group user=========
if (($_GET['module'])==sha1('3') && ($_GET['act']=='add_group_user')){
$nama_group=$_POST['nama_group'];
$inisial=$_POST['inisial'];

$query="insert into group_user (nama_group,inisial) values ('$nama_group','$inisial')";
$result=odbc_exec($connection, $query);

	if ($result)
  		{
  			logActivity("add_group_user","nama_group=$nama_group,inisial=$inisial");
  			header("location: ../index?module=$_GET[module]&message=success");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error");

  		}
}
//==============edit group menu=============
if (($_GET['module'])==sha1('3') && ($_GET['act']=='edit_group_user')){
$id_group=$_POST['id_group'];
$nama_group=$_POST['ed_nama_group'];
$inisial=$_POST['ed_inisial'];

$query="update group_user set nama_group='$nama_group',inisial='$inisial' where id_group='$id_group'";
$result=odbc_exec($connection, $query);

	if ($result)
  		{
  			logActivity("edit_group_user","nama_group=$nama_group,inisial=$inisial,id_group=$id_group");
  			header("location: ../index?module=$_GET[module]&message=success3");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error3");
  		}



}
//==============delete group menu=============
if (($_GET['module'])==sha1('3') && ($_GET['act']=='delete_group_user')){


$id_group=$_POST['id_group'];
$query="delete group_user where id_group='$id_group'";
$result=odbc_exec($connection, $query);

if ($result)
  		{
  			logActivity("delete_group_user","id_group=$id_group");
  			header("location: ../index?module=$_GET[module]&message=success2");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error2");

  		}



}
//==============add menu=============
if (($_GET['module'])==sha1('4') && ($_GET['act']=='add_menu')){

$nama_menu=$_POST['nama_menu'];
$parent=$_POST['parent'];

$query="insert into menu (nama_menu,parent) values ('$nama_menu','$parent')";
$result=odbc_exec($connection, $query);

$query2="select * from menu where nama_menu='$nama_menu' ";
$result2=odbc_exec($connection, $query2);
while ($row2 = odbc_fetch_array($result2))

{
$id_menu=$row2['id_menu'];
$id_menu_encrypt=sha1($row2['id_menu']);
$parent=$row2['parent'];
}

if ($parent != 0) {
$query3="update menu set src='$id_menu_encrypt' where id_menu='$id_menu' ";
$result3=odbc_exec($connection, $query3);
}

//echo $query2."<br>". $query3."<br>".$id_menu;
//die();
	if ($result)
  		{
  			logActivity("add_menu","nama_menu=$nama_menu,parent=$parent");
  			header("location: ../index?module=$_GET[module]&message=success");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error");

  		}


}
//==========edit menu=================
if (($_GET['module'])==sha1('4') && ($_GET['act']=='edit_menu')){


$id_menu=$_POST['id_menu'];
$parent=$_POST['parent'];
$nama_menu=$_POST['nama_menu'];


$query="update menu set   nama_menu='$nama_menu',parent='$parent' where  id_menu='$id_menu'";
$result=odbc_exec($connection, $query);

	if ($result)
  		{
  			logActivity("edit_menu","nama_menu=$nama_menu,parent=$parent,id_menu=$id_menu");
  			header("location: ../index?module=$_GET[module]&message=success3");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error3");

  		}


}
//==========delete menu=================
if (($_GET['module'])==sha1('4') && ($_GET['act']=='delete_menu')){
$id_menu=$_POST['id_menu'];
$query="delete menu where id_menu='$id_menu'";
$result=odbc_exec($connection, $query);

if ($result)
  		{	logActivity("delete_menu","id_menu=$id_menu");
  			header("location: ../index?module=$_GET[module]&message=success2");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error2");

  		}


}
//==========Ganti Password=================
if (($_GET['module'])==sha1('64') && ($_GET['act']=='change-pass')){
	
$username=$_POST['username'];
$password=sha1($_POST['password']);


$query="update user_account set   password='$password' where  username='$username'";
$result=odbc_exec($connection, $query);

	if ($result)
  		{
  			logActivity("change-pass","username=$username");
  			header("location: ../index?module=$_GET[module]&message=success");
  	} else  {
  			header("location: ../index?module=$_GET[module]&message=error");

  		}


}
//==========delete menu=================
if (($_GET['module'])==sha1('5') && ($_GET['act']=='edit_gmenu')){

}


//==========upload/add modal=================
if (($_GET['module'])==sha1('68') && ($_GET['act']=='upload-modal')){

$tanggal=date('Y-m-d',strtotime($_POST['tanggal']));
$nilai_modal=$_POST['nilai_modal'];


$q_cek =" select * from master_modal where DataDate='$tanggal' ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);

if ($q_found >=1 )
{
$query_delete=" delete master_modal where DataDate='$tanggal' ";
$result_delete=odbc_exec($connection2, $query_delete);
} 


$query="insert into master_modal (DataDate,Nominal_Modal) values ('$tanggal','$nilai_modal')";

$result=odbc_exec($connection2, $query);

if ($result)
      {
        logActivity("add master_modal","data_date=$tanggal,nominal_modal=nilai_modal");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}

//==========delete modal=================
if (($_GET['module'])==sha1('68') && ($_GET['act']=='delete_modal')){

$tanggal=date('Y-m-d',strtotime($_POST['tgl']));
$nilai_modal=$_POST['nominalmodal'];


$query=" delete master_modal where DataDate='$tanggal' and nominal_modal='$nilai_modal' ";


//echo $query;
//die();
$result=odbc_exec($connection2, $query);

//die();

if ($result)
      {
        logActivity("delete_modal","data_date=$tanggal,nominal_modal=nilai_modal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}

//==========edit modal=================
if (($_GET['module'])==sha1('68') && ($_GET['act']=='delete_modal')){

$tanggal=date('Y-m-d',strtotime($_POST['ed_tanggal']));
$nilai_modal=$_POST['ed_nilai_modal'];


$query=" delete master_modal where DataDate='$tanggal' and nominal_modal='$nilai_modal' ";


//echo $query;
//die();
$result=odbc_exec($connection2, $query);

//die();

if ($result)
      {
        logActivity("delete_modal","data_date=$tanggal,nominal_modal=nilai_modal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}

//==========upload modal=================
if (($_GET['module'])==sha1('69') && ($_GET['act']=='upload-forecast')){


//echo "testing <br>";
echo $_POST['nilai_interest']."<br>";
echo $_POST['nilai_ldr']."<br>";
}



//==========upload/add Pajak=================
if (($_GET['module'])==sha1('72') && ($_GET['act']=='upload-pajak')){

$tanggal=date('Y-m-d',strtotime($_POST['tanggal']));
$nilai_pajak=$_POST['nilai_pajak'];

$mon_pajak=date('n',strtotime($tanggal));
$year_pajak=date('Y',strtotime($tanggal));
//Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal'

//echo "ok";
//die();

$q_cek =" select * from master_pajak where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);

if ($q_found >=1 )
{
$query_delete=" delete master_pajak where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' ";
$result_delete=odbc_exec($connection2, $query_delete);
} 


$query="insert into master_pajak (DataDate,Nominal_Pajak) values ('$tanggal','$nilai_pajak')";

$result=odbc_exec($connection2, $query);

if ($result)
      {
        logActivity("add master_pajak","data_date=$tanggal,nominal_pajak=$nilai_pajak");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}

//========
//==========Delete Pajak=================
if (($_GET['module'])==sha1('72') && ($_GET['act']=='delete_pajak')){

$tanggal=date('Y-m-d',strtotime($_POST['tgl2']));
$nilai_pajak=$_POST['nominal_pajak'];

$mon_pajak=date('n',strtotime($tanggal));
$year_pajak=date('Y',strtotime($tanggal));
//Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal'

$query_delete=" delete master_pajak where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' ";
$result_delete=odbc_exec($connection2, $query_delete);


if ($result_delete)
      {
        logActivity("delete master_pajak","data_date=$tanggal,nominal_pajak=$nilai_pajak");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}



//========== input master_counter_rate =================
if (($_GET['module'])==sha1('84') && ($_GET['act']=='alco')){

$tanggal=date('Y-m-d',strtotime($_POST['tanggal']));
$mata_uang=$_POST['mata_uang'];

$min_alco_1=$_POST['min_alco_1'];
$max_alco_1=$_POST['max_alco_1'];
$min_alco_3=$_POST['min_alco_3'];
$max_alco_3=$_POST['max_alco_3'];

$mon_pajak=date('n',strtotime($tanggal));
$year_pajak=date('Y',strtotime($tanggal));
//Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal'

$q_cek =" select * from master_counter_rate where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' and JenisMataUang='$mata_uang' ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);

if ($q_found >=1 )
{
$query_delete=" delete master_counter_rate where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' and JenisMataUang='$mata_uang' ";
$result_delete=odbc_exec($connection2, $query_delete);
} 


$query="insert into master_counter_rate (DataDate,Min_Rate1,Max_Rate1,Min_Rate3,Max_Rate3,JenisMataUang) values ('$tanggal','$min_alco_1','$max_alco_1','$min_alco_3','$max_alco_3','$mata_uang')";

$result=odbc_exec($connection2, $query);

if ($result)
      {
        logActivity("add master_counter_rate","data_date=$tanggal,Min_Rate1=$min_alco_1,Max_Rate1=$max_alco_1,Min_Rate3=$min_alco_3,Max_Rate3=$max_alco_3");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}



//==========delete master_counter_rate=================
if (($_GET['module'])==sha1('84') && ($_GET['act']=='delete_alco')){

$tanggal=date('Y-m-d',strtotime($_POST['tgl']));
$matauang=$_POST['matauang'];

$mon_pajak=date('n',strtotime($tanggal));
$year_pajak=date('Y',strtotime($tanggal));
//Month(DataDate)='$mon_modal' and Year(DataDate)='$year_modal'

$query_delete =" delete from master_counter_rate where Month(DataDate)='$mon_pajak' and Year(DataDate)='$year_pajak' and JenisMataUang='$matauang' ";
$result_delete=odbc_exec($connection2, $query_delete);

if ($result_delete)
      {
        logActivity("delete master_counter_rate","data_date=$tanggal,matauang=$matauang");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}
//delete_pajak


//========== Upload Suku bunga efektif =======================
if (($_GET['module'])==sha1('70') && ($_GET['act']=='skb_efektif')){

$mata_uang=$_POST['mata_uang'];

$from =date('Y-m-d',strtotime($_POST['from']));
$to=date('Y-m-d',strtotime($_POST['to']));
$range_atas=$_POST['range_atas'];
$range_bawah=$_POST['range_bawah'];
$range_valas=$_POST['range_valas'];

$mon=date('n',strtotime($from));
$year=date('Y',strtotime($from));



$q_cek =" select * from Master_SKB_Efektif where Month(DataDate)='$mon' and Year(DataDate)='$year'  ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);
$row_found=odbc_fetch_array($q_cek_result);


 



if ($q_found >=1 )
{
  
   if ($mata_uang=='1')
    {
        //update Range Atas dan Bawah
        $query =" update Master_SKB_Efektif set Range_Atas='$range_atas',Range_Bawah='$range_bawah' WHERE Month(DataDate)='$mon' and Year(DataDate)='$year' ";
        $result=odbc_exec($connection2, $query);

      } else {
        $query =" update Master_SKB_Efektif set Range_Valas='$range_valas' WHERE Month(DataDate)='$mon' and Year(DataDate)='$year' ";
        $result=odbc_exec($connection2, $query);
        //update Range Valas

      }
      
} else { 

if ($mata_uang=='1')
    {
        //insert  Range Atas dan Bawah----> range valas==0
      $query =" insert into Master_SKB_Efektif (DataDate, Range_Atas, Range_Bawah,Range_Valas) values ('$from','$range_atas','$range_bawah','0') ";
      $result=odbc_exec($connection2, $query);

      } else {

        //insert range valas ---> range_atas=0 , range_bawah=0
      $query =" insert into Master_SKB_Efektif (DataDate, Range_Atas, Range_Bawah,Range_Valas) values ('$from','0','0','$range_valas') ";
      $result=odbc_exec($connection2, $query);

      }




}

if ($result)
      {
        logActivity("Add Master Suku Bunga Efektif","from=$from,to=$to, range_atas=$range_atas, range_bawah=$range_bawah");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}

//==========delete suku bunga efektif=================
if (($_GET['module'])==sha1('70') && ($_GET['act']=='delete_skb_efektif')){

$tanggal=date('Y-m-d',strtotime($_POST['tgl']));


$mon=date('n',strtotime($tanggal));
$year=date('Y',strtotime($tanggal));

$query_delete =" delete Master_SKB_Efektif where  Month(DataDate)='$mon' and Year(DataDate)='$year' ";
$result_delete=odbc_exec($connection2, $query_delete);


if ($result_delete)
      {
        logActivity("delete Master Suku Bunga Efektif ","data_date=$tanggal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}

//========== Input Premi LPS  =======================
if (($_GET['module'])==sha1('85') && ($_GET['act']=='input_premi')){

$semester=$_POST['semester'];
$tahun=$_POST['tahun'];

//$from =date('Y-m-d',strtotime($_POST['from']));
//$to=date('Y-m-d',strtotime($_POST['to']));
switch ($semester) {
        case '1':
        $name="Semester 1";
        $tanggal_akhir=$tahun."-06-30";
        $tanggal_start=$tahun."-01-01";
        $input=$tahun."-06-01";
        $array_bulan=array("Januari","Februari","Maret","April","Mei","Juni");
        //$array_bulan_sebelumnya=array("Juli","Agustus","September","Oktober","November","Desember");
        //$tahun_sebelumnya=date('Y',strtotime(date('Y-m-d',strtotime($input))." -1 year"));
        break;
        case '2':
        $name="Semester 2";
        $tanggal_start=$tahun."-07-31";
        $tanggal_akhir=$tahun."-12-31";
        $input=$tahun."-12-01";
        $array_bulan=array("Juli","Agustus","September","Oktober","November","Desember");
        //$array_bulan_sebelumnya=array("Januari","Februari","Maret","April","Mei","Juni");
        //$tahun_sebelumnya=date('Y',strtotime(date('Y-m-d',strtotime($input))." 0 year"));
        break;
        
     
}
        

$premi_ver=$_POST['premi_ver'];
$premi_sebelumnya=$_POST['premi_sebelumnya'];

$mon1=date('n',strtotime($tanggal_start));
$year1=date('Y',strtotime($tanggal_start));
$mon2=date('n',strtotime($tanggal_akhir));
$year2=date('Y',strtotime($tanggal_akhir));

$q_cek =" select * from Master_Saldo_Premi_LPS where Month(Periode_Awal)='$mon1' and Year(Periode_Awal)='$year1' and Month(Periode_Akhir)='$mon2' and Year(Periode_Akhir)='$year2'  ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);
$row_found=odbc_fetch_array($q_cek_result);

if ($q_found >=1 )
{

        $query =" update Master_Saldo_Premi_LPS set Jumlah_Premi_Verifikasi_LPS='$premi_ver',Saldo_Premi_Periode_Lalu='$premi_sebelumnya' ";
        $query.=" where Month(Periode_Awal)='$mon1' and Year(Periode_Awal)='$year1' and Month(Periode_Akhir)='$mon2' and Year(Periode_Akhir)='$year2' ";
        $result=odbc_exec($connection2, $query);

} else { 

      $query =" insert into Master_Saldo_Premi_LPS (Periode_Awal, Periode_Akhir, Jumlah_Premi_Verifikasi_LPS,Saldo_Premi_Periode_Lalu) ";
      $query.=" values ('$tanggal_start','$tanggal_akhir','$premi_ver','$premi_sebelumnya')  ";
      $result=odbc_exec($connection2, $query);

   
}

if ($result)
      {
        logActivity("Input Premi LPS","tgl_start=$tanggal_start,tgl_akhir=$tanggal_akhir, ver_premi=$premi_ver, saldo_premi_sblmny=$premi_sebelumnya");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}


//==========delete suku bunga efektif=================
if (($_GET['module'])==sha1('85') && ($_GET['act']=='delete_premi')){

$mon1=date('n',strtotime($_POST['awal']));
$year1=date('Y',strtotime($_POST['awal']));
$mon2=date('n',strtotime($_POST['akhir']));
$year2=date('Y',strtotime($_POST['akhir']));


$query_delete =" delete Master_Saldo_Premi_LPS where Month(Periode_Awal)='$mon1' and Year(Periode_Awal)='$year1' and Month(Periode_Akhir)='$mon2' and Year(Periode_Akhir)='$year2' ";
$result_delete=odbc_exec($connection2, $query_delete);


if ($result_delete)
      {
        logActivity("delete Master Suku Bunga Efektif ","data_date=$tanggal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}


//========== Weight Average ======================================================================================================================
if (($_GET['module'])==sha1('86') && ($_GET['act']=='weight_average')){

$xtahun=$_POST['tahun'];
$xbulan=$_POST['bulan'];
$nilai_we=$_POST['weight_average'];


$tmp_datadate=date('Y-m-t',strtotime($xtahun."-".$xbulan."-01"));


//echo $tmp_datadate;
//die();
$mon=date('n',strtotime($tmp_datadate));
$year=date('Y',strtotime($tmp_datadate));

$q_cek =" select * from weighted_average where Month(DataDate)='$mon' and Year(DataDate)='$year'  ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);
//$row_found=odbc_fetch_array($q_cek_result);

if ($q_found >=1 )
{

        $query =" update weighted_average set nilai='$nilai_we'";
        $query.=" where Month(DataDate)='$mon' and Year(DataDate)='$year' ";
        $result=odbc_exec($connection2, $query);

} else { 

      $query =" insert into weighted_average (DataDate, nilai) ";
      $query.=" values ('$tmp_datadate','$nilai_we')  ";
      $result=odbc_exec($connection2, $query);

   
}


if ($result)
      {
        logActivity("Weight Average","DataDate=$tmp_datadate , nilai=$nilai_we");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}

//==========delete weight average=================
if (($_GET['module'])==sha1('86') && ($_GET['act']=='delete_weight_average')){

$tanggal=$_POST['tgl'];
$nilai=$_POST['nilai'];

$query_delete =" delete weighted_average where DataDate='$tanggal' AND nilai='$nilai' ";
$result_delete=odbc_exec($connection2, $query_delete);


if ($result_delete)
      {
        logActivity("delete weighted_average ","data_date=$tanggal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}


//========== Input ATMR ======================================================================================================================
if (($_GET['module'])==sha1('87') && ($_GET['act']=='input_atmr')){

$tanggal=date('Y-m-d',strtotime($_POST['tanggal']));
$atmr_kredit=$_POST['atmr_kredit'];
$atmr_pasar=$_POST['atmr_pasar'];
$atmr_operasional=$_POST['atmr_operasional'];

//echo $tmp_datadate;
//die();
$mon=date('n',strtotime($tanggal));
$year=date('Y',strtotime($tanggal));

$q_cek =" select * from Master_ATMR where Month(DataDate)='$mon' and Year(DataDate)='$year'  ";
$q_cek_result=odbc_exec($connection2, $q_cek);
$q_found=odbc_num_rows($q_cek_result);
//$row_found=odbc_fetch_array($q_cek_result);

if ($q_found >=1 )
{
        $query =" update Master_ATMR set atmr_kredit='$atmr_kredit', atmr_pasar='$atmr_pasar', atmr_operasional='$atmr_operasional', DataDate='$tanggal' ";
        $query.=" where Month(DataDate)='$mon' and Year(DataDate)='$year' ";
        $result=odbc_exec($connection2, $query);

} else { 

      $query =" insert into Master_ATMR (DataDate, atmr_kredit, atmr_pasar, atmr_operasional) ";
      $query.=" values ('$tanggal','$atmr_kredit', '$atmr_pasar', '$atmr_operasional')  ";
      $result=odbc_exec($connection2, $query);
   
}


if ($result)
      {
        logActivity("Master_ATMR","DataDate=$tmp_datadate , $atmr_kredit, $atmr_pasar, $atmr_operasional ");
        header("location: ../index?module=$_GET[module]&message=success");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error");

      }
}

// delete_atmr
//========== delete_atmr =================
if (($_GET['module'])==sha1('87') && ($_GET['act']=='delete_atmr')){

$tanggal=date('Y-m-d',strtotime($_POST['tgl']));

$query_delete =" delete Master_ATMR where DataDate='$tanggal'  ";
$result_delete=odbc_exec($connection2, $query_delete);


if ($result_delete)
      {
        logActivity("delete Master_ATMR ","data_date=$tanggal");
        header("location: ../index?module=$_GET[module]&message=success2");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error2");

      }
}



//========== Edit Profile  =================
if (($_GET['module'])==sha1('89') && ($_GET['act']=='edit_profile')){

$id_user=$_POST['id_user'];

$nama_lengkap=$_POST['nama_lengkap'];

//var_dump($_FILES['file_img']);
//die();

if( $_FILES['file_img']['name']!="" ) {

  $image_temp=$_FILES['file_img']['tmp_name'];
  $nama=$_FILES['file_img']['name'];
  $type=$_FILES['file_img']['type'];
  $ext = pathinfo($nama, PATHINFO_EXTENSION);
//echo $ext;
  $directory="../images/profile/".$id_user.".$ext";
  copy($image_temp,$directory);
  
  $var_image=" image='".$id_user.".$ext' ,";
  $_SESSION['SESS_IMAGE'] = $id_user.".$ext";
  } else {
  $var_image=" "; 
  }

$query=" update user_account set $var_image username='$nama_lengkap'  where id_user='$id_user' ";
//echo $query;
//die();
$result=odbc_exec($connection, $query);

  if ($result)
      { 
        $_SESSION['SESS_USERNAME']=$nama_lengkap;
        //$_SESSION['SESS_IMAGE'] = $fix_image;
        //$_SESSION['SESS_IDUSER'] = $id_user;

        logActivity("Edit PROFILE","id_user=$id_user,nama_lengkap=$nama_lengkap, image=".$id_user.".$ext ");
        header("location: ../index?module=$_GET[module]&message=success3");
    } else  {
        header("location: ../index?module=$_GET[module]&message=error3");

      }



}







?>