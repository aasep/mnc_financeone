<?php


      if( isset($_GET['module']) ){

                if($_GET['module']==sha1('2'))			{ // Module user
				   include "module/module_user.php";				   
			     }else if($_GET['module']==sha1('3')){  // Module Group User
				   include "module/module_group_user.php";
				 }else if($_GET['module']==sha1('4')){  // Module Menu
				   include "module/module_menu.php";
				 }else if($_GET['module']==sha1('5')){  // Module Group Menu
				   include "module/module_group_menu.php";
				 }else if($_GET['module']==sha1('7')){  // Flash Report
				   include "module/module_flash_report.php";
				 }else if($_GET['module']==sha1('8')){  // NOP
				   include "module/module_nop.php";
				 }else if($_GET['module']==sha1('9')){  // SBDK
				   include "module/module_sbdk.php";
				 }else if($_GET['module']==sha1('37')){  // Laporan Simpanan Pihak ke tiga
				   include "module/module_lap_simpanan_lps.php";
				 }else if($_GET['module']==sha1('38')){  // Laporan Counter Rate
				   include "module/module_lap_counter_rate.php";
				 }else if($_GET['module']==sha1('39')){  // Laporan Suku Bunga Efektif
				   include "module/module_lap_suku_bunga_efektif.php";
				 }else if($_GET['module']==sha1('40')){  // Laporan Keuangan Bulanan
				   include "module/module_lap_keuangan_bulanan.php";
				 }else if($_GET['module']==sha1('41')){  // Premi LPS
				   include "module/module_ppp_lps.php";
				 }else if($_GET['module']==sha1('42')){  // KPMM
				   include "module/module_kpmm.php";
				 }else if($_GET['module']==sha1('43')){  // Rasio-rasio Bank
				   include "module/module_rasio_bank.php";
				 }else if($_GET['module']==sha1('44')){  // Laporan Keuangan Publikasi
				   include "module/module_lap_keu_publikasi.php";
				 }else if($_GET['module']==sha1('45')){  // Laporan LTV
				   include "module/module_lap_ltv.php";
				 }else if($_GET['module']==sha1('46')){  // Laporan UMKM
				   include "module/module_lap_umkm.php";
				 }else if($_GET['module']==sha1('47')){  // Laporan Keuangan (Longform)
				   include "module/module_lap_keu_longform.php";
				 }else if($_GET['module']==sha1('64')){  // Ganti Password
				   include "module/module_ganti_password.php";
				 }else if($_GET['module']==sha1('66')){  // parameter report
				   include "module/module_upload_parameter_report.php";
				 }else if($_GET['module']==sha1('67')){  // upload parameter report
				   include "module/module_upload_budget.php";
				 }else if($_GET['module']==sha1('68')){  // upload modal
				   include "module/module_upload_modal.php";
				 }else if($_GET['module']==sha1('69')){  // upload forecast
				   include "module/module_upload_forecast.php";
				 }else if($_GET['module']==sha1('70')){  // Suku bunga
				   include "module/module_upload_suku_bunga.php";
				 }else if($_GET['module']==sha1('71')){  // ajustment
				   include "module/module_upload_adjustment_fr.php";
				 }else if($_GET['module']==sha1('72')){  // upload corporate tax
				   include "module/module_upload_corporate_tax.php";
				 }else if($_GET['module']==sha1('82')){  // corporate tax
				   include "module/module_upload_forecast2.php";
				 }else if($_GET['module']==sha1('84')){  // Upload forecast
				   include "module/module_input_rate_alco.php";
				 }else if($_GET['module']==sha1('85')){  // Input rate Alco
				   include "module/module_input_premi_lps.php";
				 }else if($_GET['module']==sha1('86')){  // Premi LPS
				   include "module/module_weight_average.php";
				 }else if($_GET['module']==sha1('87')){  // ATMR
				   include "module/module_input_atmr.php";
				 }else if($_GET['module']==sha1('88')){  // ATMR
				   include "module/module_simpanan_bank_umum.php";
				 }else if($_GET['module']==sha1('89')){  // ATMR
				   include "module/module_edit_profile.php";
				 }else if($_GET['module']==sha1('92')){  // ATMR
				   include "module/exportGL_flash_report.php";
				 }else if($_GET['module']==sha1('93')){  // ATMR
				   include "module/exportGL_longform.php";
				 }else {   //
				 include "module/notfound.php";
				 }
} else {
	 include "module/module_home.php";
	}

?>


