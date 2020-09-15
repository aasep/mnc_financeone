<?php
$module=$_GET['module'];
$pm=$_GET['pm'];
$page_tmp = $_SERVER['PHP_SELF']."?module=$module&pm=$pm";
$page=str_replace(".php","",$page_tmp);

?>


<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- /.modal -->
			<!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- BEGIN PAGE HEADER-->
			<!--<h3 class="page-title">
			Flash Report <small></small>
			</h3>-->
			<h3 class="page-title font-blue-steel font-lg bold " >
			Export GL Long Form <small></small>
			</h3>
			
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
					<small>
						  <a href="#">Export GL</a>
					</small> 
					</li>
					<!--<a href="#">Internal</a>-->
						<i class="fa fa-angle-right"></i>
					<li>
					<small>
						<a href="#">Long Form </a>
					</small>
					<!--<a href="#">Flash Report</a>-->
						<i class="fa fa-angle-right"></i>
					</li>
				 </ul>
			</div>
			<!-- END PAGE HEADER-->
			<!-- BEGIN PAGE CONTENT-->
			<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Input Long Form 
							</div>
							
						</div>
						<div class="portlet-body form">
							<!-- BEGIN FORM-->
							<form action="" id="form_sample_3" class="form-horizontal" method="POST">
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										Form Belum Komplit, Silahkan cek kembali informasi dibawah ... !
									</div>
									<div class="alert alert-success display-hide">
										<button class="close" data-close="alert"></button>
										Your form validation is successful!
									</div>
									<!--
									<div class="form-group">
													<label class="control-label col-md-5">Jenis Long Form <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<select class="form-control" name="report_type" id="report_type">
														<option value="">- Pilih  Jenis -</option>
														<option value="NERACA"> Neraca </option>
														<option value="LABA RUGI"> Laba - Rugi </option>
														
														</select>
													</div>
												</div>
									-->			
									<div class="form-group">
										<label class="control-label col-md-5">Tahun  <span class="required">
										* </span>
										</label>
										<div class="col-md-3">
											<select class="form-control" name="tahun" id="tahun">
												<?php echo getTahunBefore();  ?>
											</select>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-5">Bulan  <span class="required">
										* </span>
										</label>
										<div class="col-md-3">
											<select class="form-control" name="bulan" id="bulan">
												<option value="">Pilih Bulan </option>
												<option value="01">Januari</option>
												<option value="02">Februari</option>
												<option value="03">Maret</option>
												<option value="04">April </option>
												<option value="05">Mei</option>
												<option value="06">Juni</option>
												<option value="07">Juli</option>
												<option value="08">Agustus</option>
												<option value="09">September</option>
												<option value="10">Oktober</option>
												<option value="11">November</option>
												<option value="12">Desember </option>
												
											</select>
										</div>
									</div>
									<!--
									<div class="form-group">
									
									<label class="control-label col-md-5 font-blue-chambray"><b>Report Date :</b></label>
										<div class="col-md-3">
											<div class="input-group input-medium date date-picker" data-date="" data-date-format="yyyy-mm-dd" data-date-viewmode="years">
												<input type="text" class="form-control" readonly>
												<input type="hidden" name="tanggal" id="tanggal" class="form-control" >
												<span class="input-group-btn">
												<button class="btn default" type="button"><i class="fa fa-calendar"></i></button>
												</span>
											</div>
											
											
										</div>
									</div>
									-->
								<div class="form-actions">
									<div class="row">
										<div class="col-md-offset-5 col-md-9">
											<button class="btn blue" type="button" id="exp-flash"> Export </button>
											
										</div>
									</div>
								</div>
							
								</div>
								<div class="alert alert-danger" style="display:none" id="kosong">
										<b>Month or Year is Empty !</b>
									</div>
								<div  class="excel"></div>
								</br>
								<div align="center" class="loading2" style="display:none">
								<img src="images/loading_image.gif"  width="100" id="loading" align="center" >
								</br></br></br></br>
								</div>

						</div>

						<!--  BEGIN AJAX  -->

						
						<!--  END AJAX  -->
					</div>
					<!-- END VALIDATION STATES-->
			<!-- END PAGE CONTENT-->
			<!-- BEGIN PAGE LEVEL STYLES -->
<script src="assets/global/plugins/jquery.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery-migrate.min.js" type="text/javascript"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/additional-methods.min.js"></script>
<script src="assets/admin/pages/scripts/form-validation.js"></script>

<!-- END PAGE LEVEL STYLES -->
<script type="text/javascript">
$(document).ready(function()
{	
$("#exp-flash").click(function()
{
var id=document.getElementById("bulan").value;
var id2=document.getElementById("tahun").value;

if ( id!='' && id2!='') {
$("#kosong").hide();
$(".excel").hide();

//alert(id2);


var dataString1 = 'bulan='+ id +'&tahun='+id2;

//alert(dataString1);

$(".loading2").show(); 
$(".excel").hide();



$.ajax
({
type: "POST",
url: "module/ajax/GL_longform.php",
data: dataString1,
cache: false,
success: function(html)
{   

	$(".loading2").hide(); 
	$(".excel").show();
	$(".excel").html(html);
} 



});

} else {

$("#kosong").show();

	
}



});




});
</script>


			<!-- END PAGE CONTENT-->
						