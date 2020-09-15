<?php
include "session_login.php";
$module=$_GET['module'];
$pm=$_GET['pm'];
$page_tmp = $_SERVER['PHP_SELF']."?module=$module&pm=$pm";
$page=str_replace(".php","",$page_tmp);
?>
<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- /.modal -->
			<!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- BEGIN PAGE HEADER-->
			<h3 class="page-title">
			Laporan Keuangan Publikasi <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Report</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Laporan Publikasi</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
			
			
			<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Laporan Keuangan Publikasi
							</div>
							<div class="tools">
								<a href="javascript:;" class="collapse">
								</a>
								<a href="#portlet-config" data-toggle="modal" class="config">
								</a>
								<a href="javascript:;" class="reload">
								</a>
								<a href="javascript:;" class="remove">
								</a>
							</div>
						</div>
						<div class="portlet-body form">
							<!-- BEGIN FORM-->
							<form action="<?php echo "module/report/report_nop.php"; ?>" id="form_sample_3" class="form-horizontal" method="POST"> 
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										You have some form errors. Please check below.
									</div>
									<div class="alert alert-danger" style="display:none" id="kosong">
										<button class="close" data-close="alert"></button>
										Form Ada yang Kosong ... !
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Tahun  <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="tahun" id="tahun">
												<?php echo getTahunBefore();  ?>
												
											</select>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Kuartal  <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="kuartal" id="kuartal">
												<option value="">Pilih Kuartal</option>
												<option value="1">Kuartal 1 [ Januari s.d Maret ]</option>
												<option value="2">Kuartal 2 [ April s.d Juni ]</option>
												<option value="3">Kuartal 3 [ Juli s.d September ]</option>
												<option value="4">Kuartal 4 [ Oktober s.d Desember </option>
												
											</select>
										</div>
									</div>
									

								<div class="form-actions">
									<div class="row">
										<div class="col-md-offset-3 col-md-9">
											<button class="btn blue" type="button" id="exp-flash"> Export </button>
											
										</div>
									</div>
								</div>


								</div>
								
								<div  class="excel"></div>
								</br>
								<div align="center" class="loading2" style="display:none">
								<img src="images/loading_image.gif"  width="100" id="loading" align="center" >
								</br></br></br></br>
								</div>
							
						</div>
					</div>
					<!-- END VALIDATION STATES-->

<script src="assets/global/plugins/jquery.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery-migrate.min.js" type="text/javascript"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/additional-methods.min.js"></script>
<script src="assets/admin/pages/scripts/form-validation.js"></script>

<script type="text/javascript">
$(document).ready(function()
{	
$("#exp-flash").click(function()
{
var tahun=document.getElementById("tahun").value;
var id=document.getElementById("kuartal").value;
if (id !=''){
$("#kosong").hide();
$(".excel").hide();
var dataString1 = 'kuartal='+ id +'&tahun='+ tahun;
//alert(id);
$(".loading2").show(); 

$.ajax
({
type: "POST",
url: "module/report/ajax_report_keu_publikasi.php",
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