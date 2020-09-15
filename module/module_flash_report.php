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
			Flash Report <small></small>
			</h3>
			
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
					<small>
						  <a href="#">Daily Report</a>
					</small> 
					</li>
					<!--<a href="#">Internal</a>-->
						<i class="fa fa-angle-right"></i>
					<li>
					<small>
						<a href="#">Flash Report</a>
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
								<i class="fa fa-gift"></i>Input Flash Report
							</div>
							<!--
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
							-->
						</div>
						<div class="portlet-body form">
							<!-- BEGIN FORM-->
							<form action="<?php echo "module/report/flash_report_dev.php"; ?>" id="form_sample_3" class="form-horizontal" method="POST">
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										Form Belum Komplit, Silahkan cek kembali informasi dibawah ... !
									</div>
									<div class="alert alert-success display-hide">
										<button class="close" data-close="alert"></button>
										Your form validation is successful!
									</div>
									<div class="form-group">
									<!--
									<label class="control-label col-md-3">Pilih Tanggal </label> 
									-->
									<label class="control-label col-md-5 font-blue-chambray"><b>Report Date :</b></label>
										<div class="col-md-3">
											<div class="input-group input-medium date date-picker" data-date="" data-date-format="yyyy-mm-dd" data-date-viewmode="years">
												<input type="text" class="form-control" readonly>
												<input type="hidden" name="tanggal" id="tanggal" class="form-control" >
												<span class="input-group-btn">
												<button class="btn default" type="button"><i class="fa fa-calendar"></i></button>
												</span>
											</div>
											<!-- /input-group -->
											
										</div>
									</div>
									
								<div class="form-actions">
									<div class="row">
										<div class="col-md-offset-5 col-md-9">
											<button class="btn blue" type="button" id="exp-flash"> Export </button>
											
										</div>
									</div>
								</div>
							
								</div>
								<div class="alert alert-danger" style="display:none" id="kosong">
										<b>You entered empty date !</b>
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
var id=document.getElementById("tanggal").value;

if ( id!='') {
$("#kosong").hide();
$(".excel").hide();




var dataString1 = 'tanggal='+ id ;

$(".loading2").show(); 
$(".excel").hide();



$.ajax
({
type: "POST",
url: "module/report/ajax_flash_report_dev.php",
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


 <script>
jQuery('.numbersOnly').keyup(function () {
    this.value = this.value.replace(/[^0-9\.]/g,'');
});

jQuery(document).ready(function () {
    var form1 = $('#form_sample_3');
    var error1 = $('.alert-danger', form1);
    var success1 = $('.alert-success', form1);


    $('#form_sample_3').validate({
        errorElement: 'span', //default input error message container
        errorClass: 'help-block', // default input error message class
        focusInvalid: false, // do not focus the last invalid input
		ignore: "",
        rules: {

			tanggal: {
			   required: true	
            }
	
        },
		

        messages: {
			tanggal: {
			required: "<span class='label label-warning'> <i>Tanggal Harus Dipilih.. !</i></span>"
            }
        },

        invalidHandler: function (event, validator) { //display error alert on form submit
            success1.hide();
            $('.alert-danger span').text("Form Belum Komplit, Silahkan cek kembali informasi dibawah ... !");
            $('.alert-danger', $('#form_sample_3')).show();
            Metronic.scrollTo(error1, -200);
        },


    });


});

</script>
			<!-- END PAGE CONTENT-->
						