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
			<h3 class="page-title">
			Perhitungan Pembayaran Premi<small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Report LPS</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Perhitungan Pembayaran Premi LPS</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
			<!-- BEGIN PAGE CONTENT-->
			<div class="row">
				<div class="col-md-12">
					<div class="portlet box blue" id="form_wizard_1">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i> Perhitungan Pembayaran Premi LPS<span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                     

							<form action="<?php echo "module/actions_master.php?module=$module&act=change-pass";?>" class="form-horizontal" id="form_sample_3" method="POST">
								
									<div class="form-body">
                             
									  	<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										You have some form errors. Please check below.
									</div>
									<div class="alert alert-danger" style="display:none" id="kosong">
										<button class="close" data-close="alert"></button>
										Form Ada yang Kosong ... !
									</div>
											
								<?php
                      if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Password Berhasil dirubah....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Password Gagal Dirubah... ! </strong> </div>";

	}
	?>
									
											
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
										<label class="control-label col-md-3">Pilih Semester  <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="semester" id="semester">
												<option value="">Pilih Semester</option>
												<option value="1">Semester 1 </option>
												<option value="2">Semester 2</option>
												
												
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



							</form>	
						</div>
					</div>
				</div>
			</div>
			<!-- END PAGE CONTENT-->
             <!-- END MODAL INSERT -->


         


             
                <!-- END MODAL DELETE -->

			<!-- BEGIN PAGE CONTENT-->
			
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
var tahun=document.getElementById("tahun").value;
var id=document.getElementById("semester").value;
if (id !=''){
$("#kosong").hide();
$(".excel").hide();
var dataString1 = 'semester='+ id +'&tahun='+ tahun;
//alert(id);
$(".loading2").show(); 

$.ajax
({
type: "POST",
url: "module/report/ajax_report_premi_lps.php",
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

			tahun: {
			   required: true
			},
			semester: {
			   required: true	
			}
            
	
        },
		

        messages: {
			tahun: {
			required: "<span class='label label-warning'> <i>Tahun Harus Dipilih.. !</i></span>"
            },
            semester: {
			required: "<span class='label label-warning'> <i>Semester Harus Dipilih.. !</i></span>"
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
			
			<!-- END PAGE CONTENT-->
			
			