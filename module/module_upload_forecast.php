<?php
$module=$_GET['module'];
$pm=$_GET['pm'];
$page_tmp = $_SERVER['PHP_SELF']."?module=$module&pm=$pm";
$page=str_replace(".php","",$page_tmp);
?>





<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- /.modal -->
<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			

			<!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- BEGIN PAGE HEADER-->
			<h3 class="page-title">
			Upload Forecast<small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Maintenance</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Upload Forecast</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
            <!-- MODAL INSERT -->
				<!-- BEGIN PAGE CONTENT-->
			<div class="row">
				<div class="col-md-12">
					<div class="portlet box blue" id="form_wizard_1">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i> Upload Parameter Forecast  <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                     

							<form action="<?php echo "module/actions_master.php?module=$module&act=upload-forecast";?>" class="form-horizontal" id="form_sample_3" method="POST">
								
									<div class="form-body">
                             
									  	<div class="alert alert-danger display-hide">
									  		<button class="close" data-close="alert"></button>
											Form tidak diisi dengan benar , Silahkan dicek kembali ...!
								  		</div>
											
								<?php
                      if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Password Berhasil dirubah....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Password Gagal Dirubah... ! </strong> </div>";

	}
	?>
									
											
												<!--<h3 class="block"><b>Ganti Password</b></h3>-->
												<div class="form-group">
													<label class="control-label col-md-3">Pilih Tanggal<span class="required">
													* </span>
													</label>
													<div class="col-md-4">
												<div class="input-group input-medium date date-picker" data-date="" data-date-format="yyyy-mm-dd" data-date-viewmode="years">
												<input type="text" class="form-control" readonly>
												<input type="hidden" name="tanggal" class="form-control" >
												<span class="input-group-btn">
												<button class="btn default" type="button"><i class="fa fa-calendar"></i></button>
												</span>
											</div>
													</div>
												</div>
												<div class="form-group">
													<label class="control-label col-md-3">Nilai Interest <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="nilai_interest" id="nilai_interest" class="form-control" required/>
														
													</div>
												</div>
												
												<div class="form-group">
													<label class="control-label col-md-3">Nilai LDR <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="nilai_ldr" id="nilai_ldrl" class="form-control" required/>
														
													</div>
												</div>
                                                <div class="form-group">
													<label class="control-label col-md-3">Nilai Asumsi EIR <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="nilai_asumsi" id="nilai_asumsi" class="form-control" required/>
														
													</div>
												</div>
                                                
                                                <div class="form-group">
													<label class="control-label col-md-3">Nilai Employee Loan Benefit <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="nilai_elb" id="nilai_elb" class="form-control" required/>
														
													</div>
												</div>
											
										

										
									
										<div class="form-actions">
											<div class="row">
												<div class="col-md-offset-3 col-md-9">
													<input type="submit" class="btn green" value="SUBMIT"/>
                                                
												
												</div>
											</div>
										</div>
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
			





<!-- BEGIN PAGE LEVEL PLUGINS -->


<!-- BEGIN PAGE LEVEL STYLES -->
<script src="assets/global/plugins/jquery.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery-migrate.min.js" type="text/javascript"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/additional-methods.min.js"></script>
<script src="assets/admin/pages/scripts/form-validation.js"></script>

<!-- END PAGE LEVEL STYLES -->
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
				
				
            },
			nilai_interest: {
			   required: true
				
				
            },
			nilai_ldr: {
			   required: true
				
				
            },
			nilai_asumsi: {
			   required: true
				
				
            },
			nilai_elb: {
			    required: true
               
            }

			
        },
		

        messages: {
			tanggal: {
			required: "<span class='label label-warning'> <i>Tanggal Harus Dipilih.. !</i></span>"
            },
			nilai_interest: {
			required: "<span class='label label-warning'> <i>Tidak Boleh Kosong ... !</i></span>"
            },
			nilai_ldr: {
			required: "<span class='label label-warning'> <i>Tidak Boleh Kosong ... !</i></span>"
            },
			nilai_asumsi: {
			required: "<span class='label label-warning'> <i>Tidak Boleh Kosong ... !</i></span>"
            },
			nilai_elb: {
			required: "<span class='label label-warning'> <i>Tidak Boleh Kosong ... !</i></span>"
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