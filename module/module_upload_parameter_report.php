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
			Upload Parameter Report <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Maintenance</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Upload Parameter Report</a>
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
								<i class="fa fa-gift"></i> Upload Parameter Report <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                     

							<form action="<?php echo "module/report/action_parameter_report.php?module=$module&act=report";?>" class="form-horizontal" id="form_sample_3" method="POST" enctype="multipart/form-data">
								
									<div class="form-body">
                             
									  	<div class="alert alert-danger display-hide">
									  		<button class="close" data-close="alert"></button>
											Form tidak diisi dengan benar , Silahkan dicek kembali ...!
								  		</div>
											
								<?php
                      if (isset($_GET['type'])){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Tabel REFERENSI $_GET[type] berhasil diupdate....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Password Gagal Dirubah... ! </strong> </div>";

	}
	?>
									
											
												<!--<h3 class="block"><b>Ganti Password</b></h3>-->
												<div class="form-group">
													<label class="control-label col-md-3">Parameter Referensi<span class="required">
													* </span>
													</label>
													<div class="col-md-4">
														<select class="form-control" name="report_type" id="report_type">
														<option value="">-Pilih Referensi-</option>
														<option value="FLASH">Referensi Flash Report</option>
														
														<option value="NII">Referensi NII</option>
                                                       
														</select>
													</div>
												</div>
												<div class="form-group">
													<label class="control-label col-md-3">File <span class="required">
													* </span>
													</label>
													<div class="col-md-4">
														<input type="file" name="nama_file" id="nama_file" class="form-control" required/>
														
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
			<!-- BEGIN PAGE CONTENT-->
             <?php 
					  if (isset($_GET['type'])){
						?>
			<div class="row">
				<div class="col-md-12">
					<div class="portlet box blue" id="form_wizard_1">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i> Upload Parameter Report <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                      <?php 
					  if ($_GET['type']=='FLASH'){
						?>  
						  <table class="table table-striped table-bordered table-hover" id="sample_2">
							<thead>
							<tr>
								<th>
									 No
								</th>
								<th>
									 F Lev 1
								</th>
								<th>
									 F Lev 1 Descr
								</th>
								<th>
									 F Lev 2
								</th>
								<th>
									 F Lev 2 Descr
								</th>
								<th>
									  F Lev 3
								</th>
                                <th>
									 F Lev 3 Descr
								</th>
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query=" select * from Referensi_Flash_Report order by FLASH_LEVEL_3 asc ";
                            $result=odbc_exec($connection2,$query);
                            while ($row=odbc_fetch_array($result)) {
                            
                             

							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>$row[FLASH_Level_1]</td>";
							echo "<td>$row[FLASH_Level_1_Description]</td>";
							echo "<td>$row[FLASH_Level_2]</td>";
							echo "<td>$row[FLASH_Level_2_Description]</td>";
							echo "<td>$row[FLASH_Level_3]</td>";
							echo "<td>$row[FLASH_Level_3_Description]</td>";
							echo "</tr>";
							$i++;
							
									}
							?>


							</tbody>
							</table>
						  
						 <?php 
						  }
					  
					  
					  ?>



						<?php 
					  if ($_GET['type']=='NII'){
						?>  
						  <table class="table table-striped table-bordered table-hover" id="sample_2">
							<thead>
							<tr>
								<th>
									 No
								</th>
								<th>
									 F Lev 3 NII
								</th>
								<th>
									 F Lev 3 NII Descr
								</th>
								
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query=" select * from Referensi_NII order by FLASH_LEVEL_3_NII asc ";
                            $result=odbc_exec($connection2,$query);
                            while ($row=odbc_fetch_array($result)) {
                            
                             

							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>$row[FLASH_Level_3_NII]</td>";
							echo "<td>$row[FLASH_Level_3_NII_Description]</td>";
							
							echo "</tr>";
							$i++;
							
									}
							?>


							</tbody>
							</table>
						  
						 <?php 
						  }
					  
					  
					  ?>

							
						</div>
					</div>
				</div>
			</div>
            <?php 
					  }
						?>
			<!-- END PAGE CONTENT-->
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
        rules: {

			report_type: {
			   required: true
				
				
            },
			nama_file: {
			    required: true,
                extension: "xls"
            }

			
        },
		

        messages: {
			report_type: {
			required: "<span class='label label-warning'> <i>ReferensiHarus Dipilih.. !</i></span>"
            },
			nama_file: {
			required: "<span class='label label-warning'> <i>Lampirkan File ... !</i></span>",
            extension: "<span class='label label-warning'> <i>extension file harus '.xls' ... !</i></span>"
			
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