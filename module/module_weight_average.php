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
			Weight Average <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Maintenance</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Weight Average</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
            <!-- MODAL INSERT -->
				<!-- BEGIN PAGE CONTENT-->

<!-- MODAL DELETE -->

             <div class="modal fade bs-modal-lg" id="delete-modal" tabindex="-1" role="dialog" aria-hidden="true" aria-labelledby="myModalLabel">
								<div class="modal-dialog">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title" id="myModalLabel"><strong>Delete Modal</strong></h4>
										</div>
										<div class="modal-body">
							<form action="<?php echo "module/actions_master.php?module=$module&act=delete_weight_average"; ?>" class="form-horizontal" method="POST">
                            <div class="table-scrollable">
							<table class="table table-striped table-bordered table-hover" id="list-user">
                            
							</table>
							<!--</div>-->
										</div>
										</div>
										<div class="modal-footer">
											<input type="submit" class="btn default blue" value="Hapus">

										</div>
									</div>
									<!-- /.modal-content -->
									</form>
								</div>
								<!-- /.modal-dialog -->
							 </div>
							<!-- /.modal -->



                <!-- END MODAL DELETE -->







			<div class="row">
				<div class="col-md-12">
					<div class="portlet box blue" id="form_wizard_1">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i> Weight Average <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                     
							<form action="<?php echo "module/actions_master.php?module=$module&act=weight_average";?>" class="form-horizontal" id="form_sample_3" method="POST" enctype="multipart/form-data">
								
									<div class="form-body">
                             
									  	<div class="alert alert-danger display-hide">
									  		<button class="close" data-close="alert"></button>
											Form tidak diisi dengan benar , Silahkan dicek kembali ...!
								  		</div>
											
								<?php
                      if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data berhasil ditambahkan....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Data Gagal Diupload... ! </strong> </div>";

	}

	if (isset($_GET['message']) && ($_GET['message']=="success2")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Delete Berhasil...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error2")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Delete  Gagal...!</strong> </div>";

	}
	?>
									
											
												<!--<h3 class="block"><b>Ganti Password</b></h3>-->
									<!--<div class="form-group">
										<label class="control-label col-md-3">Pilih Tanggal </label>
										<div class="col-md-3">
											<div class="input-group input-medium date date-picker" data-date="" data-date-format="yyyy-mm-dd" data-date-viewmode="years" data-date-minviewmode="months">
												<input type="text" class="form-control" readonly>
												<span class="input-group-btn">
												<input type="hidden" name="tanggal" id="tanggal" class="form-control" >
												<button class="btn default" type="button"><i class="fa fa-calendar"></i></button>
												</span>
											</div>
											
											
										</div>
									</div> -->
												
												<div class="form-group">
										<label class="control-label col-md-3">Tahun  <span class="required">
										* </span>
										</label>
										<div class="col-md-3">
											<select class="form-control" name="tahun" id="tahun">
												<?php echo getTahunBefore();  ?>
											</select>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Bulan  <span class="required">
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


												<div class="form-group">
													<label class="control-label col-md-3"> Nilai Weight Average <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="weight_average" id="weight_average" class="form-control" required/>
														
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


         
<!-- BEGIN EXAMPLE TABLE PORTLET-->
					<div class="portlet box blue-hoki">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-globe"></i>Daftar Weight Average
							</div>
							<div class="tools">
								<a href="javascript:;" class="reload">
								</a>
								<a href="javascript:;" class="remove">
								</a>
							</div>
						</div>
						<div class="portlet-body">
							<table class="table table-striped table-bordered table-hover" id="sample_2">
							<thead>
							<tr>
								<th>
									 No
								</th>
								<th>
									 Data Date
								</th>
								
								<th>
									 Nilai Weight Average
								</th>
								<th>
									 Action
								</th>
								
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query="select * from weighted_average ";
                            $result=odbc_exec($connection2,$query);
                            while ($row=odbc_fetch_array($result)) {
                            $tgl=date('d-m-Y',strtotime($row['DataDate']));
							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>".date('d-m-Y',strtotime($row['DataDate']))."</td>";
							echo "<td>$row[nilai]</td>";
							echo "<td> <a href='#'  data-toggle='modal' id-datadate='$row[DataDate]' id-nilai='$row[nilai]' id-tgl='$tgl' data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
							echo "</tr>";
							$i++;
							
									}
							?>


							</tbody>
							</table>
						</div>
					</div>
					<!-- END EXAMPLE TABLE PORTLET-->

             
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
$(document).ready(function() {

	$('.detailDelete').click(function() {
		var tgl = $(this).attr('id-tgl');
		var datadate = $(this).attr('id-datadate');
    	var nilai = $(this).attr('id-nilai');

			$("#list-user").empty();
			$("#list-user").append( 
                '<tr>'+
				'<td>'+'<h5>Yakin Anda Akan Mendelete Weight Average "'+'Tanggal : <strong>' + datadate +' " </strong>'+'dg nilai <strong> ' + nilai +' " </strong></h5>'+
				'<input type="hidden" name="tgl" value="'+datadate+'">'+
				'<input type="hidden" name="nilai" value="'+nilai+'">'+
				'</td></tr>');
			  });


	$('.detailEdit').click(function() {
		var tgl = $(this).attr('id-tgl');
		var datadate = $(this).attr('id-datadate');
    	var nominalmodal = $(this).attr('id-nominalmodal');

		
		document.getElementById('ed_tanggal').value=tgl;
		document.getElementById('ed_tanggal2').value=tgl;
    	
    	document.getElementById('ed_nilai_modal').value=nominalmodal;
		  
			  });

}); // document ready	$(document).ready(function() {

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