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
			Input Premi LPS <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Maintenance</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Input Premi LPS</a>
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
											<h4 class="modal-title" id="myModalLabel"><strong>Delete Premi</strong></h4>
										</div>
										<div class="modal-body">
							<form action="<?php echo "module/actions_master.php?module=$module&act=delete_premi"; ?>" class="form-horizontal" method="POST">
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
								<i class="fa fa-gift"></i> Input Premi LPS <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
							<form action="<?php echo "module/actions_master.php?module=$module&act=input_premi";?>" class="form-horizontal" id="form_sample_3" method="POST" enctype="multipart/form-data">
								
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
												<div class="form-group">
										<label class="control-label col-md-3">Tahun  <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="tahun" id="tahun">
												<option value="">Pilih Tahun </option>
												<option value="2015">2015</option>
												<option value="2016">2016</option>
												<option value="2017">2017</option>
												<option value="2018">2018</option>
												<option value="2019">2019</option>
												
											</select>
										</div>
									</div>
												<div class="form-group">
													<label class="control-label col-md-3">Semester <span class="required">
													* </span>
													</label>
													<div class="col-md-4">
														<select class="form-control" name="semester" id="semester">
														<option value="">- Pilih Semester -</option>
														<option value="1">Semester 1</option>
														<option value="2">Semester 2</option>
														</select>
													</div>
												</div>
												<div class="form-group">
													<label class="control-label col-md-3">Jml. Premi Verifikasi LPS <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="premi_ver" id="premi_ver" class="form-control" required/>
														
													</div>
												</div>

												<div class="form-group">
													<label class="control-label col-md-3">Saldo Premi Bulan Sebelumya <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="premi_sebelumnya" id="premi_sebelumnya" class="form-control" required/>
														
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
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-globe"></i>Daftar Saldo Premi LPS
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
									 Periode Awal
								</th>
								<th>
									 Periode Akhir
								</th>
								<th>
									 Jml Premi Verifikasi 
								</th>
								<th>
									 Saldo Premi Periode Lalu
								</th>
								<th>
									 Action
								</th>
								
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query="select * from Master_Saldo_Premi_LPS ";
                            $result=odbc_exec($connection2,$query);
                            while ($row=odbc_fetch_array($result)) {
                            $tgl=date('d-m-Y',strtotime($row['DataDate']));
							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>".date('m-Y',strtotime($row['Periode_Awal']))."</td>";
							echo "<td>".date('m-Y',strtotime($row['Periode_Akhir']))."</td>";
							echo "<td>".number_format((float)$row['Jumlah_Premi_Verifikasi_LPS'], 2, '.', '')."</td>";
							echo "<td>".number_format((float)$row['Saldo_Premi_Periode_Lalu'], 2, '.', '')."</td>";
						
							echo "<td> <a href='#'  data-toggle='modal' id-akhir='$row[Periode_Akhir]' id-awal='$row[Periode_Awal]' id-tgl='$tgl' data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
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
		var awal = $(this).attr('id-awal');
    	var akhir = $(this).attr('id-akhir');
		  //alert(datadate);
		 // alert(nominalmodal);
			$("#list-user").empty();
			$("#list-user").append( 
                '<tr>'+
				'<td>'+'<h5>Yakin Anda Akan Mendelete "'+'Tanggal Awal : <strong>' + awal +' " </strong>'+'dg Tanggal Akhir <strong> ' + akhir +' " </strong></h5>'+
				'<input type="hidden" name="awal" value="'+awal+'">'+
				'<input type="hidden" name="akhir" value="'+akhir+'">'+
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

			premi_sebelumnya: {
			   required: true	
            },
            premi_ver: {
			   required: true
            }, 
			nama_file: {
			    required: true,
                extension: "xls"
            }

			
        },
		

        messages: {
			premi_sebelumnya: {
			required: "<span class='label label-warning'> <i>Premi Sebelumnya  harus diisi.. !</i></span>"
            },
            premi_ver: {
			required: "<span class='label label-warning'> <i>Premi Verifikasi harus diisi.. !</i></span>"
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