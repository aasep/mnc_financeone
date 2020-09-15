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
			Upload Suku Bunga<small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Maintenance</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Upload Suku Bunga</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
            <!-- MODAL INSERT -->
				<!-- BEGIN PAGE CONTENT-->



 <div class="modal fade bs-modal-lg" id="delete-modal" tabindex="-1" role="dialog" aria-hidden="true" aria-labelledby="myModalLabel">
								<div class="modal-dialog">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title" id="myModalLabel"><strong>Delete Modal</strong></h4>
										</div>
										<div class="modal-body">
							<form action="<?php echo "module/actions_master.php?module=$module&act=delete_skb_efektif"; ?>" class="form-horizontal" method="POST">
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










			<div class="row">
				<div class="col-md-12">
					<div class="portlet box blue" id="form_wizard_1">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i> Upload Range Suku Bunga <span class="step-title">
								 </span>
							</div>
							
						</div>
						<div class="portlet-body form">
                      
                     

							<form action="<?php echo "module/actions_master.php?module=$module&act=skb_efektif";?>" class="form-horizontal" id="form_sample_3" method="POST">
								
									<div class="form-body">
                             
									  	<div class="alert alert-danger display-hide">
									  		<button class="close" data-close="alert"></button>
											Form tidak diisi dengan benar , Silahkan dicek kembali ...!
								  		</div>
											
								<?php
                      if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data Berhasil Diatambahkan....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Data gagal ditambahakan... ! </strong> </div>";
}

	 if (isset($_GET['message']) && ($_GET['message']=="success2")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Delete Data Berhasil....!  </strong> </div>";

	}
	
	if (isset($_GET['message']) && ($_GET['message']=="error2")){
	echo "<div class='alert alert-warning alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong> Delete Data Gagal... ! </strong> </div>";

	}
	?>
									
											
												<!--<h3 class="block"><b>Ganti Password</b></h3>-->
											<!--	<div class="form-group">
										<label class="control-label col-md-3">Periode Pelaporan </label>
										<div class="col-md-4">
											<div class="input-group input-large date-picker input-daterange" data-date="" data-date-format="yyyy-mm-dd">
												<input type="text" class="form-control" name="from">
												<span class="input-group-addon">
												to </span>
												<input type="text" class="form-control" name="to">
											</div>
											
										</div>
									</div>
									-->

									<div class="form-group">
										<label class="control-label col-md-3">Pilih Tanggal </label>
										<div class="col-md-3">
											<div class="input-group input-medium date date-picker" data-date="" data-date-format="yyyy-mm-dd" data-date-viewmode="years" data-date-minviewmode="months">
												<input type="text" class="form-control" readonly>
												<span class="input-group-btn">
												<input type="hidden" name="from" id="from" class="form-control" >
												<button class="btn default" type="button"><i class="fa fa-calendar"></i></button>
												</span>
											</div>
											<!-- /input-group -->
											
										</div>
									</div>

												<div class="form-group">
													<label class="control-label col-md-3">Mata Uang<span class="required">
													* </span>
													</label>
													<div class="col-md-4">
														<select class="form-control" name="mata_uang" id="mata_uang">
														<option value="">-Pilih  Mata Uang-</option>
														<option value="1">IDR</option>
														<option value="2">VALAS</option>
														
														</select>
													</div>
												</div>
												<div class="range">
												
												
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
								<i class="fa fa-globe"></i>Nilai Suku Bunga Efektif
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
									 Range Atas
								</th>
								<th>
									 Range Bawah
								</th>
								<th>
									 Range Valas
								</th>
								
								<th>
									 Action
								</th>
								
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query="select * from Master_SKB_Efektif ";
                            $result=odbc_exec($connection2,$query);
                            while ($row=odbc_fetch_array($result)) {
                            $tgl=date('d-m-Y',strtotime($row['DataDate']));
							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>".date('d-m-Y',strtotime($row['DataDate']))."</td>";
							echo "<td>".number_format((float)$row[Range_Atas], 2, '.', '')."</td>";
							echo "<td>".number_format((float)$row[Range_Bawah], 2, '.', '')."</td>";
							echo "<td>".number_format((float)$row[Range_Valas], 2, '.', '')."</td>";
							echo "<td> <a href='#'  data-toggle='modal' id-datadate='$row[DataDate]'  data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
							echo "</tr>";
							$i++;
							
									}
							?>


							</tbody>
							</table>
						</div>
					</div>


             
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

$('#mata_uang').change(function() {
		var id=$(this).val();

		var dataString = 'id='+id;
		
		//alert(dataString1);
		$.ajax({
		type: "POST",
		url: "module/ajax_sukubunga.php",
		data: dataString,
		cache: false,
		success: function(html)
		{
			$(".range").html(html);
		} 
			});
});
}); // document ready	$(document).ready(function() {

    </script>	




<script>
$(document).ready(function() {

	$('.detailDelete').click(function() {
		var datadate = $(this).attr('id-datadate');
		  //alert(datadate);
		 // alert(nominalmodal);
			$("#list-user").empty();
			$("#list-user").append( 
                '<tr>'+
				'<td>'+'<h5>Yakin Anda Akan Mendelete Suku Bunga '+'Tanggal : <strong>' + datadate +'</strong></h5>'+
				'<input type="hidden" name="tgl" value="'+datadate+'">'+
				'</td></tr>');
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
		ignore: "",
        rules: {

			from: {
			   required: true	
            },
			to: {
			   required: true
            },
			range_atas: {
			   required: true
            },
			range_bawah: {
			   required: true
            }
	
        },
		

        messages: {
			from: {
			required: "<span class='label label-warning'> <i>From Harus Dipilih.. !</i></span>"
            },
			to: {
			required: "<span class='label label-warning'> <i>To Harus Dipilih ... !</i></span>"
            
            },
			range_atas: {
			required: "<span class='label label-warning'> <i>Tdk Boleh Kosong ... !</i></span>"
            
            },
			range_bawah: {
			required: "<span class='label label-warning'> <i>Tdk Boleh Kosong ... !</i></span>"
           
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