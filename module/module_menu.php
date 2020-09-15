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
			Menu <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Management User</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Menu</a>
						<i class="fa fa-angle-right"></i>
					</li>

				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
			<!-- MODAL INSERT-->
			
							<div class="modal fade" id="basic" tabindex="-1" role="basic" aria-hidden="true">
								<div class="modal-dialog modal-lg">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title">Management Menu</h4>
										</div>
										<div class="modal-body">
											<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Add Menu
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
							<form action="<?php echo "module/actions_master.php?module=$module&act=add_menu"; ?>" id="form_sample_3" class="form-horizontal" method="POST">
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										Form tidak diisi dengan benar , Silahkan dicek kembali ...!
									</div>
									<div class="alert alert-success display-hide">
										<button class="close" data-close="alert"></button>
										Your form validation is successful!
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Nama Menu<span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="text" name="nama_menu" id="nama_menu"  class="form-control" required/>
										</div>
									</div>
									
									
									<div class="form-group">
										<label class="control-label col-md-3">Parent <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="parent" id="parent">
												<option value="0">No Parent (as Parent)</option>
												<?php
												$query_parent="select * from menu where parent='0' order by nama_menu asc";
												$result_parent=odbc_exec($connection, $query_parent);
												while ($row_parent=odbc_fetch_array($result_parent)){

													echo "<option value='$row_parent[id_menu]'>$row_parent[nama_menu]</option>"; 

												}

												?>
											
												
											</select>
										</div>
									</div>
									
									
								</div>
								
							
						</div>
					</div>
					<!-- END VALIDATION STATES-->
										</div>
										<div class="modal-footer">
											<button type="button" class="btn default" data-dismiss="modal">Close</button>
											<button type="submit" class="btn blue">Submit</button>
										</div>
									</div>
									<!-- /.modal-content -->
									</form>
							<!-- END FORM-->
								</div>
								<!-- /.modal-dialog -->
							</div>
							<!-- /.modal -->
             
					<!--  END MODAL INSERT -->


					<!-- MODAL UPDATE-->
					<?php
					if (isset($_GET['act']) && $_GET['act']=="edit_menu")
					{
					?>
			
							<div class="modal show" id="basic" tabindex="-1" role="basic" aria-hidden="true">
								<div class="modal-dialog modal-lg">
									<div class="modal-content">
										<div class="modal-header">
											<a href="<?php echo $page; ?>"><button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button></a>
											<h4 class="modal-title">Management Menu</h4>
										</div>
										<div class="modal-body">
											<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Add Menu
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
							<form action="<?php echo "module/actions_master.php?module=$module&act=edit_menu"; ?>"  class="form-horizontal" method="POST">
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										You have some form errors. Please check below.
									</div>
									<div class="alert alert-success display-hide">
										<button class="close" data-close="alert"></button>
										Your form validation is successful!
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Nama Menu<span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="text" name="nama_menu" id="nama_menu" value="<?php echo $_GET['nama_menu']?>" class="form-control" required/>
											<input type="hidden" name="id_menu" id="id_menu" value="<?php echo $_GET['id_menu']?>"/>
										</div>
									</div>
									
									
									<div class="form-group">
										<label class="control-label col-md-3">Parent <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="parent" id="parent">
												<option value="0" <?php if ($_GET['parent']=="0") {echo "selected=selected";}?>>No Parent (as Parent)</option>
												<?php
												$query_parent="select * from menu where parent='0' order by nama_menu asc";
												$result_parent=odbc_exec($connection, $query_parent);
												while ($row_parent=odbc_fetch_array($result_parent)){
                                                    if ($_GET['parent']==$row_parent['id_menu']){
													echo "<option value='$row_parent[id_menu]' selected='selected'>$row_parent[nama_menu]</option>"; 
													} else {
													echo "<option value='$row_parent[id_menu]' >$row_parent[nama_menu]</option>"; 
												        }    
												   }

												?>
												
											</select>
										</div>
									</div>
									
									
								</div>
								
							
						</div>
					</div>
					<!-- END VALIDATION STATES-->
										</div>
										<div class="modal-footer">
											<a href="<?php echo $page; ?>"><button type="button" class="btn default" data-dismiss="modal">Close</button></a>
											<button type="submit" class="btn blue">Submit</button>
										</div>
									</div>
									<!-- /.modal-content -->
									</form>
							<!-- END FORM-->
								</div>
								<!-- /.modal-dialog -->
							</div>
							<!-- /.modal -->
             
					<!--  END MODAL UPDATE -->
						<?php
					}
					?>
			

					<!-- MODAL DELETE -->
					 <!-- MODAL DELETE -->

             <div class="modal fade bs-modal-lg" id="delete-modal" tabindex="-1" role="dialog" aria-hidden="true" aria-labelledby="myModalLabel">
								<div class="modal-dialog">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title" id="myModalLabel"><strong>Delete User</strong></h4>
										</div>
										<div class="modal-body">
							<form action="<?php echo "module/actions_master.php?module=$module&act=delete_menu"; ?>" class="form-horizontal" method="POST">
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
					<!-- END MODAL DELETE -->							

             <a class="btn blue" data-toggle="modal" href="#basic">Add Menu <i class="fa fa-plus"></i> </a> </br> </br>
             

             <?php
    if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>menu Berhasil ditambahkan ...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>menu tidak berhasil diinput...!</strong> </div>";
    }
	if (isset($_GET['message']) && ($_GET['message']=="success2")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Delete Menu Berhasil...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error2")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Delete Menu Gagal...!</strong> </div>";

	}
	if (isset($_GET['message']) && ($_GET['message']=="success3")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Update Menu Berhasil...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error3")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Update Menu Gagal...!</strong> </div>";

	}
	?>

			<!-- BEGIN PAGE CONTENT-->
			
			<!-- END PAGE CONTENT-->
			<!-- BEGIN EXAMPLE TABLE PORTLET-->
					<div class="portlet box blue-hoki">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-globe"></i>Daftar List Menu 
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
									 NO
								</th>
								
								<th>
									 Nama Menu 
								</th>
								<th>
									 Parent
								</th>
							    <th>
									 Kode
								</th>
								<th>
									 Action
								</th>
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
							$no_parent=1;
                            $query="select * from menu where parent='0' order by id_menu asc ";
                            $result=odbc_exec($connection,$query);
                            while ($row=odbc_fetch_array($result)) {
                          
                             $submenu=1;	
							echo "<tr>";
							
							echo "<td width='5%'>$no_parent</td>";
							echo "<td width='50%'><b>$row[nama_menu]</b></td>";
							echo "<td width='5%'>".$row['parent']."</td>";
							echo "<td width='10%'>$row[id_menu]</td>";
							echo "<td width='30%'><a class='btn default' href='$page&act=edit_menu&nama_menu=$row[nama_menu]&parent=$row[parent]&id_menu=$row[id_menu]'>edit</a> <a href='#'  data-toggle='modal' id-menu='$row[id_menu]' id-nama='$row[nama_menu]' data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
							echo "</tr>";

                            $query2="select * from menu where parent='$row[id_menu]' order by nama_menu asc ";
                            $result2=odbc_exec($connection,$query2);
                            while ($row2=odbc_fetch_array($result2)) {

                            echo "<tr>";
							echo "<td width='5%'>$no_parent . $submenu </td>";
							echo "<td width='50%'>$row2[nama_menu]</td>";
							echo "<td width='5%'>".$row2['parent']."</td>";
							echo "<td width='10%'>$row2[id_menu]</td>";
							echo "<td width='30%'><a class='btn default' href='$page&act=edit_menu&nama_menu=$row2[nama_menu]&parent=$row2[parent]&id_menu=$row2[id_menu]'>edit</a> <a href='#'  data-toggle='modal' id-menu='$row2[id_menu]' id-nama='$row2[nama_menu]' data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
							echo "</tr>";
                              $submenu++;
                            }

							$no_parent++;


							
									}
							?>


							</tbody>
							</table>
						</div>
					</div>
					<!-- END EXAMPLE TABLE PORTLET-->



<!-- END PAGE LEVEL STYLES -->


<!-- END CORE PLUGINS -->
<!-- BEGIN PAGE LEVEL PLUGINS -->
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/jquery.validate.min.js"></script>
<script type="text/javascript" src="assets/global/plugins/jquery-validation/js/additional-methods.min.js"></script>
<!-- END PAGE LEVEL PLUGINS -->
<!-- BEGIN PAGE LEVEL STYLES -->
<script src="assets/admin/pages/scripts/form-validation.js"></script>
<script src="assets/global/plugins/jquery.min.js" type="text/javascript"></script>
<script>
$(document).ready(function() {

	$('.detailDelete').click(function() {
		var id_menu = $(this).attr('id-menu');
    	var nama_menu = $(this).attr('id-nama');
		  //alert(id_user);
			$("#list-user").empty();
			$("#list-user").append( 
                '<tr>'+
				'<td>'+'<h5>Yakin Anda Akan Mendelete Menu "'+'<strong>' + nama_menu +' " </strong></h5>'+
				'<input type="hidden" name="id_menu" value="'+id_menu+'">'+
				'</td></tr>');
			  });

}); // document ready	$(document).ready(function() {

    </script>

 <script>

 
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
		    
            nama_menu: {
                required: true
            }
        },

        messages: {
		    nama_menu: {
                 required: "<span class='label label-warning'><i>Nama Menu  Harus diisi ...! </i></span>"
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