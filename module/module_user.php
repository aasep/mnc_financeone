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
			User Account <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Management User</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">User Account</a>
						<i class="fa fa-angle-right"></i>
					</li>
				
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
            <!-- MODAL INSERT -->
							<div class="modal fade" id="basic" tabindex="-1" role="basic" aria-hidden="true">
								<div class="modal-dialog modal-lg">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title">Add User Account</h4>
										</div>
										<div class="modal-body">
											<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Add User
							</div>
							<div class="tools">
								<a href="javascript:;" class="collapse">
								</a>
								<a href="#portlet-config" data-toggle="modal" class="config">
								</a>
								<a href="javascript:;" class="reload">
								</a>
								
							</div>
						</div>
						<div class="portlet-body form">
							<!-- BEGIN FORM-->
							<form action="<?php echo "module/actions_master.php?module=$module&act=add_user"; ?>" id="form_sample_3" class="form-horizontal" method="POST">
								<div class="form-body">
									<div class="alert alert-danger display-hide">
										<button class="close" data-close="alert"></button>
										Form Belum diisi dengan benar, Silahkan Cek Kembali...!
									</div>
									<div class="alert alert-success display-hide">
										<button class="close" data-close="alert"></button>
										Your form validation is successful!
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Username <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="text" name="username" id="username" placeholder="Username" class="form-control" required/>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Password <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="password" name="password" id="password" placeholder="Password" class="form-control" required/>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Confirm Password <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="password" name="cpassword" id="cpassword" data-required="1" placeholder="Confirm Password" class="form-control"/>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Status Acoount <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="status_account" id="status_account">
												<option value="">Pilih Status Account</option>
												<option value="1">Active</option>
												<option value="0">Non Active</option>
												
											</select>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Group User <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="group_user" id="group_user">
												<option value="">Pilih Group User</option>
												<?php
												$query_group="select * from group_user";
												$result_group=odbc_exec($connection, $query_group);
												while ($row_group=odbc_fetch_array($result_group)) {

												echo "<option value='$row_group[id_group]'>$row_group[nama_group]</option>";
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
             <!-- END MODAL INSERT -->


              <!-- MODAL EDIT -->
              <?php
              if (isset($_GET['act']) && ($_GET['act']=="edit_user")){
              ?>

							<div class="modal show" id="basic" tabindex="-1" role="basic" aria-hidden="true">
								<div class="modal-dialog modal-lg">
									<div class="modal-content">
										<div class="modal-header">
											<a href="<?php echo $page; ?>"><button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button> </a>
											<h4 class="modal-title">Edit User Account</h4>
										</div>
										<div class="modal-body">
											<!-- BEGIN VALIDATION STATES-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Edit User
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
							<form action="<?php echo "module/actions_master.php?module=$module&act=edit_user"; ?>"  class="form-horizontal" method="POST">
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
										<label class="control-label col-md-3">Username <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="text" name="username" id="username"  value="<?php echo $_GET['username'];?>" class="form-control" required/>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Password <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="password" name="password" id="password"  value="" class="form-control" />
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Confirm Password <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<input type="password" name="cpassword" id="cpassword" value="" class="form-control"/>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Status Acoount <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="status_account" id="status_account">
												<option value="">Pilih Status Account</option>
												<option value="1" <?php if ($_GET['status']=="1") {echo "selected='selected'";}?>>Active</option>
												<option value="0" <?php if ($_GET['status']=="0") {echo "selected='selected'";}?>>Non Active</option>
												
											</select>
										</div>
									</div>
									<div class="form-group">
										<label class="control-label col-md-3">Group User <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="group_user" id="group_user">
												<option value="">Pilih Group User</option>
												<?php
												$query_group="select * from group_user";
												$result_group=odbc_exec($connection, $query_group);
												while ($row_group=odbc_fetch_array($result_group)) {
												if ($_GET['id_group']==$row_group['id_group']){
													echo "<option value='$row_group[id_group]' selected='selected'>$row_group[nama_group]</option>";

												}
													echo "<option value='$row_group[id_group]'>$row_group[nama_group]</option>";
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
             <!-- END MODAL EDIT -->

              <?php
             }
              ?>




             <!-- MODAL DELETE -->

             <div class="modal fade bs-modal-lg" id="delete-modal" tabindex="-1" role="dialog" aria-hidden="true" aria-labelledby="myModalLabel">
								<div class="modal-dialog">
									<div class="modal-content">
										<div class="modal-header">
											<button type="button" class="close" data-dismiss="modal" aria-hidden="true"></button>
											<h4 class="modal-title" id="myModalLabel"><strong>Delete User</strong></h4>
										</div>
										<div class="modal-body">
							<form action="<?php echo "module/actions_master.php?module=$module&act=delete_user"; ?>" class="form-horizontal" method="POST">
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

             <a class="btn blue" data-toggle="modal" href="#basic">Add User <i class="fa fa-plus"></i> </a> </br> </br>
             
 <?php
    if (isset($_GET['message']) && ($_GET['message']=="success")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data User Berhasil ditambahkan ...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data tidak berhasil diinput...!</strong> </div>";

	}

if (isset($_GET['message']) && ($_GET['message']=="success2")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data Berhasil dihapus ...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error2")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data Gagal Dihapus...!</strong> </div>";

	}
if (isset($_GET['message']) && ($_GET['message']=="success3")){
	echo "<div class='alert alert-success alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data Berhasil Diupdate ...! </strong> </div>";
     }
	if (isset($_GET['message']) && ($_GET['message']=="error3")){
	echo "<div class='alert alert-danger alert-dismissable'><button type='button' class='close' data-dismiss='alert' aria-hidden='true'></button><strong>Data Gagal Diupdate...!</strong> </div>";

	}




?>
			<!-- BEGIN PAGE CONTENT-->
			
			<!-- END PAGE CONTENT-->
			<!-- BEGIN EXAMPLE TABLE PORTLET-->
					<div class="portlet box blue-hoki">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-globe"></i>Daftar User 
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
									 Username
								</th>
								<th>
									 Status Account
								</th>
								<th>
									 Group User
								</th>
								<th>
									 Date Create
								</th>
								<th>
									 Action
								</th>
							</tr>
							</thead>
							<tbody>
							<?php
							$i=1;
                            $query="select a.id_user,a.username,b.nama_group,a.adddt,a.status_account,a.id_group from user_account a left join  group_user b on a.id_group=b.id_group ";
                            $result=odbc_exec($connection,$query);
                            while ($row=odbc_fetch_array($result)) {
                             if ($row['status_account']=='1'){
                             	$status="Active";
                             } else { $status="InActive"; }
                             //class='detail-sumRO'

							echo "<tr>";
							echo "<td>$i</td>";
							echo "<td>$row[username]</td>";
							echo "<td>$status</td>";
							echo "<td>$row[nama_group]</td>";
							echo "<td>".date('d-m-Y H:i',strtotime($row['adddt']))."</td>";
							echo "<td><a class='btn default' href='$page&act=edit_user&username=$row[username]&status=$row[status_account]&id_group=$row[id_group]'>edit</a> <a href='#'  data-toggle='modal' id-username='$row[username]' data-target='#delete-modal' class='detailDelete' > <button class='btn red'>Delete</button></a></td>";
							echo "</tr>";
							$i++;
							
									}
							?>


							</tbody>
							</table>
						</div>
					</div>
					<!-- END EXAMPLE TABLE PORTLET-->



<!-- END PAGE LEVEL STYLES -->

<!-- IMPORTANT! Load jquery-ui-1.10.3.custom.min.js before bootstrap.min.js to fix bootstrap tooltip conflict with jquery ui tooltip -->
<!--<script src="assets/global/plugins/jquery-ui/jquery-ui-1.10.3.custom.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/bootstrap/js/bootstrap.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/bootstrap-hover-dropdown/bootstrap-hover-dropdown.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery-slimscroll/jquery.slimscroll.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery.blockui.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery.cokie.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/uniform/jquery.uniform.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/bootstrap-switch/js/bootstrap-switch.min.js" type="text/javascript"></script>
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
		var id_user = $(this).attr('id-username');
    
		  //alert(id_user);
			$("#list-user").empty();
			$("#list-user").append( 
                '<tr>'+
				'<td>'+'<h5>Yakin Anda Akan Mendelete '+'<strong>' + id_user +'</strong></h5>'+
				'<input type="hidden" name="id_user" value="'+id_user+'">'+
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
		    
            username: {
                required: true,
                nowhitespace: true
                
            },
			password: {
                required: true
            },
            cpassword: {
                required: true,
                equalTo: "#password"
            },
                status_account: {
                required: true
            },
                group_user: {
                required: true
            }
        },

        messages: {
		    username: {
                 required: "<span class='label label-warning'><i>Username  Harus diisi ...! </i></span>",
                 nowhitespace:"<span class='label label-warning'><i>tidak boleh ada spasi...! </i></span>",
            },
		    password: {
                 required: "<span class='label label-warning'><i>Password  Harus diisi ...! </i></span>"
            },     
			cpassword: {
                 required: "<span class='label label-warning'><i>Confirm Password  Harus diisi ...! </i></span>",
                 equalTo: "<span class='label label-warning'> <i>Password Tidak Sama ... !</i></span>"
            },
            status_account: {
                 required: "<span class='label label-warning'><i>Status Account  Harus diisi ...! </i></span>"
            },
            group_user: {
                 required: "<span class='label label-warning'><i>Group User  Harus diisi ...! </i></span>"
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

