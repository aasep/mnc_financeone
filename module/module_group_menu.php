<?php
$module=$_GET['module'];
$pm=$_GET['pm'];
$page_tmp = $_SERVER['PHP_SELF']."?module=$module&pm=$pm";
$page=str_replace(".php","",$page_tmp);

if(isset($_POST['id_group'])){

$id_group=$_POST['id_group'];	
} else {
$id_group=$_GET['id_group'];	

}

?>
<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- /.modal -->
			<!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- BEGIN PAGE HEADER-->
			<h3 class="page-title">
			Group Menu <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<a href="#">Management User</a>
						<i class="fa fa-angle-right"></i>
					</li>
					<li>
						<a href="#">Goup Menu</a>
						<i class="fa fa-angle-right"></i>
					</li>
					
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
			<!-- BEGIN PAGE CONTENT-->
		<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-gift"></i>Group Menu
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
							<form action="<?php echo "$page"; ?>" id="form_sample_3" class="form-horizontal" method="POST">
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
										<label class="control-label col-md-3">Nama Group User <span class="required">
										* </span>
										</label>
										<div class="col-md-4">
											<select class="form-control" name="id_group" id="id_group">
												<option value="">Pilih Group User</option>
												<?php
												$query_group="select * from group_user order by nama_group asc";
												$result_group=odbc_exec($connection, $query_group);
												while ($row_group=odbc_fetch_array($result_group)){
													if($id_group==$row_group['id_group']) {
													echo "<option value='$row_group[id_group]' selected='selected'>$row_group[nama_group]</option>"; 
													} else {
													echo "<option value='$row_group[id_group]' >$row_group[nama_group]</option>"; 

													}
												}

												?>
											
												
											</select>

										</div>
									</div>

									<div class="form-group">
										<label class="control-label col-md-3">.
										</label>
										<div class="col-md-4">
											<input type="submit" value="Submit" class="btn blue"/>
										</div>
									</div>
									
									</form>
								</div>
								
							
						</div>
					</div>
			<!-- END PAGE CONTENT-->
			




            <!-- BEGIN EXAMPLE TABLE PORTLET-->
					<div class="portlet box green-haze">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-globe"></i>Daftar Group Menu
							</div>
							<div class="tools">
								<a href="javascript:;" class="reload">
								</a>
								<a href="javascript:;" class="remove">
								</a>
							</div>
						</div>
						<div class="portlet-body">
						<form name="form1" method="post" action="" > 
							<table class="table table-striped table-bordered table-hover" id="sample_2">
							<thead>
							<tr>
							<th>
									No
								</th>
								<th>
									# 
								</th>
								<th>
								
								</th>
								<th>
								Nama menu
								</th>
								<th>
									Status
								</th>
								
							</tr>
							</thead>
							<tbody>
							<?php
							 if (isset($id_group) ){
							$no=1;
							$no_parent=1;
							$query="select * from menu where parent='0' ";
							$result=odbc_exec($connection, $query);
							while ($row=odbc_fetch_array($result)) {
							//<input type="checkbox" class="checkboxes" value="1"/>
							$no_sub=1;	
                            echo "<tr>";
                            echo "<td>$no</td>";
                            echo "<td>$no_parent</td>";
                            echo "<td></td>";
                            echo "<td><b>$row[nama_menu]</b></td>";
                            echo "<td></td></tr>";
                            
                            
							$query2="select * from menu where parent='$row[id_menu]' ";
							$result2=odbc_exec($connection, $query2);
							while ($row2=odbc_fetch_array($result2)) {
								

							$query3="select id_group_menu from group_menu where id_menu='$row2[id_menu]' and id_group='$id_group' ";
							$result3=odbc_exec($connection, $query3);
							$found=odbc_num_rows($result3);

							if($found >=1){
								//$check="checked='checked'";
								$status="<a class='btn green disabled' href='$page&act=edit_menu&nama_menu=$row2[nama_menu]&parent=$row2[parent]&id_menu=$row2[id_menu] disable='disabled''>Active</a>";

							} else {
								//$check="";
								$status="<a href='#'  data-toggle='modal' id-menu='$row2[id_menu]' id-nama='$row2[nama_menu]' data-target='#delete-modal' class='detailDelete' > <button class='btn red disabled' >In Active</button></a>";
							}



							$no++;
							echo "<tr>";
							echo "<td>$no</td>";
                            echo "<td>$no_parent.$no_sub</td>";
                            echo "<td> <input type='checkbox' name='checkbox[]' class='checkboxes' value='$row2[id_menu]'  /></td>";
                            echo "<td>$row2[nama_menu]</td>";
                            echo "<td> $status</td></tr>";

                                $no_sub++;
                                

							} 
							   $no_parent++;
							    $no++;

							}//end if	
								
							}
							?>
							
							
							</tbody>
							</table>

						</div>

					</div>


					<!-- END EXAMPLE TABLE PORTLET-->
					<div class="portlet default">
						
						<div class="portlet-body form">
							<!-- BEGIN FORM-->
							
								<div class="form-body">

									<div class="form-group">
										<label class="control-label col-md-3">.
										</label>
										<div class="col-md-3">
											<input type="submit" value=" Aktifkan " name="hidup" id="hidup" class="btn green"/>
										</div>
										<div class="col-md-3">
											<input type="submit" value=" Non Aktifkan " name="mati" id="mati"  class="btn red"/>
											 <input type="hidden" name="id_group" value="<?php echo $id_group;?>">
										</div>
									</div>
									
									
								</div>
								
							
						</div>
					</div>

	<?php
//echo $sql;
// Check if delete button active, start this
if(isset($_POST['hidup'])){
for($i=0;$i<count($_POST['checkbox']);$i++){
$del_id=$_POST['checkbox'][$i];

// cek apakah id_menu dg group user X di tabel group menu sudah ada
$query_cek="select id_group_menu FROM group_menu  WHERE id_group=$id_group AND id_menu='$del_id'";
			 $result_cek = odbc_exec($connection,$query_cek);
		     $found_priv = odbc_num_rows($result_cek);

if ($found_priv >=1)
{  
	$result_priv=1;
	} else {
		//insert
		$sql_priv="insert into group_menu (id_group,id_menu) values ('$id_group','$del_id')";
	 $result_priv = odbc_exec($connection,$sql_priv);
		}

} // end loop for
// if successful redirect to delete_multiple.php
if($result_priv)
{
echo "<meta http-equiv=\"refresh\" content=\"0;URL=index.php?module=$module&id_group=$id_group\">";
}
} // end if isset post
//echo $sql_priv."</br>";

//jika tekan tombol non aktif
if(isset($_POST['mati'])){
for($i=0;$i<count($_POST['checkbox']);$i++){
$del_id=$_POST['checkbox'][$i];

// cek apakah id_menu dg group user X di tabel group menu sudah ada
$query_cek="select id_group_menu FROM group_menu  WHERE id_group='$id_group' AND id_menu='$del_id'";
			 $result_cek = odbc_exec($connection,$query_cek);
		     $found_priv = odbc_num_rows($result_cek);

if ($found_priv >=1)
{  //delete
$sql_priv="delete from group_menu where id_group='$id_group' AND id_menu='$del_id' ";
		 $result_priv = odbc_exec($connection,$sql_priv);
	
	} else {
		$result_priv=1;
	 $result_priv = pg_query($connection,$sql_priv);
		}

} 
if($result_priv)
{
echo "<meta http-equiv=\"refresh\" content=\"0;URL=index.php?module=$module&id_group=$id_group\">";
}
} 
?>    

 </form>				