<?php
require_once '../config/config.php';
//include('db.php');
if($_POST['id'])
{
$id=$_POST['id'];

if ($id=='1'){
?>
<div class="form-group">
													<label class="control-label col-md-3">Range Atas <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="range_atas" id="range_atas" class="form-control" required/>
														
													</div>
												</div>
                                                
												<div class="form-group">
													<label class="control-label col-md-3">Range Bawah <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="range_bawah" id="range_bawah" class="form-control" required/>
														
													</div>
												</div>


<?php
} else {
?>
<div class="form-group">
													<label class="control-label col-md-3">Range Valas <span class="required">
													* </span>
													</label>
													<div class="col-md-3">
														<input type="number" name="range_valas" id="range_valas" class="form-control" required/>
														
													</div>
												</div>



<?php
}
}
?>