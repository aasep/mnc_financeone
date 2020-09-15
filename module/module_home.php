<!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- /.modal -->
			<!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
			<!-- BEGIN PAGE HEADER-->
			<h3 class="page-title font-blue-steel font-lg bold " >
			Management Report <small></small>
			</h3>
			<div class="page-bar">
				<ul class="page-breadcrumb">
					<li>
						<i class="fa fa-home"></i>
						<small><a href="index.html">Home</a></small>
						<i class="fa fa-angle-right"></i>
					</li>
				
					
				</ul>
				
			</div>
			<!-- END PAGE HEADER-->
			<!-- BEGIN PAGE CONTENT-->
			


			<div class="row">
				<div class="col-md-12">
					<!-- BEGIN SAMPLE TABLE PORTLET-->
					<div class="portlet box blue">
						<div class="portlet-title">
							<div class="caption">
								<i class="fa fa-comments"></i>Welcome to Finance One Report
							</div>
							<!--
							<div class="tools">
								<a href="javascript:;" class="collapse">
								</a>
								<a href="#portlet-config" data-toggle="modal" class="config">
								</a>
								<a href="javascript:;" class="reload">
								</a>
								<a href="javascript:;" class="remove">
								</a>
							</div>-->
						</div>
						<div class="portlet-body">
							<div class="table-scrollable">
								<table class="table table-bordered table-hover" width="100%">
								<thead>
								<tr>
									<th width="20%">
										 #
									</th>
									<th width="2%">
										
									</th>
									<th width="78%">
										 Information
									</th>
								
									
								</tr>
								</thead>
								<tbody>
								<tr class="active">
									<td>
										Username
									</td>
									<td>
										 :
									</td>
									<td>
										 <?php echo getUsername();?>
									</td>
									
									
								</tr>
								<tr class="success">
									<td>
										 Group Name
									</td>
									<td>
										:
									</td>
									<td>
										 <?php echo getGroupUserName()?>
									</td>
								</tr>
								<tr class="active">
									<td>
										Last Login
									</td>
									<td>
										 :
									</td>
									<td>
										 <?php echo date('d-m-Y H:i',strtotime(lastLogin())); ?>
									</td>
									
									
								</tr>
								
								</tbody>
								</table>
							</div>
						</div>
					</div>
					<!-- END SAMPLE TABLE PORTLET-->
				</div>
				
			</div>
			<div class="row">
				<div class="col-md-12">
					
					<!-- END BORDERED TABLE PORTLET-->
					<!--<div align="center"><img src="images/financial.png" style="height: 75px;" alt=""/> -->

					</div>
				</div>
			</div>
			<!-- END PAGE CONTENT-->