<?php
#########################################
# FINCON DEV							#
# OS : WINDOWS 							#
# DB : MS SQL SERVER					#
# CREATOR : ASEP ARIFYAN				#
# EMAIL	: asep.arifyan@mncbank.co.id	#
#########################################
session_start();
    require_once('config/config.php');
	require_once('function/function.php');


	//date_default_timezone_set("Asia/Jakarta");
	if (isset($_SESSION['SESS_USERNAME']) && isset($_SESSION['SESS_STATUS_ACCOUNT']) && isset($_SESSION['SESS_PASSWORD']))
	{
		header("location: index");
		die();
	}
	if (isset($_POST['username']))
	{  
       // $password=sha1($_POST['password']);
		$password=hashEncrypted($_POST['password']);
		$query_user=" SELECT * FROM user_account where username='".strtolower($_POST['username'])."' and password='$password' ";
        $result_user = odbc_exec($connection,$query_user);
		$found = odbc_num_rows($result_user);

		 if ($found >= 1)
		 {
		 	  $r_account = odbc_fetch_array($result_user);
			  $fix_username=$r_account['username'];
              $fix_password=$r_account['password'];
			  $fix_status_account=$r_account['status_account'];
			  $fix_group_user=$r_account['id_group'];
			  $fix_id_user=$r_account['id_user'];
			  $fix_image=$r_account['image'];
              // ==========Cek again password
			 
			  if ($fix_password != $password){
					$var='1q2w3e4r';
					logActivity("login_failed","wrong password");
			        header("location: login?status=$var&id=1");
					die();
			  }
			 
			   //======== CHECK account status Active or inactive =======
			  if ($fix_status_account == 0){
			  	    logActivity("login_failed","inactive");
				    header('location: temp_session_login?status=inactive');
					die();
			  }
			  
			    $_SESSION['SESS_USERNAME']=$fix_username;
				$_SESSION['SESS_STATUS_ACCOUNT']=$fix_status_account;
			    $_SESSION['SESS_PASSWORD'] = $fix_password;
			    $_SESSION['SESS_GROUP_USER'] = $fix_group_user;
			    $_SESSION['SESS_IMAGE'] = $fix_image;
			    $_SESSION['SESS_IDUSER'] = $fix_id_user;
				logActivity("login","success");
	            header("location: index");


		 } else {
			$var='1q2w3e4r';
			logActivity("login_failed","wrong username or password");
			header("location: login?status=$var&id=2");

			 } // else not found


		}  //end isset post username


?>

<!DOCTYPE html>

<!--[if IE 8]> <html lang="en" class="ie8 no-js"> <![endif]-->
<!--[if IE 9]> <html lang="en" class="ie9 no-js"> <![endif]-->
<!--[if !IE]><!-->
<html lang="en">
<!--<![endif]-->
<!-- BEGIN HEAD -->
<head>
<meta charset="utf-8"/>
<title>Finance One System</title>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta content="width=device-width, initial-scale=1.0" name="viewport"/>
<meta http-equiv="Content-type" content="text/html; charset=utf-8">
<meta content="" name="description"/>
<meta content="" name="author"/>
<!-- BEGIN GLOBAL MANDATORY STYLES -->
<link href="assets/global/css/google_css.css" rel="stylesheet" type="text/css"/>
<link href="assets/global/plugins/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css"/>
<link href="assets/global/plugins/simple-line-icons/simple-line-icons.min.css" rel="stylesheet" type="text/css"/>
<link href="assets/global/plugins/bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css"/>
<link href="assets/global/plugins/uniform/css/uniform.default.css" rel="stylesheet" type="text/css"/>
<!-- END GLOBAL MANDATORY STYLES -->
<!-- BEGIN PAGE LEVEL STYLES -->
<link href="assets/admin/pages/css/login2.css" rel="stylesheet" type="text/css"/>
<!-- END PAGE LEVEL SCRIPTS -->
<!-- BEGIN THEME STYLES -->
<link href="assets/global/css/components.css" id="style_components" rel="stylesheet" type="text/css"/>
<link href="assets/global/css/plugins.css" rel="stylesheet" type="text/css"/>
<link href="assets/admin/layout/css/layout.css" rel="stylesheet" type="text/css"/>
<link href="assets/admin/layout/css/themes/darkblue.css" rel="stylesheet" type="text/css" id="style_color"/>
<link href="assets/admin/layout/css/custom.css" rel="stylesheet" type="text/css"/>
<!-- END THEME STYLES -->
<link rel="shortcut icon" type="image/x-icon" href="http://www.mncbank.co.id/assets/Addresbarwebsite(Favicon).png">
</head>
<!-- END HEAD -->
<!-- BEGIN BODY -->
<body class="login">
<!-- BEGIN SIDEBAR TOGGLER BUTTON -->
<div class="menu-toggler sidebar-toggler">
</div>
<!-- END SIDEBAR TOGGLER BUTTON -->
<!-- BEGIN LOGO -->
<div class="logo">
	<a href="login">
	<img src="images/logo-mnc-bank.png" style="height: 100px;" alt=""/>
	</a>
</div>
<!-- END LOGO -->
<!-- BEGIN LOGIN -->
<div class="content">
	<!-- BEGIN LOGIN FORM -->
	<form class="login-form" action="login.php" method="post">
		<div class="form-title">
			<!--<span class="form-title">MNC BANK APPLICATION REPORT.</span>-->
            <span class="form-subtitle"><center>Login Finance One Report</center></span>
			<!--<span class="form-subtitle">Please login.</span> -->
		</div>
        <?php
        if ( isset ($_GET['status']) && ($_GET['status'])=="1q2w3e4r")
		{
        echo "<div class='alert alert-danger'><span> Username or Password is incorrect ! </span></div>";
        }
        ?>
        
		<div class="alert alert-danger display-hide">
			<!--<button class="close" data-close="alert"></button>-->
			<span>
			Username and Password can't be empty ! </span>
		</div>
		<div class="form-group">
			<!--ie8, ie9 does not support html5 placeholder, so we just show field title for that-->
			<label class="control-label visible-ie8 visible-ie9">Username</label>
			<input class="form-control form-control-solid placeholder-no-fix" type="text" autocomplete="off" placeholder="Username" name="username"/>
		</div>
		<div class="form-group">
			<label class="control-label visible-ie8 visible-ie9">Password</label>
			<input class="form-control form-control-solid placeholder-no-fix" type="password" autocomplete="off" placeholder="Password" name="password"/>
		</div>
		<div class="form-actions">
			<button type="submit" class="btn btn-primary btn-block uppercase">Login</button>
		</div>
		<div class="form-actions">
			<div class="pull-left">
				<label class="rememberme check"> </label>
				
			</div>
			
		</div>
		
		
	</form>
	<!-- END LOGIN FORM -->
	<!-- BEGIN FORGOT PASSWORD FORM -->
	
	<!-- END FORGOT PASSWORD FORM -->
	<!-- BEGIN REGISTRATION FORM -->
	
	<!-- END REGISTRATION FORM -->
</div>
<div class="copyright">
	 2016 © PT Bank MNC Internasional, Tbk.
</div>
<!-- END LOGIN -->
<!-- BEGIN JAVASCRIPTS(Load javascripts at bottom, this will reduce page load time) -->
<!-- BEGIN CORE PLUGINS -->
<!--[if lt IE 9]>
<script src="assets/global/plugins/respond.min.js"></script>
<script src="assets/global/plugins/excanvas.min.js"></script> 
<![endif]-->
<script src="assets/global/plugins/jquery.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery-migrate.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/bootstrap/js/bootstrap.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery.blockui.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/uniform/jquery.uniform.min.js" type="text/javascript"></script>
<script src="assets/global/plugins/jquery.cokie.min.js" type="text/javascript"></script>
<!-- END CORE PLUGINS -->
<!-- BEGIN PAGE LEVEL PLUGINS -->
<script src="assets/global/plugins/jquery-validation/js/jquery.validate.min.js" type="text/javascript"></script>
<!-- END PAGE LEVEL PLUGINS -->
<!-- BEGIN PAGE LEVEL SCRIPTS -->
<script src="assets/global/scripts/metronic.js" type="text/javascript"></script>
<script src="assets/admin/layout/scripts/layout.js" type="text/javascript"></script>
<script src="assets/admin/layout/scripts/demo.js" type="text/javascript"></script>
<script src="assets/admin/pages/scripts/login.js" type="text/javascript"></script>
<!-- END PAGE LEVEL SCRIPTS -->
<script>
jQuery(document).ready(function() {     
	Metronic.init(); // init metronic core components
	Layout.init(); // init current layout
	Login.init();
	Demo.init();
});
</script>
<!-- END JAVASCRIPTS -->
</body>
<!-- END BODY -->
</html>