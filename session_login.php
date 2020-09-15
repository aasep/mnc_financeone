<?php
//session_start();
if (!isset($_SESSION['SESS_USERNAME']) || !isset($_SESSION['SESS_STATUS_ACCOUNT']) || !isset($_SESSION['SESS_PASSWORD']))
{
	header("location: temp_session_login");
	die();
}
?>