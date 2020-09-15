<?php

##########################################################################
#            -= YOU MAY NOT REMOVE OR CHANGE THIS NOTICE -=              #
# -----------------------------------------------------------------------#
# 																		 #
#  																		 #
#  	Developed by:	Asep Arifyan						    			 #
#	License:		Commercial											 #
#  	Copyright: 		2016. All Rights Reserved.		                     #
#                                                                        #
#  	Additional modules (embedded): 										 #
#	-- Metronic (Themes) 												 #
#																		 #
#																		 #
# -----------------------------------------------------------------------#
#	Designed and built with all the love and loyalty.					 #
##########################################################################
// FOR USER
// $dsn="db_fincon";
// $usr="hr";
// $pass="Instance13";
// $connection = odbc_connect($dsn, $usr, $pass);
	
// 	if(!$connection)
// 		die('Failed to connect to server: ' . odbc_error());	


$dsn="db_fincon";
$usr="sa";
$pass="Instance12";
$connection = odbc_connect($dsn, $usr, $pass);
	
	if(!$connection)
		die('Failed to connect to server: ' . odbc_error());	




// FOR LSMK 
$dsn2="develop";
$usr2="lsmk";
$pass2="password";
$connection2 = odbc_connect($dsn2, $usr2, $pass2);
	
	if(!$connection2)
		die('Failed to connect to server: ' . odbc_error());	

?>