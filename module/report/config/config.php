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

$dsn="db_fincon";
$usr="root3";
$pass="Instance12";
$connection = odbc_connect($dsn, $usr, $pass);
	
	if(!$connection)
		die('Failed to connect to server: ' . odbc_error());	
/*		
		echo "<br>ok <br>";

$query="select * from user_account";
$result=odbc_exec($connection,$query);
while ( $row=odbc_fetch_array($result)) {

	echo $row['username']." ". $row['status_account']."</br>";
}
*/

?>