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

$dsn="develop";
$usr="lsmk";
$pass="password";
$connection = odbc_connect($dsn, $usr, $pass);
	
	if(!$connection)
		die('Failed to connect to server: ' . odbc_error());	
		
		echo "<br>ok <br>";

$query="select top 10 * from dm_journal";
$result=odbc_exec($connection,$query);
while ( $row=odbc_fetch_array($result)) {

	echo $row['DataDate']." ". $row['JenisMataUang']." ".$row['KodeGL']." ".$row['Nominal']."</br>";
}


?>