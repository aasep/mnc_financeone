<?php

 // && !isset($_GET['module'])
	if (!isset($_SESSION['SESS_GROUP_USER']) )
	{
		session_destroy();
		header("location: temp_session_login.php");
		
		} else if (isset($_SESSION['SESS_GROUP_USER']) ){
			
			if (isset($_GET['module']))
			
			{     //cek di tabel menu
			       $module=$_GET['module'];
				
			       $sql="select id_menu from menu where src='$module'";  
				   $res= odbc_exec($connection,$sql);
				   $found_menu=odbc_num_rows($res);
				   $row = odbc_fetch_array($res);
				   $id_menu=$row['id_menu'];

				   if ($found_menu >=1)
				   {
			      $query="SELECT id_group_menu FROM group_menu where id_group=$_SESSION[SESS_GROUP_USER] AND id_menu='$id_menu'";
				  $result = odbc_exec($connection,$query);
				  $found=odbc_num_rows($result);
			      //jika tidak ketemu				  
				       if ($found ==0)
				            {
						       header("location: temp_session_group.php");
						
						      }
				
				
			        } else  { //header("location: temp_session_group.php");  
					}//end found_menu
					  
					
			  } //end isset module
				
		} // end else
				

?>