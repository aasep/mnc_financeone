<li >
					<a href="#">
					
					<span class="title">
					</span>
					</a>
				</li>


<?php
// active parent
$active_parent=$_GET['pm'];

// active sub menu
$active_submenu=$_GET['module'];

$parent_count=0;
$query_parent_menu="select distinct (a.parent)  FROM menu a, group_menu b WHERE a.id_menu=b.id_menu AND b.id_group=$_SESSION[SESS_GROUP_USER]";
$result_parent_menu = odbc_exec($connection,$query_parent_menu);
    while ($row = odbc_fetch_array($result_parent_menu))
	   { 
	   //===masang class parent menu 
	   $class_parent="";
		   if ($parent_count==0)
		   {$class='icon-docs';}
		   if ($parent_count==1)
		   {$class='fa fa-th-list';}
		   if ($parent_count==2)
		   {$class='icon-wallet';}
		   if ($parent_count==3)
		   {$class='icon-briefcase';}
		   if ($parent_count==4)
		   {$class='icon-briefcase';}
		   if ($parent_count==5)
		   {$class='icon-briefcase';}
		   
		   
		   $parent_menu=$row['parent'];
		   if ($active_parent==sha1($parent_menu)) {$class_parent="class='start active open'";} 
		   
		   $query_nama_menu="SELECT nama_menu FROM menu WHERE id_menu='$parent_menu'";
		   $result_nama_menu = odbc_exec($connection,$query_nama_menu);
		   $r_menu = odbc_fetch_array($result_nama_menu);
		 //display parent menu 
		 
		  	   
		 echo "<li $class_parent ><a href='#'><i class='$class'></i><span class='title'> $r_menu[nama_menu] </span><span class='arrow '></span></a>";
		   
		// SUBMENU=========

		   $q_submenu="select distinct (a.nama_menu),a.id_menu  FROM menu a, group_menu b WHERE a.id_menu=b.id_menu AND b.id_group=$_SESSION[SESS_GROUP_USER]  AND a.parent='$parent_menu' order by a.nama_menu asc";
		   $result_submenu = odbc_exec($connection,$q_submenu);
		   $found_submennu = odbc_num_rows($result_submenu);
		   if ($found_submennu >= 1)
		   {
		   echo "<ul class='sub-menu'>";
 while ($r_submenu = odbc_fetch_array($result_submenu))
		    {
			$class_submenu="";
         if ($active_submenu==sha1($r_submenu['id_menu'])) {$class_submenu="class='active'";} 
		echo"<li $class_submenu ><a href='index?module=".sha1($r_submenu['id_menu'])."&pm=".sha1($parent_menu)."'>  $r_submenu[nama_menu]</a></li>";

			  //$submenu++;
		    } //end while loop submenu
			echo "</ul></li>"; 
		   } // end submenu found
		   $parent_count++;
		   }
		      
		   
/*		   
?>

				<li >  <!--  class="start active open"   if active    -->
					<a href="javascript:;">
					<i class="fa fa-cogs"></i>
					<span class="title">
					Page Layouts </span>
					<span class="arrow open">
					</span>
					<!-- <span class="selected"></span>   bentuk panah -->
					</a>
					<ul class="sub-menu">
						<li class="active">  <!--  class="active"   if active    -->
							<a href="<?php echo "index.php?module=".sha1('0005');?>">
							Modul 1 </a>
						</li>
						<li>
							<a href="<?php echo "index.php?module=".sha1('0006');?>">
							Modul 2 </a>
						</li>
						
						
					</ul>
				</li>
			
            
<?php

*/
?>
				
				
				<li class="last">
					<a href="logout">
					<i class="fa fa-user"></i>
					<span class="title bold">
					Logout</span>
					</a>
				</li>
                
                
            