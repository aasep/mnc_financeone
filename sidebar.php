<div class="page-sidebar-wrapper">
		<!-- DOC: Set data-auto-scroll="false" to disable the sidebar from auto scrolling/focusing -->
		<!-- DOC: Change data-auto-speed="200" to adjust the sub menu slide up/down speed -->
		<div class="page-sidebar navbar-collapse collapse">
			<!-- BEGIN SIDEBAR MENU1 DEFAULT -->
			<ul class="page-sidebar-menu hidden-sm hidden-xs" data-auto-scroll="true" data-slide-speed="200">
				<!-- DOC: To remove the search box from the sidebar you just need to completely remove the below "sidebar-search-wrapper" LI element -->
				<!-- DOC: This is mobile version of the horizontal menu. The desktop version is defined(duplicated) in the header above -->
				
				<?php include "sidebar_menu.php";?>
			</ul>
			<!-- END SIDEBAR MENU1 DEFAULT-->
			<!-- BEGIN RESPONSIVE MENU FOR HORIZONTAL & SIDEBAR MENU MOBILE -->
			<ul class="page-sidebar-menu visible-sm visible-xs" data-slide-speed="200" data-auto-scroll="true">
				<?php include "sidebar_menu.php";?>
			</ul>
			<!-- END RESPONSIVE MENU FOR HORIZONTAL & SIDEBAR MENU MOBILE -->
		</div>
	</div>