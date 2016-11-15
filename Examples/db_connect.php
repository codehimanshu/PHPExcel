<?php
    require_once 'db_config.php';
	$link=mysql_connect(DB_SERVER,DB_USER,DB_PASSWORD);
	$db=mysql_select_db(DB_DATABASE,$link) or die(mysql_error()) or die(mysql_error());
?>