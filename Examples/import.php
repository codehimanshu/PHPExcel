<?php
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

//linking database
require_once 'db_connect.php';

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';

//searching for file.xlxs
if (!file_exists("file.xlsx")) {
	exit("file.xlxs not found." . EOL);
}

//displaying file found time
echo date('H:i:s') , " Load from Excel2007 file" , EOL;
$callStartTime = microtime(true);

//loading the file
$objPHPExcel = PHPExcel_IOFactory::load("file.xlsx");

//showing the time taken to read file
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;
echo 'Call time to read Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;

//checking the column names
$i = 1;
$Aval = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getValue();
$Bval = $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getValue();
$Cval = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
$Dval = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
$Eval = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();

//validating the template
if($Aval!='title')
	exit( "Expecting 'title' in A column<br>");
if($Bval!='details')
	exit( "Expecting 'details' in A column<br>");
if($Cval!='topicId')
	exit( "Expecting 'topicId' in A column<br>");
if($Dval!='ques_level')
	exit( "Expecting 'ques_level' in A column<br>");
if($Eval!='ques_type')
	exit( "Expecting 'ques_type' in A column<br>");


//getting value of a cell
$i = 2;
do{
$Aval = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getValue();
$Bval = $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getValue();
$Cval = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
$Dval = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
$Eval = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();

//validating the values
if($Aval=='')
{
	echo "title can't be empty<br>";
	$i++;
	break;
}
if(!is_float($Cval))
{
	echo "topicId must be numeric<br>";
	$i++;
	continue;
}
if(!is_float($Dval))
{
	echo "ques_level must be numeric<br>";
	$i++;
	continue;
}
if(!is_float($Eval))
{
	echo "ques_type must be numeric<br>";
	$i++;
	continue;
}

//checking for existing values
$result= mysql_query("SELECT id FROM questions where title='$Aval' AND topicId='$Cval'",$link) or die(mysql_error());
// var_dump($result);
$result = mysql_fetch_array($result);
// var_dump($result);
if($result)
{
	// echo $result;
	echo "This question already exists: ";
	echo $Aval." ".$Cval,'<br>' ;
}

//seeding into the DB
else
{
	$result=mysql_query("INSERT INTO questions VALUES ('','$Aval','$Bval','$Cval','$Dval','$Eval')",$link) or die(mysql_error());
	echo $i-1, " Record Added<br>";
}
$i++;
}while(!empty($Aval));