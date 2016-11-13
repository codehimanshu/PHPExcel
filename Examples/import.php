<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */

//displaying errors
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');

//linking database
$link=mysql_connect('localhost','root','root') or die("ERROR Connecting MySQL");
$db=mysql_select_db('phpexcel',$link) or die("ERROR Connecting Database");


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
if($result)
{
	echo "This question already exists<br>";
	echo $Aval." ".$Cval;
	break;
}

//seeding into the DB
$result=mysql_query("INSERT INTO questions VALUES ('','$Aval','$Bval','$Cval','$Dval','$Eval')",$link) or die(mysql_error());

$i++;
}while(!empty($Aval));