<?php
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

//linking database
require_once 'db_connect.php';

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


// Create new PHPExcel object
echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
$objPHPExcel->getProperties()->setCreator("Himanshu Agrawal")
							 ->setLastModifiedBy("Himanshu Agrawal")
							 ->setTitle("Questions")
							 ->setSubject("Questions")
							 ->setDescription("Test Questions for Exam Module")
							 ->setKeywords("office PHPExcel php")
							 ->setCategory("Test result file");

//getting table structure and creating template
echo date('H:i:s') , "Creating Template" , EOL;							 
$table = $_POST['table'];
$structure = mysql_query("DESCRIBE ".$table,$link) or die(mysql_error());
$columns = [];
// var_dump($columns);
if($structure)
{
	$col = 'A';
	$col_no = 0;
	echo "Fields found: ";
	while($field = mysql_fetch_assoc($structure))
	{
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($col.'1',$field["Field"]);
		echo $field["Field"].", ";
		// echo $col_no;
		array_push($columns,$field['Field']);
		$col_no++;
		$col++;
	}
	echo "<br>";
}
else
	return "Failure";

$total_columns = $col_no-1;
//writing records
$i=2;
echo date('H:i:s') , "Fetching Records" , EOL;
$results = mysql_query('SELECT * FROM '.$table) or die(mysql_error());
if($results)
{
	while($result = mysql_fetch_assoc($results))
	{
		// var_dump($result);
		$col_no = 0;
		$col = 'A';
		// echo $total_columns;
		while($col_no <= $total_columns)
		{
			// echo "DD<br>";
			// var_dump($columns);
			// echo $columns[$col_no];
			// echo $col_no;
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue($col.$i,$result[$columns[$col_no]]);
			$col++;
			$col_no++;
		}
		// $objPHPExcel->setActiveSheetIndex(0)
		// ->setCellValue('A'.$i,$result['id'])
		// ->setCellValue('B'.$i,$result['title'])
		// ->setCellValue('C'.$i,$result['details'])
		// ->setCellValue('D'.$i,$result['topicId'])
		// ->setCellValue('E'.$i,$result['ques_level'])
		// ->setCellValue('F'.$i,$result['ques_type']);
		$i++;
	}
}
else
 	die("Table Empty");

// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$objPHPExcel->getActiveSheet()->setTitle('Questions');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " <strong> File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL,"</strong>";
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing files" , EOL;
echo 'Files have been created in ' ,'<strong>', getcwd(),'</strong>' , EOL;
