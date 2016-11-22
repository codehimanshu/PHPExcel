<!DOCTYPE html>
<html lang="en">
<head>
  	<title>Break Interview</title>
  	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
  	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script>
  	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <script src="script.js"></script>
</head>
<body style="background-color: #d8d8d8; height: 100%;">

    <!-- the default navbar used on website -->
    <nav class="navbar navbar-inverse" id="navbar">
        <div class="container-fluid">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#myNavbar">
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                <a class="navbar-brand" href="#">Break Interview</a>
            </div>
            <div class="collapse navbar-collapse" id="myNavbar">
                <ul class="nav navbar-nav navbar-right">
                    <li><a href="#"><span class="glyphicon glyphicon-log-in"></span> Logout</a></li>
                </ul>
            </div>
        </div>
    </nav>

    <div class="container" style="height: 87%; position: fixed; width: 100%;">
    	<div clas="row" style="height: 100%;">
    	   	<div class="col-xs-6" style="border-right: 4px solid grey; height: 100%; overflow: scroll;">
                <?php
                    require_once 'Examples/db_connect.php';
                    $result = mysql_query("SHOW TABLES",$link);
                    if($result)
                    {
                        echo "Select Table:<br>";
                        while($table = mysql_fetch_assoc($result))
                        {
                            echo "<center><button type=button class='export btn btn-lg btn-primary' id='".$table["Tables_in_phpexcel"]."'>".$table["Tables_in_phpexcel"]."</button></center><br>";
                        }
                    }
                    else
                        echo "No tables found";
                ?>

	    		<!-- <center><button type="button" class="btn btn-lg btn-primary" id="export">Export DB to Excel</button></center> -->
                <span id="export_result"></span>
	    	</div>

	    	<div class="col-xs-6" style="height: 100%;">
                <center>
                <form action="#"" method="post" enctype="multipart/form-data">
                    Select Excel (.xlsx) file to upload: 
                    <input type="file" name="fileToUpload" id="fileToUpload">
                    <input type="submit" value="Upload File" name="submit">
                </form>
                </center>
                <span>
                    <?php
                    if(isset($_POST['submit'])):
                        $target_dir = "Examples/";
                        $target_file = $target_dir . basename($_FILES["fileToUpload"]["name"]);
                        $uploadOk = 1;
                        $FileType = pathinfo($target_file,PATHINFO_EXTENSION);
                        // Allow certain file formats
                        if(!($FileType == "xlsx" || $FileType =="xls")) {
                            echo "Sorry, only xlsx files are allowed.";
                            $uploadOk = 0;
                        }
                        // Check if $uploadOk is set to 0 by an error
                        if ($uploadOk == 0) {
                            echo "Sorry, your file was not uploaded.";
                        // if everything is ok, try to upload file
                        } else {
                            echo $target_file;
                            if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"],"Examples/file.xlsx")) {
                                echo "The file ". basename( $_FILES["fileToUpload"]["name"]). " has been uploaded.";
                            } else {
                                echo "Sorry, there was an error uploading your file.";
                            }
                        }
                    ?>
                </span>
                <br><br><br>
	    		<center><button type="button" class="btn btn-lg btn-primary" id="import">Import Excel to DB</button></center>
                <?php 
                    endif;
                ?>
                <span id="import_result"></span>
	    	</div>    	
    	</div>
    </div>

</body>
</html>