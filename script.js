$(document).ready(function(){
	$("#export").click(function(){
		$.ajax({
			type: 'GET',
			url: 'Examples/export.php',
			success: function(data){
				$("#export_result").html(data);
			}
		});
	});

	$("#import").click(function(){
		$.ajax({
			type: 'GET',
			url: 'Examples/import.php',
			success: function(data){
				$("#import_result").html(data);
			}
		});
	});
})