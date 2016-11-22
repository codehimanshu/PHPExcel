$(document).ready(function(){
	$(".export").click(function(){
		var table = $(this).attr('id');
		$.ajax({
			type: 'POST',
			url: 'Examples/export.php',
			data: {
				'table':table, 
			},
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