var wb, comments = {};
$(document).ready(function(e) {
	$('#input-excel').change(function(e){
		var reader = new FileReader();
		reader.readAsArrayBuffer(e.target.files[0]);
		reader.onload = function(e) {
			var data = new Uint8Array(reader.result);
			wb = XLSX.read(data, {type:'array'});
			fillTable();
		}
	});
});