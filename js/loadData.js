var wb, comments = {};
$(document).ready(function(e) {
    $("#matriz").tableHeadFixer();

	$('#input-excel').change(function(e){
		try {
            var reader = new FileReader();

            //Loads Spinner
            $("#loading").fadeIn();
            var opts = {
                lines: 12, // The number of lines to draw
                length: 7, // The length of each line
                width: 4, // The line thickness
                radius: 10, // The radius of the inner circle
                color: '#000', // #rgb or #rrggbb
                speed: 1, // Rounds per second
                trail: 60, // Afterglow percentage
                shadow: false, // Whether to render a shadow
                hwaccel: false // Whether to use hardware acceleration
            };
            var trget = document.getElementById('loading');
            var spnr = new Spinner(opts).spin(trget);

            reader.readAsArrayBuffer(e.target.files[0]);
            reader.onload = function (e) {
                var data = new Uint8Array(reader.result);
                wb = XLSX.read(data, {type: 'array'});
                fillTable();
            }
        }catch (err) {
			console.log("File Error!");
		} finally {
            setTimeout(function() {
                spnr.stop(); // Stop the spinner
            }, 4000);
		}
	});
});