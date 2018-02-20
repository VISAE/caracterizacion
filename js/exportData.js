$(document).ready(function(e) {
	$('#button-a').click(function(){
		if(wb !== undefined) {
			wb = XLSX.utils.table_to_book(document.getElementById('matriz'),{sheet:"Resultados"});
			fillWithComments();
			var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
			saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "Resultado.xlsx");
		}
	});
});

function fillWithComments() {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	for(var property in comments) {
		// console.log(property);
		if(comments.hasOwnProperty(property)) {
			sheet[property].c = [];
			sheet[property].c.push(comments[property]);
		}
	}
}

function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for(var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}
