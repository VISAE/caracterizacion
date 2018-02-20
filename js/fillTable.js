function manageUndefined(cell) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	try {
		return sheet[cell].v;
	} catch(e) {
		return "";
	}
}


function fillTable() {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	var rows = XLSX.utils.decode_range(sheet["!ref"]).e.r;
	var cols = XLSX.utils.decode_range(sheet["!ref"]).e.c;
	// document.getElementById("msg").innerHTML = rows;
	var strRows = '';
	for (var i = 3; i <= rows; i++) {
		//console.log(wb.Sheets.SocioDemograficos['A'+(i+1)].v + " " + wb.Sheets.SocioDemograficos['D'+(i+1)].v);
		strRows += "<tr>";
		strRows += "<td>"+manageUndefined('A'+(i+1))+"</td>"; // IDENTIFICACION DEL ESTUDIANTE
		strRows += "<td>"+manageUndefined('B'+(i+1))+" "+manageUndefined('C'+(i+1))+" "+manageUndefined('D'+(i+1))+"</td>"; // NOMBRE COMPLETO
		strRows += "<td>"+manageUndefined('F'+(i+1))+"</td>"; // CENTRO
		strRows += "<td>"+manageUndefined('G'+(i+1))+"</td>"; // ZONA
		strRows += "<td>"+manageUndefined('E'+(i+1))+"</td>"; // PROGRAMA
		strRows += "<td>"+manageUndefined('H'+(i+1))+"</td>"; // ESCUELA
		strRows += "<td>"+manageUndefined('I'+(i+1))+"</td>"; // EMAIL 
		strRows += "<td>"+manageUndefined('CA'+(i+1))+"</td>"; // TELEFONO
		strRows += "<td>"+manageUndefined('AD'+(i+1))+"</td>"; // EDAD
		strRows += "<td></td>"; // ASIGNACION
		strRows += "<td>"+manageUndefined('X'+(i+1))+"</td>"; // CONVENIO
		strRows += "<td></td>"; // NOVEDAD
		strRows += "<td></td>"; // GENERAL
		strRows += "<td></td>"; // CAMPUS VIRTUAL
		var rFSD = riesgoFSD(i+1);
		strRows += "<td"+addComment(rFSD, 'O'+(i-1))+">"+rFSD[0]+"</td>"; // Riesgo Factor Socio-Demográfico		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		var rFSE = riesgoFSE(i+1);
		strRows += "<td"+addComment(rFSE, 'R'+(i-1))+">"+rFSE[0]+"</td>"; // Riesgo Factor Socio-Económico		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // Riesgo Factor Academico Antecedentes		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // Riesgo Factor Académico por apropiacion al modelo		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // Riesgo Factor Institucional		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // Riesgo Factores Externos		
		strRows += "<td></td>"; // Acciones realizadas según ruta de PAPC
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // NIVEL DE RIESGO POR FACTORES
		strRows += "<td></td>"; // GRUPO COLABORATIVO
		strRows += "<td></td>"; // SITUACION DE RIESGO
		strRows += "<td></td>"; // ACCIONES REALIZADAS
		strRows += "<td></td>"; // RESULTADOS
		strRows += "<td></td>"; // No. Cursos 
		strRows += "<td></td>"; // SITUACION DE RIESGO
		strRows += "<td></td>"; // ACCIONES REALIZADAS
		strRows += "<td></td>"; // RESULTADOS

	}
	document.getElementById("tableBody").innerHTML = strRows;
}


function addComment(comment, cellExport) {
	if (comment.length>1) {
		comments[cellExport] = {a:'VISAE', t:comment[1]};
		return " title='"+comment[1]+"'";
	} else {
		return '';
	}
}

// Analisis de los riesgos
function riesgoFSD(row) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
    var na = ['No Aplica', '#N/A'];
    var riesgos = new Array;
    riesgos.push(sheet['N'+row].v == 'Rural'?'Area Residencia:\n\tRural':null);
    riesgos.push(sheet['J'+row].v== 'Madre Soltera'?'Estado Civil:\n\tMadre Soltera':null);
    riesgos.push(na.indexOf(sheet['R'+row].v) < 0?'Disc. Auditiva:\n\t'+sheet['R'+row].v:null);
    riesgos.push(na.indexOf(sheet['S'+row].v) < 0?'Disc. Cognitiva:\n\t'+sheet['S'+row].v:null);
    riesgos.push(na.indexOf(sheet['U'+row].v) < 0?'Disc. Emocional:\n\t'+sheet['U'+row].v:null);
    riesgos.push(na.indexOf(sheet['T'+row].v) < 0?'Disc. Fisica:\n\t'+sheet['T'+row].v:null);
    riesgos.push(na.indexOf(sheet['V'+row].v) < 0?'Disc. Mental:\n\t'+sheet['V'+row].v:null);
    riesgos.push(na.indexOf(sheet['Q'+row].v) < 0?'Disc. Visual:\n\t'+sheet['Q'+row].v:null);
    riesgos.push(sheet['X'+row].v == 'Interno'?'Convenio INPEC:\n\tInterno':null);
    msgFSD = '';
    riesgos.forEach(function(e){
	    msgFSD += (e!=null?'* '+e+'\n':'');
	});
	return msgFSD.length > 0?["Riesgo por condiciones personales",msgFSD]:["Sin riesgo por condiciones personales"];
}

function riesgoFSE(row) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
    var riesgos = new Array;
    riesgos.push(sheet['AO'+row].v == 'Menos de un salario mínimo'?'Ingresos Mensuales:\n\tMenos de un salario mínimo':null);
    riesgos.push(sheet['AE'+row].v == 'Desempleado'?'Situación Laboral:\n\tDesempleado':null);
    msgFSE = '';
    riesgos.forEach(function(e){
		msgFSE += (e!=null?'* '+e+'\n':'');
	});
	return msgFSE.length > 0?["Riesgo socioeconomico", msgFSE]:["Sin riesgo socioeconomico"];	
}