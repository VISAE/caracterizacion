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
	var strRows = '';
	for (var i = 3; i <= rows; i++) {
		strRows += "<tr>";
        strRows += "<td>"+(i-2)+"</td>"; // Nro.
        strRows += "<td>"+manageUndefined('A'+(i+1))+"</td>"; // IDENTIFICACION DEL ESTUDIANTE
		strRows += "<td>"+manageUndefined('B'+(i+1))+" "+manageUndefined('C'+(i+1))+" "+manageUndefined('D'+(i+1))+"</td>"; // NOMBRE COMPLETO
		strRows += "<td>"+manageUndefined('F'+(i+1))+"</td>"; // CENTRO
		strRows += "<td>"+manageUndefined('G'+(i+1))+"</td>"; // ZONA
		strRows += "<td>"+manageUndefined('E'+(i+1))+"</td>"; // PROGRAMA
		strRows += "<td>"+manageUndefined('H'+(i+1))+"</td>"; // ESCUELA
		strRows += "<td>"+manageUndefined('I'+(i+1))+"</td>"; // EMAIL
		strRows += "<td></td>"; // TELEFONO
		strRows += "<td>"+manageUndefined('AD'+(i+1))+"</td>"; // EDAD
		strRows += "<td></td>"; // ASIGNACION
		strRows += "<td></td>"; // CONVENIO
		strRows += "<td></td>"; // NOVEDAD
		strRows += "<td></td>"; // PARTICIPACION INDUCCION
        strRows += "<td></td>"; // PARTICIPACION INMERSION CAMPUS VIRTUAL
        var rFSD = validaRiesgos('FSD', i+1);
		strRows += "<td"+addComment(rFSD, 'P'+(i+3))+">"+rFSD[0]+"</td>"; // Factor Socio-Demográfico
		var rFPS = validaRiesgos('FPS', i+1);
        strRows += "<td"+addComment(rFPS, 'Q'+(i+3))+">"+rFPS[0]+"</td>"; // Factor Psicosocial
        var rFAA = validaRiesgos('FAA', i+1);
        strRows += "<td"+addComment(rFAA, 'R'+(i+3))+">"+rFAA[0]+"</td>"; // Factor Académico Antecedentes
        strRows += "<td></td>"; // Factor Socio-Económico
        strRows += "<td></td>"; // Factor Académico por apropiacion al modelo
        strRows += "<td></td>"; // Factor Institucional
        strRows += "<td></td>"; // Factores Externos
        strRows += "<td></td>"; // COMPETENCIAS DIGITALES BASICAS
        strRows += "<td></td>"; // COMPETENCIAS CUANTITATIVAS
        strRows += "<td></td>"; // COMPETENCIAS LECTO-ESCRITORA
        strRows += "<td></td>"; // COMPETENCIAS DE INGLES
        strRows += "<td></td>"; // MEDIO DE CONTACTO
        strRows += "<td></td>"; // ACCIONES REALIZADAS SEGÚN RUTA DE PAPC
        strRows += "<td></td>"; // RESULTADOS ACCIONES DE ACOMPAÑAMIENTO
        strRows += "<td></td>"; // NIVEL DE RIESGO
        strRows += "<td></td>"; // OBSERVACIONES
        strRows += "<td></td>"; // GRUPO COLABORATIVO
        strRows += "<td></td>"; // SITUACION DE RIESGO
        strRows += "<td></td>"; // ACCIONES REALIZADAS
        strRows += "<td></td>"; // RESULTADOS
        strRows += "<td></td>"; // No. Cursos
        strRows += "<td></td>"; // SITUACION DE RIESGO
        strRows += "<td></td>"; // MEDIO DE CONTACTO
        strRows += "<td></td>"; // ACCIONES REALIZADAS
        strRows += "<td></td>"; // RESULTADOS
        strRows += "<td></td>"; // SITUACION DE RIESGO
        strRows += "<td></td>"; // ACCIONES REALIZADAS
        strRows += "<td></td>"; // RESULTADOS
        strRows += "<td></td>"; // SITUACION DE ALERTA
        strRows += "<td></td>"; // COMPETENCIAS DIGITALES BASICAS
        strRows += "<td></td>"; // COMPETENCIAS CUANTITATIVAS
        strRows += "<td></td>"; // COMPETENCIAS LECTO-ESCRITORA
        strRows += "<td></td>"; // COMPETENCIAS DE INGLES
        strRows += "<td></td>"; // SITUACION DE RIESGO
        strRows += "<td></td>"; // MEDIO DE CONTACTO
        strRows += "<td></td>"; // ACCIONES REALIZADAS
        strRows += "<td></td>"; // RESULTADOS
        strRows += "<td></td>"; // SITUACION DE RIESGO
        strRows += "<td></td>"; // ACCIONES REALIZADAS
        strRows += "<td></td>"; // RESULTADOS
        strRows += "<td></td>"; // SITUACION DE ALERTA
        strRows += "<td></td>"; // MEDIO DE CONTACTO
        strRows += "<td></td>"; // APROBACION DE LA CATEDRA
        strRows += "<td></td>"; // APROBACION DE CURSOS
        strRows += "<td></td>"; // CULMINACION DEL PERIODO ACADEMICO
        strRows += "<td></td>"; // RESULTADO FINAL
        strRows += "<td></td>"; // CONSEJERO ACADEMICO
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
// Factor Socio-Demográfico
function riesgoFSD(row, sheet, cells, na) {
    cells.push(na.indexOf(sheet['L'+row].v.toString().toLowerCase()) < 0?'Etnia:\n\t'+sheet['L'+row].v:null);
    cells.push(sheet['N'+row].v.toString().toLowerCase() == 'rural'?'Area Residencia:\n\tRural':null);
    cells.push(sheet['P'+row].v.toString().toLowerCase() == 'si'?'Desplazado:\n\tSi':null);
    cells.push(na.indexOf(sheet['R'+row].v.toString().toLowerCase()) < 0?'Disc. Auditiva:\n\t'+sheet['R'+row].v:null);
    cells.push(na.indexOf(sheet['S'+row].v.toString().toLowerCase()) < 0?'Disc. Cognitiva:\n\t'+sheet['S'+row].v:null);
    cells.push(na.indexOf(sheet['U'+row].v.toString().toLowerCase()) < 0?'Disc. Emocional:\n\t'+sheet['U'+row].v:null);
    cells.push(na.indexOf(sheet['T'+row].v.toString().toLowerCase()) < 0?'Disc. Fisica:\n\t'+sheet['T'+row].v:null);
    cells.push(na.indexOf(sheet['Q'+row].v.toString().toLowerCase()) < 0?'Disc. Visual:\n\t'+sheet['Q'+row].v:null);
    cells.push(sheet['X'+row].v.toString().toLowerCase() == 'interno'?'Convenio INPEC:\n\tInterno':null);
    cells.push(sheet['AD'+row].v >= 43?'Edad:\n\t'+sheet['AD'+row].v+' años':null);
}

// Factor Psicosocial
function riesgoFPS(row, sheet, cells, na) {
    cells.push(na.indexOf(sheet['V'+row].v.toString().toLowerCase()) < 0?'Disc. Mental:\n\t'+sheet['V'+row].v:null);
    cells.push(na.indexOf(sheet['W'+row].v.toString().toLowerCase()) < 0?'Enfermedad:\n\t'+sheet['W'+row].v:null);
}

// Factor Académico Antecedentes
function riesgoFAA(row, sheet, cells, na) {
    cells.push(sheet['BI'+row].v.toString().toLowerCase() == 'no'?'Tomado Cursos Virtuales:\n\tNo':null);
    cells.push(sheet['BK'+row].v.toString().toLowerCase() == '5 años o mas'?'Tiempo sin Estudiar:\n\t5 años o mas':null);
    cells.push(sheet['BN'+row].v.toString().toLowerCase() == 'no'?'Primer opción de estudio:\n\tNo':null);
}

function prepareComment(cells) {
    var msg = '';
    cells.forEach(function(e){
        msg += (e!=null?'* '+e+'\n':'');
    });
    return msg;
}

function validaRiesgos(factor, row) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
    var na = ['no aplica', '#n/a', 'no pertenece', 'ninguna'];
    var cells = new Array;
    switch (factor) {
        case 'FSD':
            riesgoFSD(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["Riesgo por condiciones personales",msg]:["Sin riesgo por condiciones personales"];
            break;
        case 'FPS':
            riesgoFPS(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["Riesgo por situaciones personales",msg]:["Sin riesgo por situaciones personales"];
            break;
        case 'FAA':
            riesgoFAA(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["Riesgo por antecedentes académicos",msg]:["Sin Riesgo por antecedentes académicos"];
            break;
    }
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