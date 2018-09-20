function manageUndefined(cell) {
	var sheet = wb.Sheets['SocioDemograficos'];
	try {
		return sheet[cell].v;
	} catch(e) {
		return "";
	}
}

function cellStyle(info) {
    return (info!=''?" style='font-weight:bold; background-color:red;'":'');
}

function* range(start, end) {
    yield start;
    if (start === end) return;
    yield* range(start + 1, end);
}

function fillTable() {
    var sheet = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.decode_range(sheet["!ref"]).e.r;
    var cols = XLSX.utils.decode_range(sheet["!ref"]).e.c;
    var strRows = '';
    var risk;
    for (var i = 3; i <= rows; i++) {
        strRows += "<tr>";
        strRows += "<td>"+(i-2)+"</td>"; // Nro.
        strRows += "<td>"+manageUndefined('C'+(i+1))+"</td>"; // IDENTIFICACION DEL ESTUDIANTE
        strRows += "<td>"+manageUndefined('D'+(i+1))+"</td>"; // NOMBRE COMPLETO
        strRows += "<td>"+manageUndefined('R'+(i+1))+"</td>"; // CENTRO
        strRows += "<td>"+manageUndefined('Q'+(i+1))+"</td>"; // ZONA
        strRows += "<td>"+manageUndefined('T'+(i+1))+"</td>"; // PROGRAMA
        strRows += "<td>"+manageUndefined('S'+(i+1))+"</td>"; // ESCUELA
        strRows += "<td></td>"; // EMAIL
        strRows += "<td></td>"; // TELEFONO
        strRows += "<td>"+manageUndefined('F'+(i+1))+"</td>"; // ASIGNACION
        strRows += "<td>"+manageUndefined('A'+(i+1))+"</td>"; // TIPO
        risk = validaRiesgos('FSD', i+1, 0);
        riskInfo = addRisk(risk, 'L'+(i+2));
        strRows += "<td"+addComment(risk, 'L'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Socio-Demográfico
        risk = validaRiesgos('FPS', i+1, 0);
        riskInfo = addRisk(risk, 'M'+(i+2));
        strRows += "<td"+addComment(risk, 'M'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Psicosocial
        risk = validaRiesgos('FAA', i+1, 0);
        riskInfo = addRisk(risk, 'N'+(i+2));
        strRows += "<td"+addComment(risk, 'N'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Académico Antecedentes
        risk = validaRiesgos('FSE', i+1, 0);
        riskInfo = addRisk(risk, 'O'+(i+2));
        strRows += "<td"+addComment(risk, 'O'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Socio-Económico
        strRows += "<td></td>"; // Factores Internos (Indagar sobre las dificultades presentadas en la modalidad)
        strRows += "<td></td>"; // Factores Externos
        strRows += "<td></td>"; // ACCIONES DE INTERVENCION SEGÚN FACTORES DE RIESGO
        strRows += "<td></td>"; // Nivel de Riesgo Factores
        strRows += "<td></td>"; // COMPETENCIAS DIGITALES BASICAS
        strRows += "<td></td>"; // COMPETENCIAS CUANTITATIVAS
        strRows += "<td></td>"; // COMPETENCIAS LECTO-ESCRITORA
        strRows += "<td></td>"; // COMPETENCIAS DE INGLES
        strRows += "<td></td>"; // ACCIONES DE INTERVENCION SEGÚN RIESGOS COMPETENCIAS
        strRows += "<td></td>"; // Nivel de Riesgo Competencias
        strRows += "</tr>";
    }
    document.getElementById("tableBody").innerHTML = strRows;
}


/*
function fillTable() {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	var rows = XLSX.utils.decode_range(sheet["!ref"]).e.r;
	var cols = XLSX.utils.decode_range(sheet["!ref"]).e.c;
	var strRows = '';
	var risk;
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
            risk = validaRiesgos('FSD', i+1, 0);
            riskInfo = addRisk(risk, 'P'+(i+3));
    		strRows += "<td"+addComment(risk, 'P'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Socio-Demográfico
    		risk = validaRiesgos('FPS', i+1, 0);
            riskInfo = addRisk(risk, 'Q'+(i+3));
            strRows += "<td"+addComment(risk, 'Q'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Psicosocial
            risk = validaRiesgos('FAA', i+1, 0);
            riskInfo = addRisk(risk, 'R'+(i+3));
            strRows += "<td"+addComment(risk, 'R'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factor Académico Antecedentes
            risk = validaRiesgos('FSE', i+1, 0);
            strRows += "<td"+addComment(risk, 'S'+(i+3))+">"+risk[0]+addRisk(risk, 'S'+(i+3))+"</td>"; // Factor Socio-Económico
            // strRows += "<td></td>"; // Factor Socio-Económico
            // strRows += "<td></td>"; // Factor Académico por apropiacion al modelo
            strRows += "<td></td>"; // Factores Internos
            strRows += "<td></td>"; // Factores Externos
            risk = validaRiesgos('CDB', i+1, 1);
            riskInfo = addRisk(risk, 'U'+(i+3));
            strRows += "<td"+addComment(risk, 'U'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS DIGITALES BASICAS
            risk = validaRiesgos('CC', i+1, 1);
            riskInfo = addRisk(risk, 'V'+(i+3));
            strRows += "<td"+addComment(risk, 'V'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS CUANTITATIVAS
            risk = validaRiesgos('CLE', i+1, 1);
            riskInfo = addRisk(risk, 'W'+(i+3));
            strRows += "<td"+addComment(risk, 'W'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS LECTO-ESCRITORA
            risk = validaRiesgos('CI', i+1, 1);
            riskInfo = addRisk(risk, 'X'+(i+3));
            strRows += "<td"+addComment(risk, 'X'+(i+3))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS DE INGLES
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
*/

function addComment(comment, cellExport) {
	if (comment.length>1) {
		comments[cellExport] = {a:'VISAE', t:comment[1]};
		return " title='"+comment[1]+"'";
	} else {
		return '';
	}
}


function addRisk(risk, cellExport) {
    if (risk.length>1) {
        // comments[cellExport] = {a:'VISAE', t:comment[1]};
        return "<br>("+risk[1]+")";
    } else {
        return '';
    }   
}

// Analisis de los riesgos
// Factor Socio-Demográfico
function riesgoFSD(row, sheet, cells, na) {
    cells.push(sheet['F'+row].v >= 43?'Edad:\n\t'+sheet['F'+row].v+' años':null);
    cells.push(['Uno','Dos'].indexOf(sheet['L'+row].v.toString().toLowerCase())>=0?'Estrato:\n\t'+sheet['L'+row].v:null);
    cells.push(sheet['M'+row].v.toString().toLowerCase() == 'rural'?'Zona de Residencia:\n\t'+sheet['M'+row].v:null);
    cells.push(na.indexOf(sheet['V'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['V'+3].v:null);
    cells.push(na.indexOf(sheet['W'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['W'+3].v:null);
    cells.push(na.indexOf(sheet['X'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['X'+3].v:null);
    cells.push(na.indexOf(sheet['Y'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['Y'+3].v:null);
    cells.push(na.indexOf(sheet['Z'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['Z'+3].v:null);
    cells.push(na.indexOf(sheet['AA'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['AA'+3].v:null);
    cells.push(na.indexOf(sheet['AB'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['AB'+3].v:null);
    cells.push(na.indexOf(sheet['AC'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['AC'+3].v:null);
    cells.push(na.indexOf(sheet['AE'+row].v.toString().toLowerCase()) < 0?'Grupo Poblacional:\n\t'+sheet['AE'+3].v:null);
    cells.push(na.indexOf(sheet['AH'+row].v.toString().toLowerCase()) < 0?'Discapacidad:\n\t'+sheet['AH'+3].v:null);
    cells.push(na.indexOf(sheet['AI'+row].v.toString().toLowerCase()) < 0?'Discapacidad:\n\t'+sheet['AI'+3].v:null);
    cells.push(na.indexOf(sheet['AJ'+row].v.toString().toLowerCase()) < 0?'Discapacidad:\n\t'+sheet['AJ'+3].v:null);
    cells.push(na.indexOf(sheet['AK'+row].v.toString().toLowerCase()) < 0?'Discapacidad:\n\t'+sheet['AK'+3].v:null);
}

function noDiligenciados(cell, cells) {
    var field = manageUndefined(cell).toString().toLowerCase();
    if (field == '')
        cells.push('Aspecto no diligenciado');
    return field;
}

// Factor Psicosocial
function riesgoFPS(row, sheet, cells, na) {
    var field = noDiligenciados('EB'+row, cells);
    if (field != '') {
        var params = {'alto':[...range(10,16)], 'medio':[...range(17,23)], 'bajo':[...range(24,30)]};
        Object.keys(params).forEach(function(key) {
            if(params[key].includes(parseInt(field)))
                cells.push('Puntaje:\n\t'+key);
        });
    }
}

// Factor Académico Antecedentes
function riesgoFAA(row, sheet, cells, na) {
    var field = noDiligenciados('BA'+row, cells);
    cells.push(field != '' && field == 'más de 5 años'?'Tiempo sin Estudiar:\n\t'+sheet['BA'+row].v:null);
    field = noDiligenciados('BE'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Primera opción de estudio:\n\tNo':null);
    field = noDiligenciados('BN'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Uso de plataformas virtuales:\n\tNo':null);
    field = noDiligenciados('BO'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Uso de paquetes ofimáticos:\n\tNo':null);
    field = noDiligenciados('BP'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Participa en foros virtuales:\n\tNo':null);
    field = noDiligenciados('BQ'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Convierte archivos digitales:\n\tNo':null);
    field = noDiligenciados('BR'+row, cells);
    cells.push(field != '' && na.indexOf(field) >= 0?'Uso del correo electrónico:\n\t'+field:null);
}

// Factor Socio-Económico
function riesgoFSE(row, sheet, cells, na) {
    var field = noDiligenciados('BX'+row, cells);
    cells.push(field != '' && field == 'desempleado'?'Situación laboral:\n\t'+field:null);
    field = noDiligenciados('CC'+row, cells);
    cells.push(field != '' && ['ocasional o temporal','prestación de servicios'].includes(field)?'Tipo de contrato:\n\t'+field:null);
    field = noDiligenciados('CD'+row, cells);
    cells.push(field != '' && ['un salario mínimo','menos de un salario mínimo'].includes(field)?'Ingresos mensuales:\n\t'+field:null);
    field = noDiligenciados('CH'+row, cells);
    cells.push(field != '' && na.indexOf(field) < 0?'Debe buscar trabajo para continuar en la UNAD':null);
    field = noDiligenciados('AQ'+row, cells);
    cells.push(field != '' && na.indexOf(field) < 0?'Dependencia Económica':null);
    field = noDiligenciados('CI'+row, cells);
    cells.push(field != '' && ['recursos familiares','recursos empresariales','crédito financiero o icetex','subsidio o fondos gubernamentales'].includes(field)?'Origen de recursos:\n\t'+field:null);
}

// COMPETENCIAS DIGITALES BASICAS
function riesgoCDB(row, sheet, cells, na) {
    cells.push(sheet['AL'+row].v.toString().toLowerCase() == 'insuficiente'?'Nivel:\n\tInsuficiente':null);
}

// COMPETENCIAS CUANTITATIVAS
function riesgoCC(row, sheet, cells, na) {
    cells.push(sheet['AJ'+row].v.toString().toLowerCase() == 'insuficiente'?'Nivel:\n\tInsuficiente':null);
}

// COMPETENCIAS LECTO-ESCRITORA
function riesgoCLE(row, sheet, cells, na) {
    cells.push(sheet['AH'+row].v.toString().toLowerCase() == 'insuficiente'?'Nivel:\n\tInsuficiente':null);
}

// COMPETENCIAS DE INGLES
function riesgoCI(row, sheet, cells, na) {
    cells.push(sheet['AN'+row].v.toString().toLowerCase() == 'insuficiente'?'Nivel:\n\tInsuficiente':null);
}

function prepareComment(cells) {
    var msg = '';
    cells.forEach(function(e){
        msg += (e!=null?'* '+e+'\n':'');
    });
    return msg;
}

function validaRiesgos(factor, row, sheet) {
	var sheet = wb.Sheets[sheet==0?'SocioDemograficos':'competencias'];
    var na = ['no aplica', '#n/a', 'no pertenece', 'ninguna', 'n', 'ninguno', '', 'no usa correo electrónico'];
    var cells = new Array;
    switch (factor) {
        case 'FSD':
            riesgoFSD(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo por condiciones personales"];
            break;
        case 'FPS':
            riesgoFPS(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo por situaciones personales"];
            break;
        case 'FAA':
            riesgoFAA(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin Riesgo por antecedentes académicos"];
            break;
        case 'FSE':
            riesgoFSE(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo socioeconómico"];
            break;
        case 'CDB':
            riesgoCDB(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo en Competencias Digitales Basicas"];
            break;
        case 'CC':
            riesgoCC(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo en Competencias Cuantitativas"];
            break;
        case 'CLE':
            riesgoCLE(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo en Competencias Lecto-Escritora"];
            break;
        case 'CI':
            riesgoCI(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo en Competencias de Ingles"];
            break;
    }
}