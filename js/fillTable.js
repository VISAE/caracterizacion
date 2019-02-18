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
        risk = validaRiesgos('FE', i+1, 0);
        riskInfo = addRisk(risk, 'P'+(i+2));
        strRows += "<td"+addComment(risk, 'P'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factores Externos
        risk = validaRiesgos('FI', i+1, 0);
        riskInfo = addRisk(risk, 'Q'+(i+2));
        strRows += "<td"+addComment(risk, 'Q'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // Factores Internos (Indagar sobre las dificultades presentadas en la modalidad)
        strRows += "<td></td>"; // ACCIONES DE INTERVENCION SEGÚN FACTORES DE RIESGO
        strRows += "<td></td>"; // Nivel de Riesgo Factores
        risk = validaRiesgos('CDB', i+1, 0);
        riskInfo = addRisk(risk, 'T'+(i+2));
        strRows += "<td"+addComment(risk, 'T'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS DIGITALES BASICAS
        risk = validaRiesgos('CC', i+1, 0);
        riskInfo = addRisk(risk, 'U'+(i+2));
        strRows += "<td"+addComment(risk, 'U'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS CUANTITATIVAS
        risk = validaRiesgos('CLE', i+1, 0);
        riskInfo = addRisk(risk, 'V'+(i+2));
        strRows += "<td"+addComment(risk, 'V'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS LECTO-ESCRITORA
        risk = validaRiesgos('CI', i+1, 0);
        riskInfo = addRisk(risk, 'W'+(i+2));
        strRows += "<td"+addComment(risk, 'W'+(i+2))+cellStyle(riskInfo)+">"+risk[0]+riskInfo+"</td>"; // COMPETENCIAS DE INGLES
        strRows += "<td></td>"; // ACCIONES DE INTERVENCION SEGÚN RIESGOS COMPETENCIAS
        strRows += "<td></td>"; // Nivel de Riesgo Competencias
        strRows += "</tr>";
    }
    document.getElementById("tableBody").innerHTML = strRows;
}

function manageUndefined(cell) {
    var sheet = wb.Sheets['SocioDemograficos'];
    try {
        return sheet[cell].v;
    } catch(e) {
        return "";
    }
}

function cellStyle(info) {
    var expr = /(Riesgo:\n\t\b(medio|bajo)\b.*)/i;
    return (info!='' && !info.match(expr)?" style='font-weight:bold; background-color:red;'":'');
}

function* range(start, end) {
    yield start;
    if (start === end) return;
    yield* range(start + 1, end);
}

function addComment(comment, cellExport) {
	if (comment.length>1) {
		comments[cellExport] = {a:'VISAE', t:comment[1]};
		return " title='["+comment[1]+"]'";
	} else {
		return '';
	}
}


function addRisk(risk, cellExport) {
    if (risk.length>1) {
        // comments[cellExport] = {a:'VISAE', t:comment[1]};
        return "["+risk[1].replace(/\]\n\[/g,"]&diams;[")+"]";
    } else {
        return '';
    }   
}

// Analisis de los riesgos

function verificaDiligenciados(fields) {
    fieldsD = {};
    fields.forEach(function (field) {
        fieldsD[field] = manageUndefined(field).toString().toLowerCase();
    });
    return fieldsD;
}

function calculaPuntaje(cell, cells, ranges) {
    var fields = verificaDiligenciados([cell]);
    if(Object.values(fields).join('') == '')
        cells.push('Aspecto no diligenciado &diams; ');
    else {
        var params = {'alto': [...range(ranges[0], ranges[1])],'medio':[...range(ranges[2], ranges[3])],'bajo':[...range(ranges[4], ranges[5])]};
        Object.keys(fields).forEach(function (key) {
            switch (key) {
                case (cell):
                    Object.keys(params).forEach(function(keyp) {
                        if(params[keyp].includes(parseInt(fields[key])))
                            cells.push('Riesgo:\n\t'+keyp+' &diams; ');
                    });
                    break;
            }
        });
    }
}

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


// Factor Psicosocial
function riesgoFPS(row, sheet, cells, na) {
    calculaPuntaje('EB'+row,cells,[10, 16, 17, 23, 24, 30]);
}

// Factor Académico Antecedentes
function riesgoFAA(row, sheet, cells, na) {
    var fields = verificaDiligenciados(['BA'+row,'BE'+row,'BN'+row,'BO'+row,'BP'+row,'BQ'+row,'BR'+row]);
    if(Object.values(fields).join('') == '')
        cells.push('Aspecto no diligenciado');
    else {
        Object.keys(fields).forEach(function (key) {
            switch (key) {
                case('BA' + row):
                    cells.push(fields[key] != '' && fields[key] == 'más de 5 años' ? 'Tiempo sin Estudiar:\n\t' + fields[key] : null);
                    break;
                case('BE' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Primera opción de estudio:\n\tNo' : null);
                    break;
                case('BN' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Uso de plataformas virtuales:\n\tNo' : null);
                    break;
                case('BO' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Uso de paquetes ofimáticos:\n\tNo' : null);
                    break;
                case('BP' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Participa en foros virtuales:\n\tNo' : null);
                    break;
                case('BQ' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Convierte archivos digitales:\n\tNo' : null);
                    break;
                case('BR' + row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) >= 0 ? 'Uso del correo electrónico:\n\t' + fields[key] : null);
                    break;
            }
        });
    }
}

// Factor Socio-Económico
function riesgoFSE(row, sheet, cells, na) {
    var fields = verificaDiligenciados(['BX'+row,'CC'+row,'CD'+row,'CH'+row,'AQ'+row,'CI'+row]);
    if(Object.values(fields).join('') == '')
        cells.push('Aspecto no diligenciado');
    else {
        Object.keys(fields).forEach(function (key) {
            switch (key) {
                case('BX'+row):
                    cells.push(fields[key] != '' && fields[key] == 'desempleado'?'Situación laboral:\n\t'+fields[key]:null);
                    break;
                case('CC'+row):
                    cells.push(fields[key] != '' && ['ocasional o temporal','prestación de servicios'].includes(fields[key])?'Tipo de contrato:\n\t'+fields[key]:null);
                    break;
                case('CD'+row):
                    cells.push(fields[key] != '' && ['un salario mínimo','menos de un salario mínimo'].includes(fields[key])?'Ingresos mensuales:\n\t'+fields[key]:null);
                    break;
                case('CH'+row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) < 0?'Debe buscar trabajo para continuar en la UNAD:\n\tSi':null);
                    break;
                case('AQ'+row):
                    cells.push(fields[key] != '' && na.indexOf(fields[key]) < 0?'Dependencia Económica:\n\tSi':null);
                    break;
                case('CI'+row):
                    cells.push(fields[key] != '' && ['recursos familiares','recursos empresariales','crédito financiero o icetex','subsidio o fondos gubernamentales'].includes(fields[key])?'Origen de recursos:\n\t'+fields[key]:null);
                    break;
            }
        });
    }
}

// Factores Externos
function riesgoFE(row, sheet, cells, na) {
    var fields = verificaDiligenciados(['BL'+row,'BM'+row]);
    if(Object.values(fields).join('') == '')
        cells.push('Aspecto no diligenciado');
    else {
        Object.keys(fields).forEach(function (key) {
            switch (key) {
                case('BL'+row):
                    cells.push(fields[key] != '' && ['intermitente (servicio interrumpido programado al menos dos veces por semana)','escaso (con servicio disponible 1 o 2 días por semana)'].includes(fields[key])?'Servicio de energía:\n\t'+fields[key]:null);
                    break;
                case('BM'+row):
                    cells.push(fields[key] != '' && ['intermitente (servicio interrumpido)','no cuenta con el servicio (con servicio disponible 1 o 2 días por semana)'].includes(fields[key])?'Servicio de Internet:\n\t'+fields[key]:null);
                    break;
            }
        });
    }
}

// Factores Internos
function riesgoFI(row, sheet, cells, na) {
    var fields = verificaDiligenciados(['BH'+row,'BI'+row,'BJ'+row,'BK'+row,'CE'+row]);
    if(Object.values(fields).join('') == '')
        cells.push('Aspecto no diligenciado');
    else {
        var pc = false;
        Object.keys(fields).forEach(function (key) {
            switch (key) {
                case('BH'+row):
                case('BI'+row):
                    pc = pc || (fields[key] != '' && na.indexOf(fields[key]) < 0);
                    break;
                case('BJ'+row):
                case('BK'+row):
                    if(!pc)
                        cells.push(fields[key] != '' && na.indexOf(fields[key]) < 0?'Equipo electrónico para actividades académicas:\n\t'+sheet[key.slice(0,2)+3].v:null);
                    break;
                case('CE'+row):
                    cells.push(fields[key] != '' && fields[key] == '1 a 4 horas por semana'?'Tiempo para actividades académicas:\n\t'+fields[key]:null);
                    break;
            }
        });
    }
}

// COMPETENCIAS DIGITALES BASICAS
function riesgoCDB(row, sheet, cells, na) {
    calculaPuntaje('EC'+row, cells, [0, 40, 50, 80, 90, 120]);
}

// COMPETENCIAS CUANTITATIVAS
function riesgoCC(row, sheet, cells, na) {
    calculaPuntaje('EE'+row, cells, [0, 30, 40, 60, 70, 100]);
}

// COMPETENCIAS LECTO-ESCRITORA
function riesgoCLE(row, sheet, cells, na) {
    calculaPuntaje('ED'+row, cells, [0, 40, 50, 90, 100, 150]);
}

// COMPETENCIAS DE INGLES
function riesgoCI(row, sheet, cells, na) {
    calculaPuntaje('EF'+row, cells, [0, 40, 50, 80, 90, 120]);
}

function prepareComment(cells) {
    var msg = cells.filter(x => x).join(']\n[');
    /*cells.forEach(function(e){
        msg += (e!=null?'* '+e+'\n':'');
    });*/
    return msg;
}

function validaRiesgos(factor, row, sheet) {
	var sheet = wb.Sheets[sheet==0?'SocioDemograficos':'competencias'];
    var na = ['no aplica', '#n/a', 'no pertenece', 'ninguna', 'n', 'ninguno', 'no usa correo electrónico'];
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
            return msg.length > 0?["",msg + 'Puntaje: '+ manageUndefined('EB'+row)]:["Sin riesgo por situaciones personales"];
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
        case 'FE':
            riesgoFE(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo por factores externos"];
            break;
        case 'FI':
            riesgoFI(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg]:["Sin riesgo por factores internos"];
            break;
        case 'CDB':
            riesgoCDB(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg + 'Puntaje: '+ manageUndefined('EC'+row)]:["Sin riesgo en Competencias Digitales Basicas"];
            break;
        case 'CC':
            riesgoCC(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg + 'Puntaje: '+ manageUndefined('EE'+row)]:["Sin riesgo en Competencias Cuantitativas"];
            break;
        case 'CLE':
            riesgoCLE(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg + 'Puntaje: '+ manageUndefined('ED'+row)]:["Sin riesgo en Competencias Lecto-Escritora"];
            break;
        case 'CI':
            riesgoCI(row, sheet, cells, na);
            var msg = prepareComment(cells);
            return msg.length > 0?["",msg + 'Puntaje: '+ manageUndefined('EF'+row)]:["Sin riesgo en Competencias de Ingles"];
            break;
    }
}