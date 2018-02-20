<html>
<head>
	<title></title>
	<meta charset="utf-8">
  	<meta name="viewport" content="width=device-width, initial-scale=1">
  	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  	<link rel="stylesheet" type="text/css" href="css/style.css">
  	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
  	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
  	<script type="text/javascript" src="js/xlsx.full.min.js"></script>
  	<script type="text/javascript" src="js/FileSaver.min.js"></script>
  	<script type="text/javascript" src="js/exportData.js"></script>  	
  	<script type="text/javascript" src="js/loadData.js"></script>
  	<script type="text/javascript" src="js/fillTable.js"></script>
  	<script type="text/javascript" src="js/tableHeadFixer.js"></script>
  	<script>
		$(document).ready(function() {
			$("#matriz").tableHeadFixer();
		});
	</script>
</head>
<body>
<div id="wrapper">
	<input type="file" id="input-excel">
	<button id="button-a">Exportar</button>
	<label id="msg"></label>
	<div id="parent" class="table-responsive">
	<!--<table class="table table-hover" id="matriz">
		<thead>
			<tr>
				<th>IDENTIFICACION DEL ESTUDIANTE</th>
				<th>NOMBRE COMPLETO</th>
				<th>CENTRO</th>
				<th>ZONA</th>
				<th>PROGRAMA</th>
				<th>ESCUELA</th>
				<th>EMAIL </th>
				<th>TELEFONO</th>
				<th>EDAD</th>
				<th>ASIGNACION</th>
				<th>CONVENIO</th>
				<th>NOVEDAD</th>
				<th>GENERAL</th>
				<th>CAMPUS VIRTUAL</th>
				<th>Riesgo </th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>Riesgo</th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>Riesgo</th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>Riesgo</th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>Riesgo</th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>Riesgo</th>
				<th>Acciones realizadas según ruta de PAPC</th>
				<th>RESULTADOS</th>
				<th>NIVEL DE RIESGO POR FACTORES</th>
				<th>GRUPO COLABORATIVO</th>
				<th>SITUACION DE RIESGO</th>
				<th>ACCIONES REALIZADAS</th>
				<th>RESULTADOS</th>
				<th>No. Cursos </th>
				<th>SITUACION DE RIESGO</th>
				<th>ACCIONES REALIZADAS</th>
				<th>RESULTADOS</th>
			</tr>
		</thead>
		<tbody id="tableBody">

		</tbody>
	</table>-->
        <table class="table table-bordered table-hover" id="matriz">
            <thead>
                <tr>
                    <th colspan="62"></th>
                </tr>
                <tr>
                    <th class="text-center" colspan="12" rowspan="3" bgcolor="#7fff00">INFORMACION BASICA DEL ESTUDIANTE</th>
                    <th class="text-center" rowspan="4" bgcolor="#696969">NOVEDAD</th>
                    <th class="text-center" colspan="22" bgcolor="#ff8c00">PRIMER ACOMPAÑAMIENTO</th>
                    <th class="text-center" colspan="13" bgcolor="yellow">SEGUNDO ACOMPAÑAMIENTO</th>
                    <th class="text-center" colspan="8" bgcolor="blue">TERCER ACOMPAÑAMIENTO</th>
                    <th class="text-center" colspan="5" rowspan="2" bgcolor="#ff8c00">RESULTADO FINAL</th>
                    <th class="text-center" rowspan="4">CONSEJERO ACADEMICO</th>
                </tr>
                <tr>
                    <th class="text-center" colspan="9" bgcolor="#ff8c00">FACTORES DE RIESGO</th>
                    <th class="text-center" colspan="4" bgcolor="yellow">RIESGO POR COMPETENCIAS BASICAS</th>
                    <th class="text-center" rowspan="3" bgcolor="olive">MEDIO DE CONTACTO</th>
                    <th class="text-center" rowspan="3" bgcolor="#556b2f">ACCIONES REALIZADAS SEGÚN RUTA DE PAPC</th>
                    <th class="text-center" rowspan="3" bgcolor="#556b2f">RESULTADOS ACCIONES DE ACOMPAÑAMIENTO</th>
                    <th class="text-center" rowspan="3" bgcolor="#556b2f">NIVEL DE RIESGO</th>
                    <th class="text-center" rowspan="3" bgcolor="#556b2f">OBSERVACIONES</th>
                    <th class="text-center" colspan="4" bgcolor="#ff8c00">RIESGO EN CATEDRA UNADISTA</th>
                    <th class="text-center" colspan="9" bgcolor="yellow">RIESGO POR ALERTAS EN CURSOS MATRICULADOS</th>
                    <th class="text-center" colspan="4" bgcolor="yellow">RIESGO POR COMPETENCIAS BASICAS</th>
                    <th class="text-center" colspan="8" bgcolor="blue">RIESGO POR ALERTAS EN CURSOS MATRICULADOS</th>
                </tr>
                <tr>
                    <th class="text-center" colspan="2" bgcolor="#ff7f50">ACOGIDA E INTEGRACION</th>
                    <th class="text-center" rowspan="2" bgcolor="olive">Factor Socio-Demográfico</th>
                    <th class="text-center" rowspan="2" bgcolor="olive">Factor Psicosocial</th>
                    <th class="text-center" rowspan="2" bgcolor="olive">Factor Académico Antecedentes</th>
                    <th class="text-center" rowspan="2" bgcolor="#add8e6">Factor Socio-Económico</th>
                    <th class="text-center" rowspan="2" bgcolor="#f08080">Factor Académico por apropiacion al modelo</th>
                    <th class="text-center" rowspan="2" bgcolor="#e0ffff">Factor Institucional</th>
                    <th class="text-center" rowspan="2" bgcolor="#fafad2">Factores Externos</th>
                    <th class="text-center" colspan="4" bgcolor="olive">TALLERES DE NIVELACION DE COMPETENCIAS</th>
                    <th class="text-center" colspan="4" bgcolor="#00ffff">CATEDRA UNADISTA</th>
                    <th class="text-center" colspan="5" bgcolor="#9370db">RENDIMIENTO ACADEMICO CURSOS MATRICULADOS</th>
                    <th class="text-center" colspan="4" bgcolor="#00ffff">RIESGO CATEDRA UNADISTA</th>
                    <th class="text-center" colspan="4" bgcolor="#696969">TALLERES DE NIVELACION DE COMPETENCIAS</th>
                    <th class="text-center" colspan="4" bgcolor="#6495ed">RENDIMIENTO ACADEMICO</th>
                    <th class="text-center" colspan="4" bgcolor="#00ffff">CATEDRA UNADISTA</th>
                    <th class="text-center" rowspan="2" bgcolor="#6b8e23">MEDIO DE CONTACTO</th>
                    <th class="text-center" rowspan="2" bgcolor="#6b8e23">APROBACION DE LA CATEDRA</th>
                    <th class="text-center" rowspan="2" bgcolor="#6b8e23">APROBACION DE CURSOS</th>
                    <th class="text-center" rowspan="2" bgcolor="#6b8e23">CULMINACION DEL PERIODO ACADEMICO</th>
                    <th class="text-center" rowspan="2" bgcolor="#6b8e23">RESULTADO FINAL</th>
                </tr>
                <tr>
                    <th class="text-center" bgcolor="gray">Nro.</th>
                    <th class="text-center" bgcolor="gray">IDENTIFICACION DEL ESTUDIANTE</th>
                    <th class="text-center" bgcolor="gray">NOMBRE COMPLETO</th>
                    <th class="text-center" bgcolor="gray">CENTRO</th>
                    <th class="text-center" bgcolor="gray">ZONA</th>
                    <th class="text-center" bgcolor="gray">PROGRAMA</th>
                    <th class="text-center" bgcolor="gray">ESCUELA</th>
                    <th class="text-center" bgcolor="gray">EMAIL</th>
                    <th class="text-center" bgcolor="gray">TELEFONO</th>
                    <th class="text-center" bgcolor="gray">EDAD</th>
                    <th class="text-center" bgcolor="gray">ASIGNACION</th>
                    <th class="text-center" bgcolor="gray">CONVENIO</th>
                    <th class="text-center" bgcolor="#ff7f50">PARTICIPACION INDUCCION</th>
                    <th class="text-center" bgcolor="#ff7f50">PARTICIPACION INMERSION CAMPUS VIRTUAL</th>
                    <th class="text-center" bgcolor="olive">COMPETENCIAS DIGITALES BASICAS</th>
                    <th class="text-center" bgcolor="olive">COMPETENCIAS CUANTITATIVAS</th>
                    <th class="text-center" bgcolor="olive">COMPETENCIAS LECTO-ESCRITORA</th>
                    <th class="text-center" bgcolor="olive">COMPETENCIAS DE INGLES</th>
                    <th class="text-center" bgcolor="#00ffff">GRUPO COLABORATIVO</th>
                    <th class="text-center" bgcolor="#00ffff">SITUACION DE RIESGO</th>
                    <th class="text-center" bgcolor="#00ffff">ACCIONES REALIZADAS</th>
                    <th class="text-center" bgcolor="#00ffff">RESULTADOS</th>
                    <th class="text-center" bgcolor="#9370db">No. Cursos</th>
                    <th class="text-center" bgcolor="#9370db">SITUACION DE RIESGO</th>
                    <th class="text-center" bgcolor="#9370db">MEDIO DE CONTACTO</th>
                    <th class="text-center" bgcolor="#9370db">ACCIONES REALIZADAS</th>
                    <th class="text-center" bgcolor="#9370db">RESULTADOS</th>
                    <th class="text-center" bgcolor="#00ffff">SITUACION DE RIESGO</th>
                    <th class="text-center" bgcolor="#00ffff">ACCIONES REALIZADAS</th>
                    <th class="text-center" bgcolor="#00ffff">RESULTADOS</th>
                    <th class="text-center" bgcolor="#00ffff">SITUACION DE ALERTA</th>
                    <th class="text-center" bgcolor="#696969">COMPETENCIAS DIGITALES BASICAS</th>
                    <th class="text-center" bgcolor="#696969">COMPETENCIAS CUANTITATIVAS</th>
                    <th class="text-center" bgcolor="#696969">COMPETENCIAS LECTO-ESCRITORA</th>
                    <th class="text-center" bgcolor="#696969">COMPETENCIAS DE INGLES</th>
                    <th class="text-center" bgcolor="#6495ed">SITUACION DE RIESGO</th>
                    <th class="text-center" bgcolor="#6495ed">MEDIO DE CONTACTO</th>
                    <th class="text-center" bgcolor="#6495ed">ACCIONES REALIZADAS</th>
                    <th class="text-center" bgcolor="#6495ed">RESULTADOS</th>
                    <th class="text-center" bgcolor="#00ffff">SITUACION DE RIESGO</th>
                    <th class="text-center" bgcolor="#00ffff">ACCIONES REALIZADAS</th>
                    <th class="text-center" bgcolor="#00ffff">RESULTADOS</th>
                    <th class="text-center" bgcolor="#00ffff">SITUACION DE ALERTA</th>
                </tr>
            </thead>
            <tbody id="tableBody">
                <tr>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                    <td class="tg-031e"></td>
                </tr>
            </tbody>
        </table>
	</div>
</div>
</body>
</html>