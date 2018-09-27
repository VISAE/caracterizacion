<html>
<head>
	<title></title>
    <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
    <!--<meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">-->
  	<link rel="stylesheet" href="css/bootstrap.min.css">
  	<link rel="stylesheet" type="text/css" href="css/style.css">
    <link rel="stylesheet" type="text/css" href="css/tableexport.css">
  	<script src="js/jquery-3.3.1.min.js"></script>
  	<script src="js/bootstrap.min.js"></script>
  	<!--<script type="text/javascript" src="js/xlsx.full.min.js"></script>
  	<script type="text/javascript" src="js/FileSaver.min.js"></script>-->
    <script type="text/javascript" src="js/spin.js"></script>
    <script type="text/javascript" src="js/tableHeadFixer.js"></script>
    <script type="text/javascript" src="js/xlsx.core.min.js"></script>
    <script type="text/javascript" src="js/FileSaver.js"></script>
    <script type="text/javascript" src="js/tableexport.js"></script>
  	<script type="text/javascript" src="js/loadData.js"></script>
    <script type="text/javascript" src="js/fillTable.js"></script>
    <script type="text/javascript" src="js/exportData.js"></script>

</head>
<body>
<div id="wrapper">
	<input type="file" id="input-excel" class="label-info" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" >
    <input type="button" onclick="tableToExcel('matriz', 'Export HTML Table to Excel')" value="Export to Excel" />

    <!--	<button id="button-a" class="btn-primary">Exportar</button>-->

    <!--Spinner-->
    <div id="loading">
        <div id="loadingcont">
            <p id="loadingspinr">
            </p>
        </div>
    </div>

	<label id="msg"></label>
	<div id="parent" class="table-responsive">
        <table class="table table-bordered table-hover" id="matriz">
            <thead>
                <tr>
                    <th colspan="25"></th>
                </tr>
                <tr>
                    <th class="text-center" colspan="11" rowspan="2" bgcolor="#666666">INFORMACION BASICA DEL ESTUDIANTE</th>
                    <th class="text-center" colspan="14" bgcolor="#ffff00">SEGUNDO ACOMPAÑAMIENTO</th>
                </tr>
                <tr>
                    <th class="text-center" colspan="8" bgcolor="#bf9000">FACTORES DE RIESGO</th>
                    <th class="text-center" colspan="6" bgcolor="f1c232">RIESGO POR COMPETENCIAS</th>
                </tr>
                <tr>
                    <th class="text-center" bgcolor="#666666">Nro.</th>
                    <th class="text-center" bgcolor="#666666">IDENTIFICACION DEL ESTUDIANTE</th>
                    <th class="text-center" bgcolor="#666666">NOMBRE COMPLETO</th>
                    <th class="text-center" bgcolor="#666666">CENTRO</th>
                    <th class="text-center" bgcolor="#666666">ZONA</th>
                    <th class="text-center" bgcolor="#666666">PROGRAMA</th>
                    <th class="text-center" bgcolor="#666666">ESCUELA</th>
                    <th class="text-center" bgcolor="#666666">EMAIL</th>
                    <th class="text-center" bgcolor="#666666">TELEFONO</th>
                    <th class="text-center" bgcolor="#666666">ASIGNACION</th>
                    <th class="text-center" bgcolor="#666666">TIPO</th>
                    <th class="text-center" bgcolor="#bf9000">Factor Socio-Demográfico</th>
                    <th class="text-center" bgcolor="#bf9000">Factor Psicosocial</th>
                    <th class="text-center" bgcolor="#bf9000">Factor Académico Antecedentes</th>
                    <th class="text-center" bgcolor="#bf9000">Factor Socio-Económico</th>
                    <th class="text-center" bgcolor="#bf9000">Factores Externos</th>
                    <th class="text-center" bgcolor="#bf9000">Factores Internos (Indagar sobre las dificultades presentadas en la modalidad)</th>
                    <th class="text-center" bgcolor="#bf9000">ACCIONES DE INTERVENCION SEGÚN FACTORES DE RIESGO</th>
                    <th class="text-center" bgcolor="#ff9900">Nivel de Riesgo Factores</th>
                    <th class="text-center" bgcolor="#f1c232">COMPETENCIAS DIGITALES BASICAS</th>
                    <th class="text-center" bgcolor="#f1c232">COMPETENCIAS CUANTITATIVAS</th>
                    <th class="text-center" bgcolor="#f1c232">COMPETENCIAS LECTO-ESCRITORA</th>
                    <th class="text-center" bgcolor="#f1c232">COMPETENCIAS INGLES</th>
                    <th class="text-center" bgcolor="#f1c232">ACCIONES DE INTERVENCION SEGÚN RIESGOS COMPETENCIAS</th>
                    <th class="text-center" bgcolor="#ff9900">Nivel de Riesgo Competencias</th>
                </tr>
            </thead>
            <tbody id="tableBody">

            </tbody>
        </table>
	</div>
</div>
</body>
</html>