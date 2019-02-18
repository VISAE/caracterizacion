var tableToExcel = (function () {
    var uri = 'data:application/vnd.ms-excel;base64,'
        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
    return function (table, name) {
        if (!table.nodeType) table = document.getElementById(table)
        var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML }
        var blob = new Blob([format(template, ctx)]);

        window.saveAs(blob, 'reporte.xls');
        /*var blobURL = window.URL.createObjectURL(blob);

        if (ifIE()) {
            csvData = table.innerHTML;
            if (window.navigator.msSaveBlob) {
                var blob = new Blob([format(template, ctx)], {
                    type: "text/html"
                });
                navigator.msSaveBlob(blob, '' + name + '.xls');
            }
        }
        else
            window.location.href = uri + base64(format(template, ctx))*/
    }
})()

function ifIE() {
    var isIE11 = navigator.userAgent.indexOf(".NET CLR") > -1;
    var isIE11orLess = isIE11 || navigator.appVersion.indexOf("MSIE") != -1;
    return isIE11orLess;
}
/*$(document).ready(function (e) {
    TableExport.prototype.formatConfig.xlsx.buttonContent = "Exportar a XLSX";
    TableExport.prototype.formatConfig.xls.buttonContent = "Exportar a XLS";
    $('#button-a').click(function () {
        $("#matriz").tableExport({
            filename: 'Resultado',
            position: 'top',
            formats: ['xls', 'xlsx'],
            bootstrap: true
        });
    });
});*/

/*
$(document).ready(function (e) {
    $('#button-a').click(function () {
        if (wb !== undefined) {
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
            trget.appendChild(spnr.el);
            //

            wb = XLSX.utils.table_to_book(document.getElementById('matriz'), {sheet: "Resultados"});

            fillWithComments();
            var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: true, type: 'binary'});
            saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), "Resultado.xlsx");

            setTimeout(function() {
                spnr.stop(); // Stop the spinner
            }, 4000);
        }
    });
});

function fillWithComments() {
    var sheet = wb.Sheets[wb.SheetNames[0]];
    for (var property in comments) {
        // console.log(comments[property]);
        if (comments.hasOwnProperty(property)) {
            sheet[property].c = [];
            sheet[property].c.push(comments[property]);            
        }
    }
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
*/
