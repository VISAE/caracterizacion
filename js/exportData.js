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
        // console.log(property);
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
