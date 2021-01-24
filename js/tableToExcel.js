(function ($, window, document, undefined) {
    var pluginName = "tableToExcel",
        preserveColors = false,
        utf8Heading = "<meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\">",
        template = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
            "xmlns:x=\"urn:schemas-microsoft-com:office:excel\" " +
            "xmlns=\"http://www.w3.org/TR/REC-html40\">" + utf8Heading +
            "<body>" +
            "<table> {table} </table>" +
            "</body>" +
            "</html>",
        getComputedStyleByElement = function (el) {
            var additionalStyles = '';
            if (preserveColors) {
                var compStyle = getComputedStyle(el);
                additionalStyles += (compStyle && compStyle.backgroundColor ? "background-color: " + compStyle.backgroundColor + ";" : "");
                additionalStyles += (compStyle && compStyle.color ? "color: " + compStyle.color + ";" : "");
            }
            return additionalStyles;
        },
        createRows = function (rootEl) {
            var tempRows = '';

            $(rootEl).find("tr").each(function () {
                var row = this,
                    additionalStyles = getComputedStyleByElement(row);

                tempRows += '<tr style="' + additionalStyles + '">' + createColumns(row) + '</tr>';
            });

            return tempRows;
        },
        createColumns = function (rootEl) {
            var tempColumn = '';

            $(rootEl).find('td,th').each(function () {
                var column = this,
                    additionalStyles = getComputedStyleByElement(column);

                tempColumn += '<td style="' + (additionalStyles ? additionalStyles : '') + '">' + $(column).html() + '</td>';
            });

            return tempColumn;
        },
        format = function (s, c) {
            return s.replace(/{(\w+)}/g, function (m, p) {
                return c[p];
            });
        },
        downloadExcel = function (blob, fileName) {
            window.URL = window.URL || window.webkitURL;

            var link = window.URL.createObjectURL(blob),
                a = document.createElement("a"),
                fileExt = '.xls';

            a.download = fileName + fileExt;
            a.href = link;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        },
        prepareTable = function (tableEl) {
            return createRows(tableEl);
        },
        tableToExcel = function (table, options) {
            var blob = new Blob([format(template, {table: table})], {type: "application/vnd.ms-excel"});

            downloadExcel(blob, options.filename);
        }

    $.fn[pluginName] = function (options) {
        preserveColors = options.preserveColors;

        var table = prepareTable(this);

        tableToExcel(table, options);
    };

})(jQuery, window, document);
