$(function () {

	var table = $('#table'),
		downloadButton = $('#btnDownload');

	downloadButton.on('click', function () {
        table.tableToExcel({
            filename: 'Lorem',
            preserveColors: true
        });
    });
});