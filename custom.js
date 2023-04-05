jQuery(document).ready(function($){
    var excelSheets = [];

    $("body").on("click", "#upload", function () {
        var fileUpload = $("#fileUpload")[0];
 
        $('.sheetNames').each(function(){
            excelSheets[$(this).val()] = $(this).next().next().val();
        })

        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
        if (regex.test(fileUpload.value.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();

                if (reader.readAsBinaryString) {
                    reader.onload = function (e) {
                        ProcessExcel(e.target.result);
                    };
                    reader.readAsBinaryString(fileUpload.files[0]);
                } else {
                    reader.onload = function (e) {
                        var data = "";
                        var bytes = new Uint8Array(e.target.result);
                        for (var i = 0; i < bytes.byteLength; i++) {
                            data += String.fromCharCode(bytes[i]);
                        }
                        ProcessExcel(data);
                    };
                    reader.readAsArrayBuffer(fileUpload.files[0]);
                }
            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            alert("Please upload a valid Excel file.");
        }
    });
    function ProcessExcel(data) {
        var workbook = XLSX.read(data, {
            type: 'binary'
        });
 
        var firstSheet = workbook.SheetNames[0];
        
        var table = $("<table id='export-table' />");
        table[0].border = "1";
 
        var row = $(table[0].insertRow(-1));
 
        var headerCell = $("<th />");
        headerCell.html("Key");
        row.append(headerCell);
 
        var headerCell = $("<th />");
        headerCell.html("Question");
        row.append(headerCell);
        
        for (var key in excelSheets) {

            var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[key]);

            var questionCols = excelSheets[key].split(',');
    
            for (var i = 0; i < excelRows.length; i++) {
                $.each(questionCols, function(index,value){
                    var row = $(table[0].insertRow(-1));
                    
                    var cell = $("<td />");
                    cell.html(excelRows[i].Key+"_"+value);
                    row.append(cell);

                    cell = $("<td />");
                    cell.html(excelRows[i]["Question" + value]);
                    row.append(cell);
                })
            }
        }
 
        var dvExcel = $("#dvExcel");
        dvExcel.html("");
        dvExcel.append(table);
    };

    $('#exportBtn').on("click", function() {
        var table = $('#export-table');

        var wb = XLSX.utils.table_to_book(table[0], {sheet:"Sheet1"});
        
        XLSX.writeFile(wb, "table.xlsx");
    });
})

$(document).ready(function() {
    var addButton = $('#add-text');
    var wrapper = $('#text-repeater');
    var fieldHTML = $('#text-repeater').html();

    addButton.click(function() {
        $(wrapper).append(fieldHTML);
    });

    $(wrapper).on('click', '.remove-text', function(e) {
        e.preventDefault();
        $(this).parent('div').remove();
    });

    $('#importBtn').click(function() {
		var xhr = new XMLHttpRequest();
		xhr.open('GET', 'testfile.xlsx', true);
		xhr.responseType = 'arraybuffer';
		xhr.onload = function(e) {
			if (xhr.status === 200) {
				var data = new Uint8Array(xhr.response);
				var workbook = XLSX.read(data, {type: 'array'});

				var sheetName = 'TestSheet';
				var worksheet = XLSX.utils.table_to_sheet($('#export-table')[0]);
				XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

				var wbout = XLSX.write(workbook, {bookType:'xlsx', type:'array'});
				var blob = new Blob([wbout], {type: 'application/octet-stream'});
				var url = URL.createObjectURL(blob);
				var a = document.createElement("a");
				a.href = url;
				a.download = 'testfile_new.xlsx';
				a.click();
				setTimeout(function() { URL.revokeObjectURL(url); }, 0);
			}
		};
		xhr.send();
	});
});