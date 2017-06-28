/*
Author: Jyoti Rani
*/

 $(function() {
	var saveDataInterval = 30000; // save data after every half minute 
	reconfigureRow(); // make existing tabel resizable
	reconfigureCol();
	//fetchData(); // fetch the existing saved data from local storage
	
	/*
		save data in Local Storage
	*/
	$('#saveData').click(function () {
		saveData();
	});
	function saveData(){
		if (typeof(Storage) !== "undefined") {
			// Store
			console.log("saving");
			//console.log($("#dataTable")[0].innerHTML);
			localStorage.setItem("dataTable",$("#dataTable")[0].innerHTML);
			$("#savedAt")[0].innerHTML = "<b>Last Saved:" + new Date(); +"</b>"
		} else {
			console.log("Sorry, your browser does not support Web Storage...");
		}
	}
	setInterval(function(){saveData(); }, saveDataInterval);
	$('#json').click(function () {
		
		jsonFile = new Blob([tableToJson()], {type: "text/json"});
		downloadLink = document.createElement("a");

		// File name
		downloadLink.download = "tabelData.json";

		// Create a link to the file
		downloadLink.href = window.URL.createObjectURL(jsonFile);

		// Hide download link
		downloadLink.style.display = "none";

		// Add the link to DOM
		document.body.appendChild(downloadLink);
		// Click download link
		downloadLink.click();

	});
	$('#csv').click(function () {
		
		csvFile = new Blob([tabelToCsv()], {type: "text/csv"});
		downloadLink = document.createElement("a");

		// File name
		downloadLink.download = "tabelData.csv";

		// Create a link to the file
		downloadLink.href = window.URL.createObjectURL(csvFile);

		// Hide download link
		downloadLink.style.display = "none";

		// Add the link to DOM
		document.body.appendChild(downloadLink);

		// Click download link
		downloadLink.click();
	});
	
	$(document).on("click", "#copy" , function() {
		copyData();
	});
	
	$(document).on("click", "#paste" , function() {
		pasteData();
	});
	$(document).on("click", ".removeRow" , function() {
		var currRow = $(this).parent().parent().parent().parent()[0];
		var rowIndex = currRow.rowIndex;
		//console.log(rowIndex);
		//console.log(currRow.filter("td:eq("+rowIndex+")"));
		currRow.remove();
		reconfigureRow();
	});
	
	$(document).on("click", ".addRow" , function() {
		var rowIndex = $(this).parent().parent().parent().parent()[0].rowIndex;
		var firstRow = $("#excel-table").find('tr:first-child th');
		var row = $("#excel-table")[0].insertRow(rowIndex+1);
		
		firstRow.each(function(index){
			var cell = row.insertCell(index);
			
			if(index == 0){
				cell.innerHTML = '<span><p>'+ (+rowIndex+1)+'</p><div class="dropdown"><i title ="Add Row"class="fa fa-plus-circle fa-lg addRow"></i><i title = "Remove Row"class="fa fa-minus-circle fa-lg removeRow"></i></div></span>'
			}else{
				cell.contentEditable = true;
			}
		});
		reconfigureRow();
	});
	
	$(document).on("click", ".removeCol" , function() {
		var index = $(this).parent().parent().parent()[0].cellIndex
		$("tr").each(function() {
			$(this).find("td:eq("+index+"),th:eq("+index+")").remove();
			//$(this).filter("td:eq("+index+")").remove(); with tr
		});
		reconfigureCol();
	});
 
	$(document).on("click", ".addCol" , function() {
		var cellIndex = $(this).parent().parent().parent()[0].cellIndex;
		$("#excel-table").find('tr').each(function(){
			   var trow = $(this);
			   
				 if(trow.index() === 0){
					header = numToChar(cellIndex+1);
					th = document.createElement('th');
					
					th.innerHTML = '<span><p>'+header+
									'</p><div class="dropdown">'+
										'<i title ="Add Row"class="fa fa-plus-circle fa-lg addCol"></i>'+       
										'<i title = "Remove Row"class="fa fa-minus-circle fa-lg removeCol"></i>'+
									'</div></span>';
					trow.children()[cellIndex].after(th);
					//trow[cellIndex].appendChild(th); 
				 }else{
					var cell = trow[0].insertCell(cellIndex+1);
					 cell.contentEditable = true;
				 }
				
			 });
			 reconfigureCol();
	});
 });
			
var copiedData = [];
function copyData(){
	$("#paste").prop('disabled', false);
	copiedData = [];
	$(".selected").each(function(index){
		data = {rowIndex : 0, cellIndex:0 , data:""};
		data.cellIndex = $(this)[0].cellIndex;
		data.rowIndex = $(this).parent()[0].rowIndex;
		data.data = $(this)[0].innerText;
		copiedData.push(data);
	});
}
function pasteData(){
	//console.log(copiedData)
	var currentCellIndex = $(".selected")[0].cellIndex;
	var currentRowIndex = $(".selected").parent()[0].rowIndex
	var cpRowFirstIndex = copiedData[0].rowIndex;
	var selectedCell = $(".selected");
	var currCell = selectedCell;
	for(var i = 0; i < copiedData.length ; i++){
		if(cpRowFirstIndex != copiedData[i].rowIndex){
			currCell = selectedCell.closest('tr').next().find('td:eq(' +currentCellIndex + ')');
			cpRowFirstIndex = copiedData[i].rowIndex;
			selectedCell = currCell;
		}
		if(currCell[0] != null){
			currCell[0].innerText = copiedData[i].data;
			currCell = currCell.next();
		}
			
	}
}
var reconfigureCol = function(){
	$("#excel-table").find('th:not(:first) span p').each(function(index){
			
			$(this)[0].innerText = numToChar(index+1);
			
	});
	$("#excel-table tr th:not(:first)").resizable({
		handles : 'e'
	});
}
var reconfigureRow = function(){
	$("#excel-table").find('tr:not(:first) td span p').each(function(index){
			$(this)[0].innerText = index+1;
			
	});
	$("td:first-child").resizable({
		minWidth : 50,
		handles : " s"
	})
}

charToNum = function(alpha) {
	var index = 0
	for(var i = 0, j = 1; i < j; i++, j++)  {
		if(alpha == numToChar(i))   {
			index = i;
			j = i;
		}
	}
	console.log(index);
}

numToChar = function(number)    {
	var numeric = (number - 1) % 26;
	var letter = chr(65 + numeric);
	var number2 = parseInt((number - 1) / 26);
	if (number2 > 0) {
		return numToChar(number2) + letter;
	} else {
		return letter;
	}
}
chr = function (codePt) {
	if (codePt > 0xFFFF) { 
		codePt -= 0x10000;
		return String.fromCharCode(0xD800 + (codePt >> 10), 0xDC00 + (codePt & 0x3FF));
	}
	return String.fromCharCode(codePt);
}

//selection of the cells on mouseover

var isMouseDown = false;
var startRowIndex = null;
var startCellIndex = null;
var currCell = $('td').first();
currCell = $(this); 
function selectTo(cell) {
    
    var row = cell.parent();    
    var cellIndex = cell.index();
    var rowIndex = row.index();
    
    var rowStart, rowEnd, cellStart, cellEnd;
    
    if (rowIndex < startRowIndex) {
        rowStart = rowIndex;
        rowEnd = startRowIndex;
    } else {
        rowStart = startRowIndex;
        rowEnd = rowIndex;
    }
    
    if (cellIndex < startCellIndex) {
        cellStart = cellIndex;
        cellEnd = startCellIndex;
    } else {
        cellStart = startCellIndex;
        cellEnd = cellIndex;
    }        
    
    for (var i = rowStart; i <= rowEnd; i++) {
        var rowCells = $("table").find("tr").eq(i).find("td");
        for (var j = cellStart; j <= cellEnd; j++) {
            rowCells.eq(j).addClass("selected");
        }        
    }
}
$(document).on("mousedown", "td", function (e) {
	
	 if (!$(this).hasClass('ui-resizable')) {
		$("#copy").prop('disabled', false);
		currCell = $(this);  
		
		isMouseDown = true;
		var cell = $(this);	
		$("table").find(".selected").removeClass("selected"); // deselect everything
		
		if (e.shiftKey) {
			selectTo(cell);                
		} else {
			cell.addClass("selected");
			startCellIndex = cell.index();
			startRowIndex = cell.parent().index();
		}
		currCell.focus()
	}
	
	return false; // prevent text selection
});
$(document).on("mouseover", "td", function () {
	if (!isMouseDown) return;
    $("table").find(".selected").removeClass("selected");
    selectTo($(this));
});
$(document).on("bind", "selectstart", function () {
	return false;
});
$(document).mouseup(function () {
    isMouseDown = false;
	
});
// key down captures
$(document).on("keydown", "table",function (e) {
	$("#copy").prop('disabled', false);
	var flag = true;
	var c = "";
	if (e.which == 39) {
		// Right Arrow
		c = currCell.next();
	} else if (e.which == 37) { 
		// Left Arrow
		c = currCell.prev();
	} else if (e.which == 38) { 
		// Up Arrow
		c = currCell.closest('tr').prev().find('td:eq(' + 
		  currCell.index() + ')');
	} else if (e.which == 40) { 
		// Down Arrow
		c = currCell.closest('tr').next().find('td:eq(' + 
		  currCell.index() + ')');
    } else if (e.which == 9 && !e.shiftKey) {
        // Tab
       // e.preventDefault();
        c = currCell.next();
    } else if (	e.which == 9 && e.shiftKey) {
        // Shift + Tab
	}else if (	e.which == 67 && e.ctrlKey) {
        // control + c
		e.preventDefault();
		console.log("copied")
		copyData();
		flag = false;
	}else if (	e.which == 86 && e.ctrlKey) {
        // Shift + Tab
		e.preventDefault();
		pasteData();
	}else if (e.ctrlKey) { 
		flag = false;
    }
	
	if(flag)
		$(".selected").removeClass("selected")
	// If we didn't hit a boundary, update the current cell
	if (c.length > 0) {
		currCell = c;
		currCell.focus();
	}
});

//convert the tabel data to csv format
tabelToCsv = function(){
	csv = []
	$("tr").each(function(index) {
		$cells = $(this).find("td:not(.ui-resizable)")
		csv_row = [];
		$cells.each(function(cellIndex) {
			txt = $(this)[0].innerText.trim();
			csv_row.push(txt.replace(",", "-"));
		});
		if(csv_row.length !=0)
			csv.push(csv_row.join(","));
	});
	
	output = csv.join("\n")
	//$("#csvData")[0].innerText = output;
	return output;
}

//convert the
tableToJson = function(){
	var myRows = [];
	var $headers = $("th");
	var $rows = $("tr:not(:first)").each(function(index) {
	  $cells = $(this).find("td");
	  myRows[index] = {};
	  $cells.each(function(cellIndex) {
		myRows[index][$($headers[cellIndex])[0].innerText.trim()] = $(this)[0].innerText.trim();
		//console.log($headers[cellIndex].innerText)
	  });    
	});

	// Let's put this in the object like you want and convert to JSON (Note: jQuery will also do this for you on the Ajax request)
	var myObj = {};
	myObj.myrows = myRows;
	//$("#jsonData")[0].innerHTML = JSON.stringify(myObj)
	return JSON.stringify(myObj);
}


// fetch data from local storage
function fetchData(){
	if (typeof(Storage) !== "undefined") {	
		// Retrieve
		var innerHTML = localStorage.getItem("dataTable")
		if(innerHTML != null){
			$("#dataTable")[0].innerHTML = localStorage.getItem("dataTable");
		  
		}
			console.log("fetching data");
		   //console.log(innerHTML);
	   
	} else {
		console.log("Sorry, your browser does not support Web Storage...");
	}
}
