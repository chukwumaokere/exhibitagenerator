// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
// https://ourcodeworld.com/articles/read/286/how-to-execute-a-python-script-and-retrieve-output-data-and-errors-in-node-js
//
var pss = require('python-shell');
var button = document.getElementById('go');

if (button){
	button.addEventListener('click', function(){
		var response = document.getElementById('menu_options').value;
		console.log(`response is ${response}`);
		window.location.replace(`./${response}.html`);
	});
}
var createdocbutton = document.getElementById('create_document');
var rangeinput = document.getElementById('range');
var allrows = document.getElementById('all-rows');

if(allrows){
	allrows.addEventListener('click', function(){
		if (allrows.checked == true){
			rangeinput.setAttribute('disabled', true);
		}else{
			rangeinput.removeAttribute('disabled');
		}
	});
}
if (createdocbutton){
	createdocbutton.addEventListener('click', function(){ 
		var spreadsheet_raw = document.getElementById('spreadsheet').value;
		var pieces = spreadsheet_raw.split('\\');
		var spreadsheet = pieces[pieces.length-1];

		var template_raw = document.getElementById('template').value;
		var pieces = template_raw.split('\\');
		var template = pieces[pieces.length-1];
	
		if (allrows.checked == true){
			var range = 'all';
		}else{
			var range = document.getElementById('range').value;
		}
		
		var sheetname = document.getElementById('sheetname').value;
		var path = document.location.pathname;
		var dir = path.substring(path.indexOf('/'), path.lastIndexOf('/'));
		
		options = {
			mode: 'text',
			args: [spreadsheet, template, sheetname, range]
		}
		pss.PythonShell.run('./execution.py', options, function(err, results){
			//if(err) throw err;
			console.log('results: ' + results);
			if (results[1] == 'Success'){
				alert("Documents Created!");
			}
			if (results[1] == undefined || !results){
				alert("An error occured: " + err);
			}
		});
		
	});
}