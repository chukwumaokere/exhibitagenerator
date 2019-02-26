// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
// https://ourcodeworld.com/articles/read/286/how-to-execute-a-python-script-and-retrieve-output-data-and-errors-in-node-js
//
/* Checking Application Version, then Python version */
const {shell} = require('electron');
var pjson = require('./package.json');
var version = pjson.version;

/* Deprecated Update Method. Using Squirrel to handle updates now.
var xhttp = new XMLHttpRequest();
xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
       // Action to be performed when the document is read;
	  var currentVersion = xhttp.responseText;
	    if (version != currentVersion){
		console.log(`Version mismatch. Installed Version: ${version} vs Current Version ${currentVersion}`);
		var update = confirm(`There is a newer version available! Would you like to download and update now?`);
		    if (update){
			window.location.replace("https://chukwumaokere.com/exhibitagenerator/exhibitagenerator.zip");
		    }
	    }else{
		console.log(`Installed Version: ${version} vs Current Version ${currentVersion}`);
	    }
    }
};
xhttp.open("GET", "https://chukwumaokere.com/exhibitagenerator/version.php", true);
xhttp.send();
*/

//Begin Python checking
var pss = require('python-shell');
var OS = navigator.platform;
var checked = false;
var URL = [];
        URL['Win32'] = 'https://www.python.org/ftp/python/3.7.1/python-3.7.1-amd64.exe';
        URL['MacIntel'] = "https://www.python.org/ftp/python/3.7.1/python-3.7.1-macosx10.9.pkg";

var python;
var pythonVersion = pss.PythonShell.getVersion().then(resolve => {
        var pythonVer = resolve.stdout; console.log(resolve.stdout);

        if (pythonVersion && typeof(pythonVer) === 'string'){
                var python = true;
                console.log('Python installed');
        }else{
                var python = false;
                console.log('Python not installed');
                console.log(typeof(pythonVer));
                console.log(pythonVersion);
        }
		checkVersion(python, checked);
}).catch(function(err){
        if (err.message.includes("Command failed")){
                var python = false;
                console.log('Python not installed');
        }
        console.log(((err)));
		checkVersion(python, checked);
});
/* End Version Checking */

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
		var fullSSPath = document.getElementById('spreadsheet').files[0].path;
		var fullDOCXPath = document.getElementById('template').files[0].path;
		var outputPath = fullDOCXPath.substring(0, fullDOCXPath.lastIndexOf('\\')) + "\\" + "output\\";

		console.log('outputPath: ' + outputPath );
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
			args: [fullSSPath, fullDOCXPath, sheetname, range, outputPath]
		}

		var href = window.location.href;
		var dir = href.substring(8, href.lastIndexOf('/')) + "/";
		pss.PythonShell.run(dir+'./execution.py', options, function(err, results){
			console.log('results: ' + results);
			if(results){
				if (results[1] == 'Success'){
					var resp = confirm("Documents Created! Would you like to view the documents now?");
					if(resp == true){
						shell.openItem(outputPath);
					}
				}
				if (results[1] == undefined || !results){
					alert("An error occured: " + err);
				}
			}
			else{
				if(err){
					alert(err);
					console.log(err);
				}else{
					alert("An unknown error has occured");
				}
			}
		});
		
	});
}

function checkVersion(ver, checked){
        if (ver !== true && checked == false){
        var resp = confirm("Python 3.7.1 is needed for this app to run properly. Download now?");
                if (resp == true){
                        window.location.replace(URL[OS]);
                        checked = true;
                }else{
                        checked = true;
                }
        }
}


