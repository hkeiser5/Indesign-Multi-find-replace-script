/* JavaScript Find & Replace with Grep Indd script for catalog product copy
written by Heather C Keiser, copyright 2019.
------------------------------------------------------------------

Instructions on how to use
Create a csv file from an excell sheet. The excell sheet must have headers.
One column must have the header "find" and another column must have
the header "replace". Open the csv file when prompted by the script,
then every value in the find coolumn is found and replaced by the 
corosponding value in the replace column.

----------------------------------------------------------------*/

//work with all open docs get linked csv file called "propogate.csv"
var myCurrentDoc = app.activeDocument; //selects current active document

 

var myPropogate = File.openDialog('Select a CSV File','comma-separated-values(*.csv):*.csv;');

//create 2 arrays based off two columns in the propogate csv called "find" and "replace"
var myFind = [];
var myReplace = [];

if (myPropogate != null)
{
      
    myPropogate.open('r',undefined, undefined);//read the file
    var myContents =  myPropogate.read(); //get info in file
    myPropogate.close();//close the file
    myPropogate = myContents;//reassign mycontents to mypropogate
    //we can split this into lines, and those lines can be turned into arrays. The first line has the headers, so from there we can get the index for the appropriate headers. Then for each line, the arrayed index gets pushed into the appropriate array.
	
	//create an array of every row in the csv
	var rows = myPropogate.split('\n');

	//now turn every element in the rows into an array (break it by columns)
	for (var i = 0; i< rows.length; i++){
		rows[i] = rows[i].split(',');
	}
	
	//find the index of find & replace
	var findIndex;
	var replaceIndex;
	for (var i=0; i< rows[0].length; i++){
		if (rows[0][i] == 'find'){
			findIndex = i;
		} else if (rows[0][i] == 'replace'){
			replaceIndex = i;
		}
	}
	
	//push values from every row by those indexes into appropriate arrays skip index0, as that is just headers
	for (var i = 1; i< rows.length; i++){
        if (rows[i][findIndex] != null){   
            //do nothing
        }else{
		myFind.push(rows[i][findIndex]);
		myReplace.push(rows[i][replaceIndex]);
        }
	}
}

//need to escape characters in the find data so that grep search finds correctly
var slash = String('\\');


//for loop i=0, i<= array length, i++, find match from find, and replace with replace data utilizing grep replace.
for (var i=0; i< myFind.length; i++){
	
	//escape special characters in find to prepare for grep   
    myFind[i] = myFind[i].replace(/\\/g,slash + '\\');
    myFind[i] = myFind[i].replace(/\*/g,slash + '*');
    myFind[i] = myFind[i].replace(/\./g,slash + '.');
    myFind[i] = myFind[i].replace(/\+/g,slash + '+');
    myFind[i] = myFind[i].replace(/\?/g,slash + '?');
    myFind[i] = myFind[i].replace(/\^/g,slash + '^');
    myFind[i] = myFind[i].replace(/\$/g,slash + '$');
    myFind[i] = myFind[i].replace(/\(/g,slash + '(');
    myFind[i] = myFind[i].replace(/\)/g,slash + ')');
    myFind[i] = myFind[i].replace(/\</g,slash + '<');
    myFind[i] = myFind[i].replace(/\>/g,slash + '>');
    myFind[i] = myFind[i].replace(/\{/g,slash + '{');
    myFind[i] = myFind[i].replace(/\[/g,slash + '[');
    myFind[i] = myFind[i].replace(/\|/g,slash + '|');
     
    
	//Clear the find/change text preferences.
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
    
    

	//Set the GREP find options 
	app.findChangeGrepOptions.includeFootnotes = false;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = true;

	//Look for the pattern and change to
	app.findGrepPreferences.findWhat = myFind[i];
	app.changeGrepPreferences.changeTo = myReplace[i];
	myCurrentDoc.changeGrep();

	//Clear the find/change text preferences.
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
}// JavaScript Document