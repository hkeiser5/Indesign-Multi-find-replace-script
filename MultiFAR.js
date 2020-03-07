/* JavaScript Find & Replace with Grep Indd script for catalog product copy
written by Heather C Keiser, copyright 2019.
------------------------------------------------------------------

Instructions on how to use
Create a csv UTF-8 file from an excell sheet. The excell sheet must have headers.
One column must have the header "find" and another column must have
the header "replace". Open the csv file when prompted by the script,
then every value in the find coolumn is found and replaced by the 
corosponding value in the replace column.

Note this will not work if any of your fields have a quote followed by comma
within the fields.
example: "hello", she said
example2: 5",
both of the above will fail and cause the wrong values to be found/replaced.

----------------------------------------------------------------*/

//work with all open docs get linked csv file called "propogate.csv"
var myCurrentDoc = app.activeDocument; //selects current active document

var myPropogate = File.openDialog('Select a CSV File','comma-separated-values(*.csv):*.csv;');

//create 2 arrays based off two columns in the propogate csv called "find" and "replace"
var myFind = [];
var myReplace = [];

if (myPropogate != null){
      
    myPropogate.open('r',undefined, undefined);//read the file
    var myContents =  myPropogate.read(); //get info in file
    myPropogate.close();//close the file
    myPropogate = myContents;//reassign mycontents to mypropogate
    //we can split this into lines, and those lines can be turned into arrays. The first line has the headers, so from there we can get the index for the appropriate headers. Then for each line, the arrayed index gets pushed into the appropriate array.
	
	//create an array of every row in the csv
	var rows = myPropogate.split('\n');
    //im getting an empty line at the end of this from a PC.
    for (var empty = 0; empty<rows.length; empty++){
            if (rows[empty] === ''){
                rows.splice(empty, 1);
                }
        }
	//now turn every element in the rows into an array (break it by columns)
	for (var l = 0; l< rows.length; l++){
		//add a comma to front and end of each line
		var comma = ',';
		rows[l] = comma.concat(rows[l], comma);
		var comIndices = [];
		for (var n = 0; n < rows[l].length; n++) {
			if (rows[l][n] === ',') {
				comIndices.push(n);
			}
		}
		var skipIndex = -1;//starts neg so that first , always is a starter
		var newrow = [];
		//create array based off comma indicies
		for (var r =0; r < comIndices.length-1; r++){
			var curIndex = comIndices[r];//comma we are checking
			var checknext = 1;
			var nextIndex = comIndices[r+checknext];
			var lastIndex = comIndices[comIndices.length-1];
			//need to reset skip index if on the last 2nd to last index in string & to set rows[i]
			if (r === comIndices.length-2){//this is last iteration and has not been represented in the while loop below
				if (skipIndex === lastIndex){
					//do nothing already pushed
				}else{
                       newrow.push(rows[l].slice(curIndex + 1,nextIndex));
				}
				
				skipIndex = -1;//reset skip for next loop set
				rows[l] = newrow;
				newrow = [];//reset newrow for next set
			} else if (rows[l][curIndex+1] === '"'){//this section is a quote group
				//while loop to find the end of the quote group
				var quoteend = false;
				while (quoteend === false ){//need a check if nextIndex is last index
					if (nextIndex === lastIndex){//this is last group
							
						newrow.push(rows[l].slice(curIndex + 1,nextIndex));
						skipIndex = lastIndex;
						quoteend = true;
					} else if (rows[l][nextIndex-1] === '"'){//this is end of quotes
							
						newrow.push(rows[l].slice(curIndex + 1,nextIndex));
						skipIndex = nextIndex-1;
						quoteend = true;	
					} else {
						checknext += 1;
						nextIndex = comIndices[r+checknext];
					}
					tester += 1;
				}//end while loop
                   
			} else if (curIndex > skipIndex){//this section has no comma, and is not part of another comma push to next indici
                newrow.push(rows[l].slice(curIndex + 1,nextIndex));
			} //else do nothing and go to next loop
		}//end making the array based off commas

        //now clean up all the extra quotes that the csv data created
		for (var j=0; j<rows[l].length; j++){
           
			//delete single quotes at front and end, as those were added in creation of csv
			rows[l][j] = rows[l][j].replace(/^"(?!")/,'');
			//neg look behind doesn't work-need to check another way
			if (rows[l][j][rows[l][j].length-2] !== '"'){
				rows[l][j] = rows[l][j].replace(/"($|\n)/,'');
			}
			//change triple and double quotes to single quotes
			rows[l][j] = rows[l][j].replace(/"""/g,'"');
			rows[l][j] = rows[l][j].replace(/""/g,'"');
		}//end clean up quotes
	
		//find the index of find & replace
		var findIndex;
		var replaceIndex;
		for (var i=0; i< rows[0].length; i++){
			if (rows[0][i] == 'find'){
				findIndex = i;
			} else if (rows[0][i] == 'replace'){
				replaceIndex = i;
			}
		}//end find indexes for find/replace

    
		//push values from every row by those indexes into appropriate arrays skip index0, as that is just headers
        if (l > 0){
            myFind.push(rows[l][findIndex]);
            myReplace.push(rows[l][replaceIndex]);
        }///end push to find/replace
	}//end breakdown csv into arrays

	//need to escape characters in the find data so that grep search finds correctly
	var slash = String('\\');
	//for loop i=0, i<= array length, i++, find match from find, and replace with replace data utilizing grep replace.
	for (var f=0; f< myFind.length; f++){
	
		//escape special characters in find to prepare for grep   
		myFind[f] = myFind[f].replace(/\\/g,slash + '\\');
		myFind[f] = myFind[f].replace(/\*/g,slash + '*');
		myFind[f] = myFind[f].replace(/\./g,slash + '.');
		myFind[f] = myFind[f].replace(/\+/g,slash + '+');
		myFind[f] = myFind[f].replace(/\?/g,slash + '?');
		myFind[f] = myFind[f].replace(/\^/g,slash + '^');
		myFind[f] = myFind[f].replace(/\$/g,slash + '$');
		myFind[f] = myFind[f].replace(/\(/g,slash + '(');
		myFind[f] = myFind[f].replace(/\)/g,slash + ')');
		myFind[f] = myFind[f].replace(/\</g,slash + '<');
		myFind[f] = myFind[f].replace(/\>/g,slash + '>');
		myFind[f] = myFind[f].replace(/\{/g,slash + '{');
		myFind[f] = myFind[f].replace(/\[/g,slash + '[');
		myFind[f] = myFind[f].replace(/\|/g,slash + '|');
     
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
		app.findGrepPreferences.findWhat = myFind[f];
		app.changeGrepPreferences.changeTo = myReplace[f];
		myCurrentDoc.changeGrep();

		//Clear the find/change text preferences.
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
	}
	
}// JavaScript Document