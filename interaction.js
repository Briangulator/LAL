/*
	Library Activity Logger(LAL), Version 1.1
	© 2017, Brian Jameson, Blacksburg High School
	Licensed under GPL Version 3 licenses.
	Requires the ExcelPlus JavaScript Library
	http://aymkdn.github.io/ExcelPlus/
*/

/*
	Global variables
	Each 2D Array stores data to be exported in a different sheet within the same Excel file.  The data for the first row is initialized here.
	The ExcelPlus object allows us to read and write Excel files.
	The bell schedule for the school is defaulted to "normal".
	Three switches are declared which change whether or not logged out students are hidden.
*/
var dataStorage = [[["Name", "Time Out", "Time In", "Destination"]],
				[["Name", "Time In", "Time Out", "Reason", "Teacher"]],
				[["Name", "Time In", "Time Out", "Reason", "Student Count"]],
				[["Name", "Time", "Tardy?"]],
				[["Name", "Time", "Tardy?"]],
				[["Name", "Time", "Tardy?"]],
				[["Name", "Time", "Tardy?"]]],
	sheetNames = ["Hall Pass", "Visitors", "Class Visitors", "Class 1", "Class 2", "Class 3", "Class 4"],
	schedule = "normal",
	epx = new ExcelPlus(),
	lIOV = false,
	lIOC = false,
	lIOH = false;

/**
	@name addToXLSX()
	
	@description Reads an uploaded file, appends today's data to the end, and then saves as "data.xlsx".
	
	@params None
*/
function addToXLSX(){
	//Checks correct sheet amount.  Could check for actual sheet names, but this is a small enough project that a file with exactly 7 sheets probably won't ever be opened.
	if (epx.getSheetNames().length != 7){
		alert("Invalid file uploaded!");
		return;
	}
	
	//Signs out all students so that there isn't HTML code in the saved file
	endAll();
	
	//Create in-function data storage so that .shift() doesn't modify the actual data
	for (var i = 0, sheet = dataStorage[i].slice(0); i < dataStorage.length; i++){
		//Gets rid of the first row header, which an existing file should have already.
		sheet.shift();
		//Add the date to the end of each row, as data only needed an hour/minute timestamp up to this point.
		sheet = addDate(sheets);
		//Select each sheet in the loaded file and write data to it
		epx.selectSheet(sheetNames[i]);
		writeData(epx, sheet);
	}
	
	//epx.saveAs(monthWord()+".xlsx"); //Use this if Ms. Christle wants monthly reports.  Better IMO than the current implementation.
	epx.saveAs("data.xlsx");//One big file for the whole year may become too big for Excel to open.
}

/**
	@name writeData()
	
	@description Adds inputted data to the end of an Excel spreadsheet.
	
	@param [Object] ep The ExcelPlus object to add data to.
	
	@param [2D Array] data The data to write to Excel.
*/
function writeData(ep, data){
	for (i = 0; i < data.length; i++)
		ep.writeNextRow(data[i])
}

/**
	@name monthWord()
	
	@description Returns the current month as a string.
	
	@params None
*/
function monthWord(){
	//This code is a slight modification of http://www.w3schools.com/jsref/jsref_getmonth.asp
	//It is functionally the same.
	var d = new Date(),
		month = ["January",
			"February",
			"March",
			"April",
			"May",
			"June",
			"July",
			"August",
			"September",
			"October",
			"November",
			"December"];
	return month[d.getMonth()];
}

/**
	@name loggedInOnly()
	
	@description Updates the page to either display only students that are still logged in or return to normal viewing.
	
	@param [String] caller A single character identifying whether to update the "Visitors", "Class Visit", or "Hallpass" tab
*/
function loggedInOnly(caller){
	//Switch based on what called this function
	switch (caller){
		//Visitors tab
		case "v":
			//Update display
			constructDisplay(dataStorage[1], "vdisplay");
			document.getElementById("liov").innerHTML = lIOV ? "View Logged In Only" : "View All";
			
			//Flip switch
			lIOV = !lIOV;
			break;
		
		//Hallpass tab
		case "h":
			//Update display
			constructDisplay(dataStorage[0], "hdisplay");
			document.getElementById("lioh").innerHTML = lIOH ? "View Logged In Only" : "View All";
			
			//Flip switch
			lIOH = !lIOH;
			break;
		
		//Class visit tab
		case "c":
			//Update display
			constructDisplay(dataStorage[2], "cdisplay");
			document.getElementById("lioc").innerHTML = lIOC ? "View Logged In Only" : "View All";
			
			//Flip switch
			lIOC = !lIOC;
			break;
		
		//Error message if called incorrectly
		default: alert("Error in loggedInOnly()!");
	}
}

/**
	@name lIODisplay()
	
	@description Removes all signed out students from the input data set.
	
	@param [2D Array] data The data to modify
*/
function lIODisplay(data){
	for (var i = 0; i < data.length; i++){
		
		//The greatest length a time can have is 5, and the sign out button is far longer.
		if (data[i][2].length <= 5){
			data.splice(i,1);
			
			//If we spliced, we have to cancel out the loop increment
			i--;
		}
	}
	return data;
}

/**
	@name addDate()
	
	@description Adds the date onto the end of each row.
	
	@param [2D Array] data The data to modify
*/
function addDate(data){
	for (var i = 0, d = new Date(); i < data.length; i++)
		data[i].push(d.toLocaleDateString());
	return data;
}

/**
	@name ExportXLSX()
	
	@description Downloads the data collected in an Excel file
	
	@param [Boolean] end Whether or not to log out all students
*/
function exportXLSX(end){
	var ep = new ExcelPlus();
	
	//Checks if we want to sign out everyone
	if (end)
		endAll();
	
	//Creates the file
	ep.createFile(sheetNames);
	
	//Writes data to each sheet
	for(var i = 0; i < dataStorage.length; i++)
		ep.write({"sheet":sheetNames[i], "content":dataStorage[i]});
	
	//Exports using a timestamp as the filename, as this function will be used up to 1080 times a year.
	var d = new Date();
	ep.saveAs(d.toLocaleString().replace("/", "-")+".xlsx");
}

/**
	@name tabSelect()
	
	@description Facilitates change between inpage tabs
	
	@param [event] evt Event that triggered this method.  Lets us know which button called this function.
*/
function tabSelect(evt,tab){
	//This code is basically copypasta'd from http://www.w3schools.com/howto/howto_js_tabs.asp
	var tabcontent = document.getElementsByClassName("tabcontent"), i,
		tablinks = document.getElementsByClassName("tablinks");
	
	//Don't display anything
	for (i = 0; i < tabcontent.length; i++)
		tabcontent[i].style.display = "none";
	
	//Set all tab buttons to inactive
	for (i = 0; i < tablinks.length; i++)
		tablinks[i].className = tablinks[i].className.replace(" active","");
	
	//Display the correct tab
	document.getElementById(tab).style.display = "block";
	
	//Set the correct button as active so the user knows which tab they're in
	evt.currentTarget.className += " active";
}

/**
	@name endEntry()
	
	@description Logs target student out
	
	@param [Integer] index Where the student is in the log
	
	@param [2D Array] dispArray The array to sign the student out from
	
	@param [String] Identifies the display to update after the change is made
*/
function endEntry(index, dispArray, dispTarget){
	//Replace the button(always index 2) with the current 12-hour time.  24-hour is better, but I have my orders.
	dispArray[index][2] = getTime("string12");
	constructDisplay(dispArray, dispTarget);
}

/**
	@name removeEntry()
	
	@description Lets the librarian delete bad entries, such as profane names.
	
	@param [Integer] index Where the student is in the log
	
	@param [String] tgt A string noting the array to sign the student out from
*/
function removeEntry(index, tgt){
	switch (tgt){
		//Erroneous visitor
		case "vdisplay":
			dataStorage[1].splice(index, 1);
			constructDisplay(dataStorage[1], tgt);
			break;
		
		//Erroneous class
		case "cdisplay":
			dataStorage[2].splice(index, 1);
			constructDisplay(dataStorage[2], tgt);
			break;
		
		//Erroneous hallpass user
		case "hdisplay":
			dataStorage[0].splice(index, 1);
			constructDisplay(dataStorage[0], tgt);
			break;
	}
}

/**
	@name endAll()
	
	@description Signs out EVERYBODY
	
	@params None
*/
function endAll(){
	//Simply refer to other functions
	endVisitors();
	endCVisits();
	endHallPass();
}

/**
	@name endVisitors()
	
	@description Signs out all of the library's visitors
	
	@params None
*/
function endVisitors(){
	for (var i = 0; i < dataStorage[1].length; i++)
		//The sign out button has a length far greater than 10, and the timestamp there after endEntry() is never more than 5 long.
		if (dataStorage[1][i][2].length > 10)
			endEntry(i, dataStorage[1], "vdisplay");
}

/**
	@name endCVisits()
	
	@description Signs out all of the library's class visitors
	
	@params None
*/
function endCVisits(){
	for (var i = 0; i < dataStorage[2].length; i++)
		//The sign out button has a length far greater than 10, and the timestamp there after endEntry() is never more than 5 long.
		if (dataStorage[2][i][2].length > 10)
			endEntry(i, dataStorage[2], "cdisplay");
}

/**
	@name endHallPass()
	
	@description Signs in all students that are in the halls
	
	@params None
*/
function endHallPass(){
	for (var i = 0; i < dataStorage[0].length; i++)
		//The sign in button has a length far greater than 10, and the timestamp there after endEntry() is never more than 5 long.
		if (dataStorage[0][i][2].length > 10)
			endEntry(i, dataStorage[0], "hdisplay");
}

/**
	@name constructDisplay()
	
	@description Updates the display of currently signed in/out students.
	
	@param [2D Array] dS The data to display
	
	@param [String] caller A string identifying which display to update
*/
function constructDisplay(dS, caller){
	var dataSource = dS.slice(0), //Make sure not to affect actual data with functions like .reverse()
		dispTable = document.getElementById(caller), //Get the table to display data in
		displayString = ""; //Display starts empty and then fills
	
	//If we only want to display the logged in students, call lIODisplay().  Only do so if the target display is set to logged-in only.
	switch (caller){
		case "vdisplay":
			if (lIOV)
				dataSource = lIODisplay(dataSource);
			break;
		case "cdisplay":
			if (lIOC)
				dataSource = lIODisplay(dataSource);
			break;
		case "hdisplay":
			if (lIOH)
				dataSource = lIODisplay(dataSource);
			break;
		default: alert("Error in constructDisplay!");
	}
	
	//The students should be displayed from newest to oldest.  The first row is a header, so it still goes first.
	dataSource.reverse();
	dataSource.unshift(dataSource.pop());
	
	//For each row, add a <tr> tag
	for(var i = 0, j; i < dataSource.length; i++){
		displayString += "<tr>";
		//For each item in each row, add a <td> tag or a <th> if it's the first row header.
		for(j = 0; j < dataSource[0].length; j++){
			displayString += i == 0 ? "<th>" : "<td>";
			displayString += dataSource[i][j];
			displayString += i == 0 ? "</th>" : "</td>";
			j++;
		}
		
		//If the admin menu is open, add an X button to get rid of wrong or profane entries
		if (document.getElementById("sselector") != null && i > 0)
			displayString+="<td><button type=\"button\" onclick=\"removeEntry("+(dataSource.length-i)+", \'"+caller+"\')\">X</button></td>";
		
		//End table row
		displayString+="</tr>";
	}
	
	//Update the table with new contents
	dispTable.innerHTML = displayString;
}

/**
	@name submitData()
	
	@description Adds a new entry to the data sheets
	
	@param [String] caller Identifies why the student is signing in, and therefore where their data should go
*/
function submitData(caller){
	var tStamp = getTime("string12"), name;
	
	//Data validation
	if (document.getElementById(caller + "fname").value.length == 0) {
		alert("You need to input a first name!");
		return;
	}
	if (document.getElementById(caller + "lname").value.length == 0) {
		alert("You need to input a last name!");
		return;
	}
	//Change depending on where the input is coming from
	switch (caller){
		//Attendance tab
		case "a":
			//Get the name and then reset the name fields to null
			name = document.getElementById("afname").value+" "+document.getElementById("alname").value;
			document.getElementById("afname").value = "";
			document.getElementById("alname").value = "";
			
			//Changes where the data goes and whether the student is tardy depending on the time
			switch (whichClass()){
				case 0:
					dataStorage[3][dataStorage[3].length] = new Array(name, tStamp, false);
					break;
				case 1:
					dataStorage[3][dataStorage[3].length] = new Array(name, tStamp, true);
					break;
				case 2:
					dataStorage[4][dataStorage[4].length] = new Array(name, tStamp, false);
					break;
				case 3:
					dataStorage[4][dataStorage[4].length] = new Array(name, tStamp, true);
					break;
				case 4:
					dataStorage[5][dataStorage[5].length] = new Array(name, tStamp, false);
					break;
				case 5:
					dataStorage[5][dataStorage[5].length] = new Array(name, tStamp, true);
					break;
				case 6:
					dataStorage[6][dataStorage[6].length] = new Array(name, tStamp, false);
					break;
				case 7:
					dataStorage[6][dataStorage[6].length] = new Array(name, tStamp, true);
					break;
				//If the school day is over, there isn't any point to this program, so tell the student to go home
				case 8:
					alert("The school day has ended!");
					return;
				default:
					alert("Error in submitData()!");
					return;
			}
			alert(name + " is now signed in!");
			break;
		
		//Hallpass tab
		case "h":
			var destBox = document.getElementById("hdestination"),
				dest = destBox.value;
			
			if (document.getElementById("hteacher") != null) {
				var teach = document.getElementById("hteacher").value;
				if (teach.length == 0) {
					alert("You need to input the teacher you're going to!");
					return;
				}
				dest += ": "+teach;
			}
			if (document.getElementById("hcounselor") != null) {
				var couns = document.getElementById("hcounselor").value;
				if (couns.length == 0) {
					alert("You need to input the counselor you're going to!");
					return;
				}
				dest += ": " + couns;
			}
			
			//Get the name and then reset the name fields to null
			name = document.getElementById("hfname").value + " " + document.getElementById("hlname").value;
			document.getElementById("hfname").value = "";
			document.getElementById("hlname").value = "";
			
			//Add the entry.
			dataStorage[0][dataStorage[0].length] = new Array(name,
				tStamp,
				"<button type=\"button\" onclick=\"endEntry(" + dataStorage[0].length + ", dataStorage[0], \'hdisplay\')\">In</button>",
				dest);
			
			//Reset the destination boxes and update the display
			destBox.value = "Bathroom";
			hallPassTeachers();
			constructDisplay(dataStorage[0], "hdisplay");
			break;
		
		//Visitors tab
		case "v":
			//Get the name and then reset the name fields to null
			name = document.getElementById("vfname").value + " " + document.getElementById("vlname").value;
			document.getElementById("vfname").value = "";
			document.getElementById("vlname").value = "";
			
			//Add the entry.
			dataStorage[1][dataStorage[1].length] = new Array(name,
				tStamp,
				"<button type=\"button\" onclick=\"endEntry(" + dataStorage[1].length + ", dataStorage[1], \'vdisplay\')\">Out</button>",
				document.getElementById("vreason").value,
				document.getElementById("vteacher").value);
			
			//Reset the reason box and update the display
			document.getElementById("vreason").value = "Work independently on a class assignment";
			document.getElementById("vteacher").value = "C. Baker";
			constructDisplay(dataStorage[1], "vdisplay");
			break;
		
		//Class visits tab
		case "c":
			//Data validation
			if(document.getElementById("creason").value.length == 0){
				alert("You need to input a reason!");
				return;
			}
			if(document.getElementById("classsize").value.length == 0){
				alert("You need to input the number of students!");
				return;
			}
			
			//Get the name and then reset the name fields to null
			name = document.getElementById("cfname").value + " " + document.getElementById("clname").value;
			document.getElementById("cfname").value = "";
			document.getElementById("clname").value = "";
			
			//Add the entry.
			dataStorage[2][dataStorage[2].length] = new Array(name,
				tStamp,
				"<button type=\"button\" onclick=\"endEntry(" + dataStorage[2].length + ", dataStorage[2], \'cdisplay\')\">Out</button>",
				document.getElementById("creason").value,
				document.getElementById("classsize").value);
			
			//Reset the reason and class size boxes and update the display
			constructDisplay(dataStorage[2], "cdisplay");
			document.getElementById("creason").value = "";
			document.getElementById("classsize").value = "";
			break;
		default:
			//Error message in case of problems
			alert("Error in submitData()!");
	}
}
/**
	@name getTime()
	
	@description Gives the time in the formats needed for this program
	
	@param [String] format Specifies which data format is needed.
*/
function getTime(format){
	var d = new Date(),
		h = d.getHours(),
		m = d.getMinutes();
	switch (format){
		
		//Returns an array with hours and minutes
		case "array":
			return [h, m];
		
		//Returns a string with the time in 12-hour format.
		case "string12":
			if(h > 12)
				h -= 12;
			return h + (m < 10 ? ":0" : ":") + m;
		
		//Returns a string with the time in a 24-hour format.  Superior to 12-hour format, but is unused.
		case "string24":
			return h + (m < 10 ? ":0" : ":") + m;
		
		//Returns the number of minutes passed today
		case "minct":
			return 60 * h + m;
		
		//Error just in case
		default:
			alert("Invalid getTime() argument!");
	}
}
/**
	@name passWord()
	
	@description Employs simple password protection for the admin menu
	
	@param [String] pW The password given by the user
*/
function passWord(pW){
	if(passWordCheck(document.getElementById(pW).value)){
		//If the password is correct, display what is REALLY in the Librarian Only tab.
		document.getElementById("l").innerHTML="<p>\
				<button type=\"button\" onclick=\"exportXLSX(true)\">Export Data</button>\
			</p>\
			<p>\
				<object id=\"file-object\"></object>\
				<button type=\"button\" onclick=\"addToXLSX()\">Add Data to Selected File</button>\
				<a href=\"https://chrome.google.com/webstore/detail/downloads-overwrite-exist/fkomnceojfhfkgjgcijfahmgeljomcfk\">This Chrome extension is required to update an existing spreadsheet.</a>\
			</p>\
			<p>\
				Schedule for Day:\
				<select id=\"sselector\" onchange=\"schedule = document.getElementById(\"sselector\").value;\">\
					<option value=\"normal\">Standard day</option>\
					<option value=\"earlyReleaseOrClubDay\">Early Release/Club Day</option>\
					<option value=\"oneHourDelay\">1 Hour Delay</option>\
					<option value=\"twoHourDelay\">2 Hour Delay</option>\
				</select>\
			</p>\
			<button type=\"button\" onclick=\"endVisitors(); endCVisits()\">Sign Out All Visitors</button>\
			<button type=\"button\" onclick=\"closeMenu()\">Lock Librarian Menu</button>";
		
		//Set up the schedule selector and file selector
		document.getElementById("sselector").value = schedule;
		epx.openLocal({"labelButton":"Select Excel File"}, function(){});
		
		//Refresh the displays
		constructDisplay(dataStorage[0], "hdisplay");
		constructDisplay(dataStorage[1], "vdisplay");
		constructDisplay(dataStorage[2], "cdisplay");
	}
	else //Self-explanatory
		alert("Incorrect password!");
}

/**
	@name closeMenu()
	
	@description Closes the Librarian Only menu
	
	@params None
*/
function closeMenu(){
	//Simple text replacement
	document.getElementById("l").innerHTML = "<span><input type=\"password\" id=\"pwbox\" placeholder=\"Password\"></span>\
		<span><button onclick=\"passWord('pwbox')\">Submit</button></span>";
	
	//Refresh the displays
	constructDisplay(dataStorage[0], "hdisplay");
	constructDisplay(dataStorage[1], "vdisplay");
	constructDisplay(dataStorage[2], "cdisplay");
}

/**
	@name hallPassTeachers()
	
	@description Adds an additional input field to the Hallpass menu if needed
	
	@params None
*/
function hallPassTeachers(){
	//Get destination data and a target to aim the possible field at
	var hp = document.getElementById("hdestination").value,
		tgt = document.getElementById("moreinputs");
	
	//Switch based on the selection
	switch(hp){
		case "Teacher":
			tgt.innerHTML = "<input type=\"text\" id=\"hteacher\" placeholder=\"Teacher\">";
			break;
		case "Guidance":
			tgt.innerHTML = "<input type=\"text\" id=\"hcounselor\" placeholder=\"Guidance Counselor\">";
			break;
		default:
			tgt.innerHTML = "";
	}
}

/**
	@name passWordCheck()
	
	@description Checks the inputted password against the correct password.  This is not remotely secure, but it's enough to keep an average student out.
	
	@param [String] pW The inputted password
*/
function passWordCheck(pW){
	//ASCII codes for each character in the password
	var codes = [66, 72, 83, 76, 105, 98, 114, 97, 114, 121];
	
	//Make sure each character is the same as in the password
	if (codes.length != pW.length)
		return false;
	for (var i = 0; i < codes.length; i++){
		if ((codes[i]) != pW.charCodeAt(i))
			return false;
	}
	return true;
}

/**
	@name whichClass()
	
	@description Returns the # of the current class, starting at 0.  Odd #'s are returned when the student is late.
	
	@params None
*/
function whichClass(){
	var times = beginTimes(), i, time = getTime("minct");
	for (i = 0; i < times.length && time >= times[i]; i++){}
	return i;
}

/**
	@name beginTimes()
	
	@description Returns an array of class start times depending on the bell schedule.
	
	@params None
*/
function beginTimes(){
	//Choose based on an admin menu selector
	switch (schedule){
		case "earlyReleaseOrClubDay":
			return [486, 540, 546, 595, 601, 685, 691, 740];
		case "oneHourDelay":
			return [546, 620, 626, 700, 706, 810, 816, 890];
		case "twoHourDelay":
			return [606, 665, 671, 730, 736, 825, 831, 890];
		default:
			return [486, 575, 581, 670, 676, 795, 801, 890];
	}
}

//Repeat the exportXLSX function every hour for data backup
setInterval(function(){
	exportXLSX(false);
}, 3600000);
