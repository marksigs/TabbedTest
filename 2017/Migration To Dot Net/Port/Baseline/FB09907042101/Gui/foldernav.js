/*
 ______________________________________________________
/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯\
|           Folder Navigation 2.0 by EAE               |
|                                                      |
|   Based on a cross browser outline sample found      |
|   at 'www.webreference.com/dhtml'                    |
|                                                      |
|   Feel free to copy, use and change this script as   |
|   long as this part remains unchanged.               |
|                                                      |
|   If you have any questions and or comments please   |
|   E-mail me 'eae@eae.net'. If you're looking for     |
|   more JavaScripts etc, please check out my webpage  |
|                 'www.eae.net/weebfx'                 |
|                                                      |
|              Last Updated: 17 July 1998              |
\______________________________________________________/
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/

document.onmouseover = mOver ;
document.onmouseout = mOut ;

function mOver() {
	var eSrc = window.event.srcElement ;
	if (eSrc.className == "item") {
		window.event.srcElement.className = "highlight";
	}
}

function mOut() {
	var eSrc = window.event.srcElement ;
	if (eSrc.className == "highlight") {
		window.event.srcElement.className = "item";
	}
}


var bV=parseInt(navigator.appVersion);
NS4=(document.layers) ? true : false;
IE4=((document.all)&&(bV>=4))?true:false;
ver4 = (NS4 || IE4) ? true : false;

isExpanded = false;

function getIndex($1) {
	ind = null;
	for (i=0; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		if (whichEl.id == $1) {
			ind = i;
			break;
		}
	}
	return ind;
}

function arrange() {
	nextY = document.layers[firstInd].pageY + document.layers[firstInd].document.height;
	for (i=firstInd+1; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		if (whichEl.visibility != "hide") {
			whichEl.pageY = nextY;
			nextY += whichEl.document.height;
		}
	}
}

function FolderInit(){
	if (NS4) {
	firstEl = "mParent";
	firstInd = getIndex(firstEl);
	showAll();
		for (i=0; i<document.layers.length; i++) {
			whichEl = document.layers[i];
			if (whichEl.id.indexOf("Child") != -1) whichEl.visibility = "hide";
		}
		arrange();
	}
	else {
		tempColl = document.all.tags("DIV");
		for (i=0; i<tempColl.length; i++) {
			if (tempColl(i).className == "child") tempColl(i).style.display = "none";
		}
	}
}

function FolderExpand($1,$2) {
	if (!ver4) return;
	if (IE4) { ExpandIE($1,$2) } 
	else { ExpandNS($1,$2) }
}

function ExpandIE($1,$2) {
	Expanda = eval($1 + "a");
	Expanda.blur()
	ExpandChild = eval($1 + "Child");
        if ($2 != "top") { 
		ExpandTree = eval($1 + "Tree");
		ExpandFolder = eval($1 + "Folder");
	}
	if (ExpandChild.style.display == "none") {
		ExpandChild.style.display = "block";
                if ($2 != "top") { 
                	if ($2 == "last") { ExpandTree.src = "images/Lminus.gif"; }
			else { ExpandTree.src = "images/Tminus.gif"; }
			ExpandFolder.src = "images/openfoldericon.gif";	
		}
		else { mTree.src = "images/topopen.gif"; }
	}
	else {
		ExpandChild.style.display = "none";
                if ($2 != "top") { 
	                if ($2 == "last") { ExpandTree.src = "images/Lplus.gif"; }
			else { ExpandTree.src = "images/Tplus.gif"; }
			ExpandFolder.src = "images/foldericon.gif";
		}
		else { mTree.src = "images/topopen.gif"; }
	}
}
function ExpandNS($1,$2) {
	ExpandChild = eval("document." + $1 + "Child")
        if ($2 != "top") { 
		ExpandTree = eval("document." + $1 + "Parent.document." + $1 + "Tree")
		ExpandFolder = eval("document." + $1 + "Parent.document." + $1 + "Folder")
	}	
	if (ExpandChild.visibility == "hide") {
		ExpandChild.visibility = "show";
                if ($2 != "top") { 
               		if ($2 == "last") { ExpandTree.src = "images/Lminus.gif"; }
			else { ExpandTree.src = "images/Tminus.gif"; }
			ExpandFolder.src = "images/openfoldericon.gif";	
		}
		else { mTree.src = "images/topopen.gif"; }
	}
	else {
		ExpandChild.visibility = "hide";
                if ($2 != "top") { 
               		if ($2 == "last") { ExpandTree.src = "images/Lplus.gif"; }
			else { ExpandTree.src = "images/Tplus.gif"; }
			ExpandFolder.src = "images/foldericon.gif";	
		}
		else { mTree.src = "images/topopen.gif"; }
	}
	arrange();
}

function showAll() {
	for (i=firstInd; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		whichEl.visibility = "show";
	}
}


with (document) {
	write("<STYLE TYPE='text/css'>");
	if (NS4) {
		write(".parent { color: white; font-size:8pt; font-family:sans-serif; line-height:0pt; color:white; text-decoration:none; margin-top: 0px; margin-bottom: 0px; position:absolute; visibility:hidden }");
		write(".child { text-decoration:none; font-size:8pt; font-family:sans-serif; line-height:15pt; position:absolute }");
	        write(".item { color: white; text-decoration:none; font-family:sans-serif}");
	        write(".highlight { color: red; text-decoration:none; font-family:sans-serif }");
	}
	else {
		write(".parent { font-size: 8pt; font-family:sans-serif; text-decoration: none; color: white }");
		write(".child { font-size: 8pt; font-family:sans-serif; display:none }");
	        write(".item { color: white; text-decoration:none; cursor: hand }");
	        write(".highlight { color: red; text-decoration:none }");
	        write(".icon { margin-right: 5 }")
	}
	write("</STYLE>");
}

onload = FolderInit;