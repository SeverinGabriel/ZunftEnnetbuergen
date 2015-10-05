<!--
/***************************************************************************************/
/**** 		   JavaScript zur Steuerung von Menu's				    ****/
/***************************************************************************************/
/***************************************************************    XML-Creator V7.0  **/
/***************************************************************************************/
/*************************************************************** created by P. B�schi **/
/***************************************************************************************/
var kzif = 0;
var cnt = 0;
var tdef, stop;
var col1, col2;
var txt1, txt2;
var on, off;
var frame, frame2, link2
var xmlDoc, ObjXML, xmlVar;
var bgnam = "bgimg";
/***************************************************************************************/
/**** 				ACHTUNG: Im Code nichts ändern!!!		    ****/
/****				Damit die Funktionalität gewährleistet bleibt.	    ****/
/***************************************************************************************/

/* Den XML Parser laden */
function loadXML(xmlFile)
{	
	// code for Internet Explorer9
	if (window.ActiveXObject) {
		xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		xmlDoc.async = false;
		xmlDoc.onreadystatechange = verify;
		xmlDoc.load(xmlFile);
		readXML()
	}
	// code for Mozilla, etc.
	else if (document.implementation && document.implementation.createDocument) {	
		var xmlhttp = new window.XMLHttpRequest();
		xmlhttp.open("GET",xmlFile,false);
		xmlhttp.send(null);
		xmlDoc = xmlhttp.responseXML.documentElement;
	/*
	else if (document.implementation && document.implementation.createDocument) {
		xmlDoc = document.implementation.createDocument("","",null);
		xmlDoc.load(xmlFile);
		xmlDoc.onload = readXML
	*/
	}
	// 
	else
	{
		alert('Your browser cannot handle this treemenu!');
	}
}

/* Die Daten aus dem XML-File auslesen */
function readXML()
{
	xmlVar = xmlDoc.documentElement.getElementsByTagName("bgColor");
	col1   = xmlVar.item(0).getAttribute("bgOver");
	col2   = xmlVar.item(0).getAttribute("bgOut");
		
	xmlVar = xmlDoc.documentElement.getElementsByTagName("txtColor");
	txt1   = xmlVar.item(0).getAttribute("txtOver");
	txt2   = xmlVar.item(0).getAttribute("txtOut");
	
	xmlVar = xmlDoc.documentElement.getElementsByTagName("frame");
	frame  = xmlVar.item(0).getAttribute("num");
	
	xmlVar = xmlDoc.documentElement.getElementsByTagName("picture");
	on = new Image();
	on.src = xmlVar.item(0).getAttribute("on");
	off = new Image();
	off.src = xmlVar.item(0).getAttribute("off");
}

/* Überprüfen, dass das File richtig geladen wurde */
function verify()
{
  // 0 Object is not initialized
  // 1 Loading object is loading data
  // 2 Loaded object has loaded data
  // 3 Data from object can be worked with
  // 4 Object completely initialized
  if (xmlDoc.readyState != 4)
  {
      return false;
  }
}

/* Hier wird auf die Startseite verlinkt */
function goHome(td,url,name) {
	top.location.href = url;
}

/* Hier wird die Hintergrundfarbe im Menü verändert */
function menuchange(td,mod,ccol,tcol,name){			
	if (tcol == '') { tcol = txt2; }
	if (tdef!=td) {
		if (kzif == 0) {		
			if (col1 != '') {
				if (mod==1){td.style.background=col1; td.style.color=txt1;} 
			}
			if (mod==1){td.style.color=txt1;} 
			if (mod==0){td.style.background=ccol; td.style.color=tcol;}
		}
	}
	else {
		kzif = 0;		
	}
	if (name!="") {
		if (name!=stop){
			if (mod==1) {this.document[name].on;}
			if (mod==0) {this.document[name].off;}
		}
	}
}

/* Hier wird auf die gewählte Seite, im richtigen Frame, gelinkt */
function menuclick(td,url,ccol,tcol,alink,name,frameZwo) {

	kzif = 1;
	if (ccol != '') { td.style.background=ccol;	}
	if (tcol == '') { tcol = txt2; }
	else { td.style.color=tcol; }
	
	if (cnt>0) {
		if (name!="") {
			this.document[name].on;
			this.document[stop].off;
		}
		tdef.style.background=ccol;
		tdef.style.color=tcol;
	}
	cnt = cnt + 1;
	stop=name;
	tdef = td;
	
	if (alink == "#") {
		location.href = url;
	}
	else {
		if (alink != "") {			
			if (url != "") {				
				location.href = url;				
			}
			//parent.frames[frame].location = alink;
			parent.Main.location.href = alink;
			//parent[frame].location.href = alink;
		}	
	}	
	
	// Zweites Frame wechseln
	if (frameZwo != "") {
		frame2 = frameZwo.substring(0,1);
		link2  = frameZwo.substr(1);
		parent.frames[frame2].location = link2;
	}
}

//-->
