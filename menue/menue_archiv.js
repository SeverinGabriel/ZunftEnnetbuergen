
<!--
/******************************************************************/
/**** ACHTUNG: Nur die Werte zwischen den Sternen verändern!!! ****/
/******************************************************************/
/****************************************** created by P. Büschi **/
/******************************************************************/

var kzif = 0;
var cnt = 0;
var tdef, stop;
var col1, col2;
var txt1, txt2;
var on, off;

/***************************************************************************************/
/****             Hier können die Farben und Pfeile bestimmt werden                 ****/
/***************************************************************************************/

col1 = '#08429C';		// Die Farbe für den Effekt
col2 = '#ECF5FB';		// Hintegrund Farbe des Frames

txt1 = '#FFFFFF';		// Die Textfarbet für den Effekt
txt2 = '#000000';		// Die Standardfarbe des Textes


on   = 'pfeil.gif';		// Filename des Pfeiles
off  = 'blank.gif';		// Filename des Blanks (muss gleich gross sein wie das Bild!!!)

/**************************************************************************************/


/* Hier wird auf die Startseite verlinkt */
function goHome(td,url,name) {
	top.location.href = url;
}


/* Hier wird die Hintergrundfarbe im Menü verändert */
function menuchange(td,MO,name){			
	if (tdef!=td) {
		if (kzif==0) {
			if (MO==1){td.style.background=col1; td.style.color=txt1;} 
			if (MO==0){td.style.background=col2; td.style.color=txt2;}
		}
	}
	else {
		kzif = 0;		
	}
	
	if (name!="") {
		if (name!=stop){
			if (MO==1) {this.document[name].src=on;}
			if (MO==0) {this.document[name].src=off;}
		}
	}
}


/* Hier wird auf die gewählte Seite, im richtigen Frame, gelinkt */
function menuclick(td,url,name) {
	kzif = 1
	td.style.background=col1;	
	if (cnt>0) {
		if (name!="") {
			this.document[name].src=on;
			this.document[stop].src=off;
		}
		tdef.style.background=col2;
		tdef.style.color=txt2;
	}
	cnt = cnt + 1;
	stop=name;
	tdef = td;

	/*****************************************************/
	/** Hier kann die Zielframenummer angepasst werden **/
	/*****************************************************/
	
		parent.frames[1].location = url;
	
	/****************************************************/	
}
//-->
