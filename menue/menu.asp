<%@ Language=VBScript %>
<%option explicit%>

<!-- created for XML-Creator V7.0 -->
<!-- (c) 2006 by P. Büschi for SSC-DESIGN-->
<!-- updated: 12.09.2008 - Link für Parentlinks! -->

<html>
<head>

<STYLE type="text/css">
<!--
BODY,HTML {
scrollbar-face-color:#F7FBFD;
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#FFFFFF;
scrollbar-darkshadow-color:#FFFFFF;
scrollbar-shadow-color:#000000;
scrollbar-arrow-color:#0A5989;
scrollbar-track-color:#FFFFFF;
}

	a { color:#000000; text-decoration:none; font-size: 10pt; font-family: Verdana, Arial, Helvetica, sans-serif }
b{width:152px;}
-->
</STYLE>

<meta http-equiv="Content-Language" content="de-ch">
<title>Dynamic XML-Menu</title>
<Meta http-equiv=content-type content=text/html;charset=iso-8859-1>
<link rel="stylesheet" type="text/css" href="style.css">
<link rel="stylesheet" type="text/css" href="stylePlaketten.css">

<style type="text/css">
<!-- Tabelleneintrag für zusätzlichen Abstand (NICHT verändern!!!) -->		
th { height:0px; }

<!-- Abstand des Textes vom linken Rand der Tabelle (NICHT verändern!!!) -->
.sub1 {	padding-left:5px; }
.sub2 {	padding-left:9px; }
.sub3 {	padding-left:12px; }
.line { height:1; }
</style>

<script src="scripts.js" type="text/javascript">
</script>

 
</head>

<%
' Die Schriftart bestimmen
Function setParams (sName, opt)
	If opt = "bold" Then 
		setParams = "<b>" & sName & "</b>"
	Elseif opt = "italic" Then 
		setParams = "<i>" & sName & "</i>"
	ElseIf opt = "underline" Then 
		setParams = "<u>" & sName & "</u>"
	Else 
		setParams = sName 	
	End If
End Function


' Die Links für die 1. Ebene werden aufbereiten!
Function writeLevel1 (sCnt, sName, sLink, vPic, sClicked, sBorder)

	' Anzeigename setzen
	sName = setParams (sName, strText)			
	
	' Wenn nötig, Linien zeichnen
	If strBor <> "um" Then
		writeLevel1 = writeLevel1 & "<tr><td class=""line""><img src=""strich.gif"" width=""100%"" height=""1""></td></tr>" & vbnewline
	End If
	
	' Wenn es sich um die Startseite handelt, im Parent Window öffnen
	If strMain = "1" Then
		writeLevel1 = writeLevel1 & "<tr><td onclick=""goHome(this,'" & sLink & "','')"" "  & vbnewline
	Else
		writeLevel1 = writeLevel1 & "<tr><td onclick=""menuclick(this,'menu.asp?id=" & sCnt & "&mod=cl','" & strBGClck & "',"
		writeLevel1 = writeLevel1 & "'" & strTXTClck & "','" & sLink & "','','" & strZwo & "')"" "  & vbnewline
	End If
	
	' Wenn keine eigene Farbe für das Item angegeben wurde, die MouseOut-Farbe verwenden
	If strBGL1 = "" Then strBGL1 = strBGOut End if
	If strTxL1 = "" Then strTxL1 = strTXTOut End if
	
	' Prüfen ob ein Border angezeigt werden muss
	dim sMyBorder
	If sBorder = "1" Then sMyBorder= "border:1px solid "& strBCol &";" End If
	
	' Eintrag aufbauen (MouseOver und MouseOut bei gelicktem Link weglassen!)
	If sClicked = "1" Then
		writeLevel1 = writeLevel1 & "bgcolor=""" & strBGClck & """ class=""sub1"" style=""" & sMyBorder & " color:" & strTXTClck & """><b>" & vbnewline
	Else
		writeLevel1 = writeLevel1 & "onmouseover=""menuchange(this,1,'" & strBGOver & "','" & strTXTOver & "','" & vPic & "')"" "
		writeLevel1 = writeLevel1 & "onmouseout=""menuchange(this,0,'" & strBGL1 & "','" & strTxL1 & "','" & vPic & "')"" "
		writeLevel1 = writeLevel1 & "bgcolor=""" & strBGL1 & """ class=""sub1"" style=""" & sMyBorder & " color:" & strTxL1 & """>" & vbnewline
	End If
	
	' Bei geklicktem Link (wenn nötig) das Bild anzeigen
	dim sMyPic
	If sClicked = "1" And FixPic = "1" Then 
		sMyPic = strPic 
	Else
		sMyPic = strPicOff	
	End If
	
	' Prüfen ob ein Bild angezeigt werden soll oder nicht
	If showPic = "1" Then		
		writeLevel1 = writeLevel1 & "<img name=""" & vPic & """ src=""" & sMyPic & """ class=""sub1"">&nbsp;"
		writeLevel1 = writeLevel1 & sName & "</b></td></tr>" & vbnewline
	Else
		writeLevel1 = writeLevel1 & sName & "</b></td></tr>" & vbnewline
	End If
	
	' Wenn nötig, einen zusätlichen Platzhalter einbinden
	If strHei <> "" Then writeLevel1 = writeLevel1 &  "<tr><th style=""height:" & strHei & "px;""></th></tr>" & vbnewline End If
	
End Function


' Die Links für die 2. Ebene werden aufbereiten!
Function writeLevel2 (sCnt, sID, sName, sLink, vPic, sClicked, sBorder)

	' Anzeigename setzen
	sName = setParams (sName, strText)
	
	' Wenn nötig, Linien zeichnen
	If strBor = "ll" Or strBor = "umll" Then
		writeLevel2 = writeLevel2 & "<tr><td class=""line"">"
		writeLevel2 = writeLevel2 & "<img border=""0"" src=""strich.gif"" width=""100%"" height=""1"">"
		writeLevel2 = writeLevel2 & "</td></tr>" & vbnewline
	End If
	
	' Wenn keine eigene Farbe für das Item angegeben wurde, die MouseOut-Farbe verwenden
	If strBGL2 = "" Then strBGL2 = strBGOut End if
	If strTxL2 = "" Then strTxL2 = strTXTOut End if
	
	' Prüfen ob ein Border angezeigt werden muss
	dim sMyBorder
	If sBorder = "1" Then sMyBorder= "border:1px solid "& strBCol &";" End If
	
	' Eintrag aufbauen (MouseOver und MouseOut bei gelicktem Link weglassen!)
	If sClicked = "1" Then
		writeLevel2 = writeLevel2 & "<tr><td bgcolor=""" & strBGClck & """ class=""sub2"" style=""" & sMyBorder & " color:" & strTXTClck & """><b>" & vbnewline
	Else
		writeLevel2 = writeLevel2 & "<tr><td onclick=""menuclick(this,'menu.asp?id=" & sCnt & "&sid=" & sID & "&mod=cl','" & strBGClck & "',"
		writeLevel2 = writeLevel2 & "'" & strTXTClck & "','" & sLink & "','" & vPic & "','" & strZwo & "')"" "
		writeLevel2 = writeLevel2 & "onmouseover=""menuchange(this,1,'" & strBGOver & "','" & strTXTOver & "','" & vPic & "')"" "
		writeLevel2 = writeLevel2 & "onmouseout=""menuchange(this,0,'" & strBGL2 & "','" & strTxL2 & "','" & vPic & "')"" "
		writeLevel2 = writeLevel2 & "bgcolor=""" & strBGL2 & """ class=""sub2"" style=""" & sMyBorder & " color:" & strTxL2 & """>" & vbnewline
	End If
	
	' Bei geklicktem Link (wenn nötig) das Bild anzeigen
	dim sMyPic
	If sClicked = "1" And FixPic = "1" Then 
		sMyPic = strPic 
	Else
		sMyPic = strPicOff	
	End If
			  
	' Prüfen ob ein Bild angezeigt werden soll oder nicht
	If showPic = "1" Then
		writeLevel2 = writeLevel2 & "<img name=""" & vPic & """ src=""" & sMyPic & """ class=""sub2"">&nbsp;"
		writeLevel2 = writeLevel2 & sName & "</b></td></tr>" & vbnewline
	Else
		writeLevel2 = writeLevel2 & sName & "</b></td></tr>" & vbnewline
	End If		
	
	' Wenn nötig, einen zusätlichen Platzhalter einbinden
	If strHei <> "" Then writeLevel2 = writeLevel2 &  "<tr><th style=""height:" & strHei & "px;""></th></tr>" & vbnewline End If
	  				  
End Function


' Die Links für die 3. Ebene werden aufbereiten!
Function writeLevel3 (sCnt, sID, sSID, sName, sLink, vPic, sClicked, sBorder)
	
	' Anzeigename setzen
	sName = setParams (sName, strText)			
	
	' Wenn nötig, Linien zeichnen
	If strBor = "ll" Or strBor = "umll" Then
		writeLevel3 = writeLevel3 & "<tr><td class=""line"">"
		writeLevel3 = writeLevel3 & "<img border=""0"" src=""strich.gif"" width=""100%"" height=""1""></td></tr>" & vbnewline
	End If	
	
	' Wenn keine eigene Farbe für das Item angegeben wurde, die MouseOut-Farbe verwenden
	If trim(strBGL3) = "" Then strBGL3 = strBGOut End if
	If trim(strTxL3) = "" Then strTxL3 = strTXTOut End if		
	
	' Prüfen ob ein Border angezeigt werden muss
	dim sMyBorder
	If sBorder = "1" Then sMyBorder= "border:1px solid "& strBCol &";" End If
	
	' Eintrag aufbauen (MouseOver und MouseOut bei gelicktem Link weglassen!)
	If sClicked = "1" Then
		writeLevel3 = writeLevel3 & "<tr><td bgcolor=""" & strBGClck & """ class=""sub3"" style=""" & sMyBorder & " color:" & strTXTClck & """><b>" & vbnewline
	Else
		writeLevel3 = writeLevel3 & "<tr><td onclick=""menuclick(this,'menu.asp?id=" & sCnt & "&sid=" & sID & "&ssid=" & sSID & "&mod=cl',"
		writeLevel3 = writeLevel3 & "'" & strBGClck & "','" & strTXTClck & "','" & sLink & "','" & vPic & "','" & strZwo & "')"" "
		writeLevel3 = writeLevel3 & "onmouseover=""menuchange(this,1,'" & strBGOver & "','" & strTXTOver & "','" & vPic & "')"" "
		writeLevel3 = writeLevel3 & "onmouseout=""menuchange(this,0,'" & strBGL3 & "','" & strTxL3 & "','" & vPic & "')"" "
		writeLevel3 = writeLevel3 & "bgcolor=""" & strBGL3 & """ class=""sub3"" style=""" & sMyBorder & " color:" & strTxL3 & """>" & vbnewline
	End If
		
	' Bei geklicktem Link (wenn nötig) das Bild anzeigen
	dim sMyPic
	If sClicked = "1" And FixPic = "1" Then 
		sMyPic = strPic 
	Else
		sMyPic = strPicOff	
	End If
		  
	' Prüfen ob ein Bild angezeigt werden soll oder nicht
	If showPic = "1" Then
		writeLevel3 = writeLevel3 & "<img name=""" & vPic & """ src=""" & sMyPic & """ class=""sub3"">&nbsp;"
		writeLevel3 = writeLevel3 & sName & "</td></tr>" & vbnewline
	Else
		writeLevel3 = writeLevel3 & sName & "</td></tr>" & vbnewline
	End If	
	
	' Wenn nötig, einen zusätlichen Platzhalter einbinden
	If strHei <> "" Then writeLevel3 = writeLevel3 &  "<tr><th style=""height:" & strHei & "px;""></th></tr>" & vbnewline End If			 
	
End Function
%>

<%
' Variablen setzten
dim strFilename, XMLDoc, strXMLLocation, bSuccess, hasClicked
strFilename = "data.xml"

' Das XML-File laden
Set XMLDoc = Server.CreateObject("Microsoft.XMLDOM")
XMLDoc.async = False
strXMLLocation = Server.MapPath(strFilename)
bSuccess = XmlDoc.load(strXMLLocation)
If Not bSuccess Then
	Response.Write "Loading the XML file <b>" & _
		strXMLLocation & "</b> failed!"
	Response.End
End If
 
' Prüfen ob ein Link angeklickt wurde
If Request.QueryString("mod") = "cl" Then 
	hasClicked = 1 
Else 
	hasClicked = 0 
End If
  
' Attirbute aus dem XML-File holen
Dim rootNode, xmlVar, strBGCol, strMTop, strMLeft, strMRight, strBGOver, strBGOut, strBGClck
Dim strTXTOver, strTXTOut, strTXTClck, strBor, strBCol, strPic, strPicOff, ShowPic, FixPic
Dim strTopImg, strTopSpc, strBotImg, strBotSpc, strGAlign

Set rootNode  = XMLDoc.documentElement
Set xmlVar 	  = rootNode.getElementsByTagName("body")
	strBGCol  = xmlVar.item(0).getAttribute("color")
	strMTop   = xmlVar.item(0).getAttribute("mtop")
	strMLeft  = xmlVar.item(0).getAttribute("mleft")
	strMRight = xmlVar.item(0).getAttribute("mright")
Set xmlVar    = rootNode.getElementsByTagName("bgColor")
	strBGOver = xmlVar.item(0).getAttribute("bgOver")
	strBGOut  = xmlVar.item(0).getAttribute("bgOut")
	strBGClck = xmlVar.item(0).getAttribute("bgClicked")
Set xmlVar    = rootNode.getElementsByTagName("txtColor")
	strTXTOver = xmlVar.item(0).getAttribute("txtOver")
	strTXTOut = xmlVar.item(0).getAttribute("txtOut")
	strTXTClck = xmlVar.item(0).getAttribute("txtClicked")
Set xmlVar	  = rootNode.getElementsByTagName("border")
	strBor	  = xmlVar.item(0).getAttribute("art")
	strBCol	  = xmlVar.item(0).getAttribute("color")
Set xmlVar    = rootNode.getElementsByTagName("picture")
	strPic    = xmlVar.item(0).getAttribute("on")
	strPicOff = xmlVar.item(0).getAttribute("off")
	ShowPic   = xmlVar.item(0).getAttribute("showIt")
	FixPic    = xmlVar.item(0).getAttribute("fix")
Set xmlVar	  = rootNode.getElementsByTagName("graphics")
	strTopImg = xmlVar.item(0).getAttribute("topGraphic")
	strTopSpc = xmlVar.item(0).getAttribute("topGSpace")
	strBotImg = xmlVar.item(0).getAttribute("bottomGraphic")	
	strBotSpc = xmlVar.item(0).getAttribute("bottomGSpace")
	strGAlign = xmlVar.item(0).getAttribute("graphicAlign")
	

' Ab hier beginnt der Aufbau des Menus
If rootNode.hasChildNodes() Then
	
	' Bodytag aufbauen
	Response.Write "<body onLoad=""loadXML('" & strFilename & "')"" bgcolor=""" & strBGCol & """ style=""margin-top:"&strMTop&"px;" & _
				   "margin-left:"&strMLeft&"px; margin-right:"&strMRight&"px;"">" & vbnewline
	
	' Graphic oberhalb von Menue anzeigen
	If strTopImg <> "" Then
		If strGAlign <> "" Then
			Response.Write "<" & strGAlign & "><img src="""& strTopImg & """ border=""0"" style=""margin-bottom:" & strTopSpc & "px;""></" & strGAlign & ">"
		Else
			Response.Write "<img src="""& strTopImg & """ border=""0"" style=""margin-bottom:" & strTopSpc & "px;"">"
		End If
	End If
	
	' Tabelle aufbauen
	Response.Write "<table id=""menu"" style=""border-width:0px; width:156px; border-collapse:collapse"" border=""0"" cellspacing=""0"">" & vbnewline

	' Erste Stufe befüllen
	Dim myItem, i, Cnt, VarPic
	Set myItem = rootNode.getElementsByTagName("item")
	
	For i = 1 To myItem.length
		
		'Prüfen ob ein Image als Zeiger aufgebaut werden muss
		If ShowPic = "1" Then
			Cnt = Cnt + 1
			VarPic = "Pic" & Cnt
		End If		

		' Die einzelnen Attribute aus dem XML-File auslesen
		Dim strCnt, strBGL1, strTXL1, strText, strMain, strHei, strZwo, strName, strLink
		strCnt   = myItem.item(i-1).getAttribute("id")
		strBGL1  = myItem.item(i-1).getAttribute("bgcol")
		strTXL1  = myItem.item(i-1).getAttribute("txtcol")
		strText  = myItem.item(i-1).getAttribute("txt")
		strMain  = myItem.item(i-1).getAttribute("isMain")		
		strHei   = myItem.item(i-1).getAttribute("space")		
		strZwo   = myItem.item(i-1).getAttribute("Frame2")
		strName  = myItem.item(i-1).childNodes(0).text
		strLink  = myItem.item(i-1).childNodes(1).text
		
		' Um zu prüfen ob der Pfeil stehen bleiben muss, bereits hier setzen
		Dim subNode
		Set subNode = myItem.item(i-1).getElementsByTagName("subitem")		
				
		' Texteigenschaft übernehmen
		strName = setParams (strName, strText)	

		' Wenn auf einen Eintrag geklickt wurde, folgendes setzen						
		If hasClicked = 1 Then
		
			If strCnt = Request.QueryString("id") AND strLink <> "#" Then	
				
				' Mit Picture bzw. ohne Picture
				If FixPic = "1" and ShowPic = "1" Then					
					' Wenn nötig, Rahmen zeichnen
					If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then						
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "1")
					Else
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "0")
					End If
					
				ElseIf FixPic = "0" and ShowPic = "1" Then					
					' Wenn nötig, Rahmen zeichnen
					If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "1")
					Else
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "0")
					End If
					
				Else					
					' Wenn nötig, Rahmen zeichnen
					If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "1")
					Else
						Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "1", "0")
					End If
				End If
			
			Else
				Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "0", "0")				
			End If
		
		Else
			Response.Write writeLevel1 (strCnt, strName, strLink, VarPic, "0", "0")			
		End If				


		' Zweite Stufe befüllen		
		Dim ii
		If subNode.length > 0 Then
			For ii = 1 To subNode.length
				
				'Prüfen ob ein Image als Zeiger aufgebaut werden muss
				If ShowPic = "1" Then
					Cnt = Cnt + 1
					VarPic = "Pic" & Cnt
				End If		

				' Die einzelnen Attribute aus dem XML-File auslesen
				Dim strID, strUpID, strBGL2, strTXL2
				strID    = subNode.item(ii-1).getAttribute("id")
				strUpID  = subNode.item(ii-1).getAttribute("upid")
				strBGL2  = subNode.item(ii-1).getAttribute("bgcol")
				strTXL2  = subNode.item(ii-1).getAttribute("txtcol")
				strText  = subNode.item(ii-1).getAttribute("txt")
				strHei   = subNode.item(ii-1).getAttribute("space")
				strZwo   = subNode.item(ii-1).getAttribute("Frame2")
				
				' Um zu prüfen ob der Pfeil stehen bleiben muss, bereits hier setzen
				Dim subSubNode
				Set subSubNode = subNode.item(ii-1).getElementsByTagName("subitems")		
				
				' Texteigenschaft übernehmen
				strName = setParams (strName, strText)	

				If strUpID = Request.QueryString("id") Then					
					' Name und Link holen				
					strName = subNode.item(ii-1).childNodes(0).text			
					strLink = subNode.item(ii-1).childNodes(1).text				
				
					If hasClicked = 1 Then
					
						If strID = Request.QueryString("sid") AND strLink <> "#" Then
																												
							' Mit Picture bzw. ohne Picture
							If FixPic = "1" and ShowPic = "1" Then
								' Wenn nötig, Rahmen zeichnen
								If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "1")
								Else
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "0")							
								End If	
																																	
							ElseIf FixPic = "0" and ShowPic = "1" Then
								' Wenn nötig, Rahmen zeichnen
								If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then								
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "1")
								Else
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "0")								
								End If
							Else								
								' Wenn nötig, Rahmen zeichnen
								If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then									
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "1")
								Else
									Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "1", "0")					
								End If
							End If
							
						Else
							Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "0", "0")							
						End If
						
					Else
						Response.Write writeLevel2 (strCnt, strID, strName, strLink, VarPic, "0", "0")						
					End If
					
				End If	
				
				
				' Dritte Stufe befüllen									
				Dim iii
				If subSubNode.length > 0 Then					
					For iii = 1 To subSubNode.length
						
						'Prüfen ob ein Image als Zeiger aufgebaut werden muss
						If ShowPic = "1" Then
							Cnt = Cnt + 1
							VarPic = "Pic" & Cnt
						End If		

						' Die einzelnen Attribute aus dem XML-File auslesen
						Dim strIDL3, strBGL3, strTXL3
						strIDL3  = subSubNode.item(iii-1).getAttribute("id")
						strUpID  = subSubNode.item(iii-1).getAttribute("upid")
						strBGL3  = subSubNode.item(iii-1).getAttribute("bgcol")
						strTXL3  = subSubNode.item(iii-1).getAttribute("txtcol")
						strText  = subSubNode.item(iii-1).getAttribute("txt")
						strHei   = subSubNode.item(iii-1).getAttribute("space")
						strZwo   = subSubNode.item(iii-1).getAttribute("Frame2")
						
						' Texteigenschaft übernehmen
						strName = setParams (strName, strText)	
						
						If strUpID = Request.QueryString("sid") Then						
							' Name und Link holen
							strName = subSubNode.item(iii-1).childNodes(0).text
							strLink = subSubNode.item(iii-1).childNodes(1).text
							
							If hasClicked = 1 Then
															
								If strIDL3 = Request.QueryString("ssid") Then	
																
									' Mit Picture bzw. ohne Picture
									If FixPic = "1" and ShowPic = "1" Then
										' Wenn nötig, Rahmen zeichnen
										If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "1")
										Else
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "0")
										End If				
																
									ElseIf FixPic = "0" and ShowPic = "1" Then
										' Wenn nötig, Rahmen zeichnen
										If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "1")										
										Else
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "0")																			
										End If
										
									Else										
										' Wenn nötig, Rahmen zeichnen
										If strBor = "um" Or strBor = "umhl" Or strBor = "umll" Then
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "1")								
										Else
											Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "1", "0")							
										End If
									End If
									
								Else																
									Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "0", "0")
								End If
								
							Else
								Response.Write writeLevel3 (strCnt, strID, strIDL3, strName, strLink, VarPic, "0", "0")
							End If
							
						End If
						
					Next
					
				End If					
				
			Next
			
		End If
		
	Next
	
	' Letze Linie zeichnen
	If strBor <> "um" Then
		Response.Write "<tr><td class=""line""><img border=""0"" src=""strich.gif"" width=""100%"" height=""1""></td></tr>"
	End If	
	Response.Write "</table>"
	
	' Graphic unterhalb von Menue anzeigen
	If strBotImg <> "" Then
		If strGAlign <> "" Then
			Response.Write "<" & strGAlign & "><img src="""& strBotImg & """ border=""0"" style=""margin-top:" & strBotSpc & "px;""></" & strGAlign & ">"
		Else
			Response.Write "<img src="""& strBotImg & """ border=""0"" style=""margin-top:" & strBotSpc & "px;"">"
		End If
	End If
	
Else
  	Response.Write "Keine Items gefunden!"
End If
%>

</body>
<base target="Main">

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" bgcolor="#FFFFFF" background="../images/ZE_02.jpg">
























<div align="center">
























<table border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" width="117">
	<tr>
		<td valign="top">
		<p style="text-align: center">
		&nbsp;</td>
	</tr>
		
	<tr style="padding-top:30px;">
	
		<td>
		<p style="word-spacing: 1px; margin-bottom: 0; text-align:right">
		<span style="letter-spacing: 1px">
		<font face="Tahoma" color="#333333" size="1"><span style="font-size: 8pt">
		<img src="http://counter.genotec.ch/?33485_zunft" alt="Counter Genotec" border="0" align="absbottom" hspace="8"></span></font><font face="Tahoma"><span style="font-size: 8pt">B<font style="font-size: 8pt; letter-spacing:1px">esucher</font></span></font><font style="font-size: 8pt" face="Tahoma"><br></font>
		<font style="font-size: 8pt; letter-spacing:1px" face="Tahoma">unserer Website</font><span style="font-size: 7pt"><br></span>
		<font face="Tahoma"><span style="font-size: 8pt">&nbsp;<font style="font-size: 8pt; letter-spacing:1px">E2010 l A2014</font></span></font></td>
	</tr>
</table>




















































</div>


























</html>