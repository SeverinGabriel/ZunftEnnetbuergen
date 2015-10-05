<%
on error resume next

' ==============================================================================================================================
'                           Überprüfe obligatorische Felder                                                            
' ==============================================================================================================================

referrer = split(Request.ServerVariables("HTTP_REFERER"),"/") 

if request("_Mandatory") <> "" then
  Mandatory = split(request("_Mandatory"),";")
    for i = lbound(Mandatory) to Ubound (Mandatory)
		if request(Mandatory(i)) = "" then message = message & "<br>Das Feld : " & Mandatory(i) & " darf nicht leer sein!<br>"
	next
End if
if request("_Referer") = ""           then message = message & "Es wurde keine Variable _Referer angegeben !<br>"
if request("_Host") = ""              then message = message & "Es wurde keine Variable _Host angegeben !<br>"
if request("_EMail") = ""             then message = message & "Es wurde keine Variable _EMail angegeben !<br>"
if request("_DankePage") = ""         then message = message & "Es wurde keine Variable _DankePage angegeben !<br>"
if request("_FehlerPage") = ""        then message = message & "Es wurde keine Variable _FehlerPage angegeben !<br>"

arrReferers = split(request("_Referer"), ";")
iRef = False
for x = 0 to Ubound(arrReferers)
  if arrReferers(x) = referrer(2) then
    iRef = True
  end if
next
if iRef = False then
  message = message & "<br>Der Referrer (<b>"& referrer(2) &"</b>) stimmt nicht dem von Ihnen definierten Referer (<b>"& request("_Referer") &"</b>) überein !<br>"
end if

if message <> "" then 
   response.write "<h2>Formmail - Fehler !</h2>"
   response.write "<font color=red>" & message & "</font>"
   response.write "<hr>"
   response.write "Weitere Informationen und Anleitungen unter : <a href=http://support.genotec.ch>http://support.genotec.ch</a>"
   response.end
end if

' ==============================================================================================================================
'                            Speichere Die Daten in einen Array                                                                 
' ==============================================================================================================================

If request.querystring <> "" then
For i=1 To request.querystring.count
	if not left(request.querystring.key(i),1) = "_" then 
	   c = c + 1 
	   redim preserve Formular(2,c)
	   Formular(0,c) = request.querystring.key(i)
	   Formular(1,c) = request.querystring.item(i)
	 end if
Next
End if
If request.form <> "" then
For i=1 To request.form.count
	if not left(request.form.key(i),1) = "_" then 
	   c = c + 1 
	   redim preserve Formular(2,c)
	   Formular(0,c) = request.form.key(i)
	   Formular(1,c) = request.form.item(i)
	 end if
Next
End if

' ==============================================================================================================================
'                            Sortiere Die Daten Gemäss Wunsch                                                                   
' ==============================================================================================================================

select case lcase(request("_Sort"))
case "desc"
        min = lbound(Formular,2)
        max = ubound(Formular,2)
	  	For i = min To max - 1
    	 smallest_value = Formular(0,i)
    	 smallest_j = i
    	 For j = i + 1 To max
		 if Formular(0,j) < smallest_value then
    		smallest_value = Formular(0,j)
    		smallest_j = j
    	 End if
    	 Next
    	 if smallest_j <> i Then
    		For intA = 0 To ubound(Formular,2)
    			temp = Formular(intA,smallest_j)
    			Formular(intA,smallest_j) = Formular(intA,i)
    			Formular(intA,i) = temp
    		Next
    	  End if
    	Next
		for fw = ubound(Formular,2) to 1 step -1
		    counter = counter + 1
			redim preserve tmpSearch(2,counter)
		    tmpSearch(0,counter) = Formular(0,fw)
			tmpSearch(1,counter) = Formular(1,fw)
		 next
		 sortFormular = tmpSearch
case "asc"
        min = lbound(Formular,2)
        max = ubound(Formular,2)
    	For i = min To max - 1
    	smallest_value = Formular(0,i)
    	smallest_j = i
    	For j = i + 1 To max
    		if strComp(Formular(0,j),smallest_value,vbTextCompare) = -1 Then
    			smallest_value = Formular(0,j)
    			smallest_j = j
    		End if
    		Next
    		if smallest_j <> i Then
    			For intA = 0 To ubound(Formular,1)
    				temp = Formular(intA,smallest_j)
    				Formular(intA,smallest_j) = Formular(intA,i)
    				Formular(intA,i) = temp
    			Next
    		End if
    	Next
    sortFormular = Formular
case ""
    sortFormular = Formular
case else
    Formular = split(replace(request("_Sort"),"fields:",""),";")
	for c = 0 to ubound(Formular)
	   redim preserve sortFormular(2,c)
	   sortFormular(0,c) = Formular(c)
	   sortFormular(1,c) = request(Formular(c))
	Next
end select

' ==============================================================================================================================
'                                   Erstelle den Mailtext                                                                       
' ==============================================================================================================================

	strTmp = ""
	for i = 1 to ubound(sortFormular,2)
	    if len(sortFormular(0,i)) > len(strTmp) then 
		   strTmp = sortFormular(0,i)
		end if
	next
	
	for i = 1 to ubound(sortFormular,2)
	    Daten = Daten & sortFormular(0,i)
	    Daten = Daten & string(len(strTmp)-len(sortFormular(0,i))," ")
	    Daten = Daten & " : "
	    Daten = Daten & sortFormular(1,i) & vbCRLF
	next
	
	Vorlage = vbCRLF  & Daten
	Vorlage = Vorlage & "Preis der Plaketten      : " & request("_Preis") & " CHF" & vbCRLF
	Vorlage = Vorlage & "---------------------------------------------------------"  & vbCRLF
	Vorlage = Vorlage & "IP Adresse des Absenders : " & Request.ServerVariables("REMOTE_ADDR") & vbCRLF
	Vorlage = Vorlage & "Datum beim Versenden     : " & date  & vbCRLF
	Vorlage = Vorlage & "Uhrzeit beim Versenden   : " & time  & vbCRLF
	Vorlage = Vorlage & "---------------------------------------------------------"  & vbCRLF
	Vorlage = Vorlage & "Formmailer 2.1 | Copyright Genotec Internet Consulting AG"  & vbCRLF
    Vorlage = Vorlage & "                 http://www.webhosting.ch                "  & vbCRLF

' ==============================================================================================================================
'                                   Sende das Mail                                                                              
' ==============================================================================================================================
Set Mail 		= Server.CreateObject("Persits.MailSender")
    Mail.Host 		= request("_Host")
    Mail.From 		= request("_MailFrom")
    Mail.FromName 	= request("_MailFromName")
    Mail.Subject    = request("_MailSubject")
    if instr(request("_EMail"),";") > 0  then
       MailAdressen = split(request("_EMail"),";")
       for i = 0 to Ubound(MailAdressen)
           Mail.AddAddress  trim(MailAdressen(i))
       next
    else
       Mail.AddAddress   request("_EMail")
    end if
    Mail.Body = Vorlage
    err.clear
    Mail.Send
    If Err = 0 Then
		Response.Cookies("preis")=request("_Preis")
		Response.Cookies("gold")=request("Goldplaketten")
		Response.Cookies("silber")=request("Silberplaketten")
		Response.redirect request("_DankePage")
    else
	   Response.redirect request("_FehlerPage")
    End If
Set Mail = Nothing

' ==============================================================================================================================
'                                          Ende                                                                                 
' ==============================================================================================================================

%>