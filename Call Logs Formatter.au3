#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Chris NeJame

 Script Function:
	Takes a CSV through a GUI, creates an HTML file with a table 
	containing each column, using the first row as the ehaders. 
	Then it styles it. Then it asks where to save the generated 
	file.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <File.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>

$message = "Please Select the desired call log file, and select ""Open"""

$chosenOpen = FileOpenDialog($message, @MyDocumentsDir & "\", "Comma Seperated Values (*.csv)", 1)

If @error Then
   Exit
EndIf

ProgressOn("Call Log Formatter", "Please Wait", "Generating department list...", -1, -1,16)

ProgressSet(1, "Opening the CSV...")
$opened = FileOpen($chosenOpen)

ProgressSet(10, "Reading the CSV...")
$openedContents = FileRead($opened)

ProgressSet(11, "Trimming the fat...")
$openedContents = StringRegExpReplace($openedContents, '"', "")

ProgressSet(20, "Locating date bounds...")
$date1 = StringRegExp($openedContents, '\A.*\n\s*([^,\n]+),', 1) ; start
$date2 = StringRegExp($openedContents, '\n\s*([^,\n]+)(?=(,.*\s*\Z))', 1) ; end

ProgressSet(22, "Locating unique department names...")
$departments = StringRegExp($openedContents, '(?s)([^,\n]+)(?=(,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*[\n$]))(?!(.*\1,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*,[^,\n]*[\n$]))', 3)

ProgressSet(95, "Compiling results...")
$numRegexResults = UBound($departments)
Local $uniqueDepts[1]
For $i = 2 To $numRegexResults-1 Step 2
   _ArrayAdd($uniqueDepts, $departments[$i])
Next

ProgressSet(95, "Sorting results...")
_ArraySort($uniqueDepts)


ProgressSet(100, "Complete!")
ProgressOff()

$timestamp = "Call Logs " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & "-" & @MIN & "-" & @SEC & ".html"
 
$chosenSave = FileSaveDialog("Save As", @MyDocumentsDir, "HTML (*.html)", 2 + 16, $timestamp)

If @error Then
   Exit
EndIf

ProgressOn("Call Log Formatter", "Please Wait", "Generating initial headers...", -1, -1, 16)

$html = _
"<!DOCTYPE HTML>" & @CRLF & _
"<html>" & @CRLF & _
"    <head>" & @CRLF & _
"        <title>Call logs</title>" & @CRLF & _
"        <style>" & @CRLF & _
"            @media screen, print{" & @CRLF & _
"                *{" & @CRLF & _
"                    padding: 0;" & @CRLF & _
"                    margin: 0;" & @CRLF & _
"                    font-family:Arial, Helvetica, sans-serif;" & @CRLF & _
"                }" & @CRLF & _
"                html, body{" & @CRLF & _
"                    height:100%;" & @CRLF & _
"                    margin: 0;" & @CRLF & _
"                }" & @CRLF & _
"                body, .page{" & @CRLF & _
"                    background-color: #b3d4fc;" & @CRLF & _
"                }" & @CRLF & _
"                .page{" & @CRLF & _
"                    width:100%;" & @CRLF & _
"                    height:100%;" & @CRLF & _
"                    text-align:center;" & @CRLF & _
"                }" & @CRLF & _
"                .sectionTitle {" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding-top:20px;" & @CRLF & _
"                    padding-bottom:20px;" & @CRLF & _
"                    border-bottom:1px solid black;" & @CRLF & _
"                }" & @CRLF & _
"                .sectionTitleSmall {" & @CRLF & _
"                    font-size:1.5em;" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding-top:20px;" & @CRLF & _
"                    text-align:center;" & @CRLF & _
"                    border-bottom:1px solid black;" & @CRLF & _
"                }" & @CRLF & _
"                .disclaimer {" & @CRLF & _
"                    font-size:.6em;" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding-bottom:10px;" & @CRLF & _
"                    border-bottom:1px solid black;" & @CRLF & _
"                    text-align:center;" & @CRLF & _
"                }" & @CRLF & _
"                .caveats {" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    text-align:left;" & @CRLF & _
"                }" & @CRLF & _
"                .cavTitle {" & @CRLF & _
"                    font-size:0.8em;" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding-left:20px;" & @CRLF & _
"                }" & @CRLF & _
"                #more-button {" & @CRLF & _
"                    border-style:none;" & @CRLF & _
"                    background:none;" & @CRLF & _
"                }" & @CRLF & _
"                #more-button {" & @CRLF & _
"                    border-style:none;" & @CRLF & _
"                    background:none;" & @CRLF & _
"                }" & @CRLF & _
"                .removedCol {" & @CRLF & _
"                    font-size:.6em;" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding-left:40px;" & @CRLF & _
"                }" & @CRLF & _
"                .sumTitle {" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    padding:20px;" & @CRLF & _
"                }" & @CRLF & _
"                .orgSum .sumTitle {" & @CRLF & _
"                    font-size:1.5em;" & @CRLF & _
"                }" & @CRLF & _
"                .deptSum .sumTitle {" & @CRLF & _
"                    font-size:1.1em;" & @CRLF & _
"                    padding:12px;" & @CRLF & _
"                }" & @CRLF & _
"                .pnSum .sumTitle {" & @CRLF & _
"                    font-size:1.0em;" & @CRLF & _
"                    padding:5px;" & @CRLF & _
"                }" & @CRLF & _
"                .deptSum .sumTitle p:hover, .pnSum .sumTitle p:hover, .tableTitle p:hover, .cavTitle p:hover {" & @CRLF & _
"                    cursor:pointer;" & @CRLF & _
"                    background-color:#b3d4fc;" & @CRLF & _
"                }" & @CRLF & _
"                .sumTable {" & @CRLF & _
"                    width:100%;" & @CRLF & _
"                }" & @CRLF & _
"                .deptSum .sumTableDiv, .pnSum .sumTableDiv {" & @CRLF & _
"                    padding-left:15px;" & @CRLF & _
"                    padding-right:15px;" & @CRLF & _
"                }" & @CRLF & _
"                .content {" & @CRLF & _
"                    padding-bottom:150px;" & @CRLF & _
"                }" & @CRLF & _
"                .logTable  td:first-child, .logTable  th:first-child, .orgSum td:first-child {" & @CRLF & _
"                    border-left:hidden;" & @CRLF & _
"                }" & @CRLF & _
"                .logTable  td:last-child, .logTable  th:last-child, .orgSum td:last-child {" & @CRLF & _
"                    border-right:hidden;" & @CRLF & _
"                }" & @CRLF & _
"                .tableTitle {" & @CRLF & _
"                    font-size:2.0em;" & @CRLF & _
"                    padding:20px;" & @CRLF & _
"                }" & @CRLF & _
"                table.sumTable td:nth-child(2) {" & @CRLF & _
"                    width:200px;" & @CRLF & _
"                }" & @CRLF & _
"                table{" & @CRLF & _
"                    border-collapse:collapse;" & @CRLF & _
"                    margin: 0 auto;" & @CRLF & _
"                    " & @CRLF & _
"                }" & @CRLF & _
"                table, th, td{" & @CRLF & _
"                    border:1px solid black;" & @CRLF & _
"                }" & @CRLF & _
"                th{" & @CRLF & _
"                    background-color: #B9CDE4;" & @CRLF & _
"                }" & @CRLF & _
"                tr:nth-child(odd){" & @CRLF & _
"                    background-color: #DDECFE;" & @CRLF & _
"                }" & @CRLF & _
"                tr:nth-child(even){" & @CRLF & _
"                    background-color: #CEE4FD;" & @CRLF & _
"                }" & @CRLF & _
"                th, td {" & @CRLF & _
"                    width:auto;" & @CRLF & _
"                    white-space:normal;" & @CRLF & _
"                    padding-left:2px;" & @CRLF & _
"                    padding-right:2px;" & @CRLF & _
"                    padding-top:2px;" & @CRLF & _
"                    padding-bottom:2px;" & @CRLF & _
"                }" & @CRLF & _
"                th.logTable, td.logtable {" & @CRLF & _
"                    max-width:105px;" & @CRLF & _
"                }" & @CRLF & _
"                td {" & @CRLF & _
"                    font-size: .9em;" & @CRLF & _
"                }" & @CRLF & _
"                .Date {" & @CRLF & _
"                    max-width:75px;" & @CRLF & _
"                }" & @CRLF & _
"                .Time{" & @CRLF & _
"                    max-width:60px;" & @CRLF & _
"                }" & @CRLF & _
"                .CallType {" & @CRLF & _
"                    max-width:75px;" & @CRLF & _
"                }" & @CRLF & _
"                .CallingNumber, .CalledNumber {" & @CRLF & _
"                    max-width:85px;" & @CRLF & _
"                    overflow: auto;" & @CRLF & _
"                }" & @CRLF & _
"                .CallingDepartment, .CalledDepartment {" & @CRLF & _
"                    max-width:105px;" & @CRLF & _
"                }" & @CRLF & _
"                .CallingExtension, .CalledExtension {" & @CRLF & _
"                    max-width:80px;" & @CRLF & _
"                }" & @CRLF & _
"                .CallConnected {" & @CRLF & _
"                    max-width:85px;" & @CRLF & _
"                }" & @CRLF & _
"                .Duration {" & @CRLF & _
"                    max-width:70px;" & @CRLF & _
"                }" & @CRLF & _
"                .QueuingTime {" & @CRLF & _
"                    max-width:70px;" & @CRLF & _
"                }" & @CRLF & _
"                .CarrierCode {" & @CRLF & _
"                    max-width:55px;" & @CRLF & _
"                }" & @CRLF & _
"                .summaries{" & @CRLF & _
"                    max-width:100%;" & @CRLF & _
"                    text-align:left;" & @CRLF & _
"                }" & @CRLF & _
"            }" & @CRLF & _
"            @media screen{" & @CRLF & _
"                .content{" & @CRLF & _
"                    display:inline-block;" & @CRLF & _
"                    min-height:100%;" & @CRLF & _
"                    min-width:900px;" & @CRLF & _
"                    max-width:100%;" & @CRLF & _
"                    background-color:white;" & @CRLF & _
"                    border-left:1px solid black;" & @CRLF & _
"                    border-right:1px solid black;" & @CRLF & _
"                }" & @CRLF & _
"                .sectionTitle {" & @CRLF & _
"                    font-size:2.5em;" & @CRLF & _
"                }" & @CRLF & _
"                .grow {" & @CRLF & _
"                    border-bottom:1px solid black;" & @CRLF & _
"                    -moz-transition: height .1s;" & @CRLF & _
"                    -ms-transition: height .1s;" & @CRLF & _
"                    -o-transition: height .1s;" & @CRLF & _
"                    -webkit-transition: height .1s;" & @CRLF & _
"                    transition: height .1s;" & @CRLF & _
"                    height: 0;" & @CRLF & _
"                    overflow: hidden;" & @CRLF & _
"                }" & @CRLF & _
"            }" & @CRLF & _
"            @media print{" & @CRLF & _
"                .sectionTitleSmall {" & @CRLF & _
"                    page-break-before:always;" & @CRLF & _
"                }" & @CRLF & _
"                .showHide {" & @CRLF & _
"                    display:none;" & @CRLF & _
"                }" & @CRLF & _
"                .dataTable {" & @CRLF & _
"                    display:none;" & @CRLF & _
"                }" & @CRLF & _
"                .caveats {" & @CRLF & _
"                    display:none;" & @CRLF & _
"                }" & @CRLF & _
"                .disclaimer {" & @CRLF & _
"                    display:none;" & @CRLF & _
"                }" & @CRLF & _
"                .sectionTitle {" & @CRLF & _
"                    font-size:2.0em;" & @CRLF & _
"                }" & @CRLF & _
"                .deptSum .sumTitle {" & @CRLF & _
"                    font-size:0.9em;" & @CRLF & _
"                }" & @CRLF & _
"                .orgSum .sumTitle {" & @CRLF & _
"                    font-size:1.2em;" & @CRLF & _
"                    padding:10px;" & @CRLF & _
"                }" & @CRLF & _
"                div.deptSum:nth-child(6n+6), div.pnSum:nth-child(7n+2) {" & @CRLF & _
"                    page-break-after:always;" & @CRLF & _
"                }" & @CRLF & _
"            }" & @CRLF & _
"        </style>" & @CRLF & _
"        <script>" & @CRLF & _
"            var browserName = ""not IE"";" & @CRLF & _
"            var fullVersion = ""doesn't matter"";" & @CRLF & _
"            function growDiv(event) {" & @CRLF & _
"                var target = event.target || event.srcElement;" & @CRLF & _
"                while(target.tagName.toUpperCase() != ""SPAN""){" & @CRLF & _
"                    target = target.children[0];" & @CRLF & _
"                }" & @CRLF & _
"                var span = target;" & @CRLF & _
"                while(target.tagName.toUpperCase() != ""DIV""){" & @CRLF & _
"                    target = target.parentNode;" & @CRLF & _
"                }" & @CRLF & _
"                var parentDiv = target.parentNode;" & @CRLF & _
"                if(browserName == ""Microsoft Internet Explorer"" && fullVersion < 10){"  & @CRLF & _
"                    var growDiv = parentDiv.querySelectorAll('.grow')[0];" & @CRLF & _
"                } else {" & @CRLF & _
"                    var growDiv = parentDiv.getElementsByClassName('grow')[0];" & @CRLF & _
"                }" & @CRLF & _
"                if (growDiv.clientHeight) {" & @CRLF & _
"                    growDiv.style.height = 0;" & @CRLF & _
"                } else {" & @CRLF & _
"                    growDiv.style.height = growDiv.children[0].clientHeight + ""px"";" & @CRLF & _
"                }" & @CRLF & _
"                span.innerHTML=span.innerHTML=='Show '?'Hide ':'Show ';" & @CRLF & _
"            }" & @CRLF & _
"            function sortTable(event, column){" & @CRLF & _
"                var table = event.target;" & @CRLF & _
"                while(table.tagName.toUpperCase() != 'TABLE'){" & @CRLF & _
"                    table = table.parentNode;" & @CRLF & _
"                }" & @CRLF & _
"                var tbl = table.tBodies[0];" & @CRLF & _
"                var column = event.target.cellIndex;" & @CRLF & _
"                var store = [];" & @CRLF & _
"                for(var i=0, len=tbl.rows.length; i<len; i++){" & @CRLF & _
"                    var row = tbl.rows[i];" & @CRLF & _
"                    var sortnr = parseFloat(row.cells[column].textContent || row.cells[column].innerText);" & @CRLF & _
"                    if(!isNaN(sortnr)) store.push([sortnr, row]);" & @CRLF & _
"                }" & @CRLF & _
"                store.sort();" & @CRLF & _
"                for(var i=0, len=store.length; i<len; i++){" & @CRLF & _
"                    tbl.appendChild(store[i][1]);" & @CRLF & _
"                }" & @CRLF & _
"                store = null;" & @CRLF & _
"            }" & @CRLF & _
"            window.onload=function(){" & @CRLF & _
"                var nAgt = navigator.userAgent;" & @CRLF & _
"                if ((verOffset=nAgt.indexOf(""MSIE""))!=-1) {" & @CRLF & _
"                    browserName = ""Microsoft Internet Explorer"";" & @CRLF & _
"                    fullVersion = nAgt.substring(verOffset+5);" & @CRLF & _
"                    fullVersion = fullVersion=fullVersion.substring(0,fullVersion.indexOf("";""));" & @CRLF & _
"                    fullVersion = parseFloat(fullVersion);" & @CRLF & _
"                }" & @CRLF & _
"            }" & @CRLF & _
"        </script>" & @CRLF & _
"    </head>" & @CRLF & _
"    <body>" & @CRLF & _
"        <div class=""page"">" & @CRLF & _
"            <div class=""content"">" & @CRLF & _
"                <div class=""sectionTitle"">" & @CRLF & _
"                    <p>Call Logs " & $date1[0] & "-" & $date2[0] & "*</p>" & @CRLF & _
"                </div>" & @CRLF & _
"                <div class=""disclaimer"">" & @CRLF & _
"                    <p>*WARNING: Some columns may be removed if they contain no entries.</p>" & @CRLF & _
"                    <p>Some calls do not have associated departments due to no department being tied to them </p>" & @CRLF & _
"                    <p>when the call was made. They are not included in their respective department, or</p>" & @CRLF & _
"                    <p>oranization totals. But the calls are still summarized and in the data table.</p>" & @CRLF & _
"                    <p>This issue only effects calls made before 9/26/2013.</p>" & @CRLF & _
"                </div>" & @CRLF & _
"<REPLACEME>" & @CRLF
$htmlTable = _
"                <div class=""dataTable"">" & @CRLF & _
"                    <div class=""tableTitle"">" & @CRLF & _
"                        <p onclick=""growDiv(event)""><span class=""showHide"">Show </span>Data Table</p>" & @CRLF & _
"                    </div>" & @CRLF & _
"                    <div class=""logTable grow"">" & @CRLF & _
"                        <table>" & @CRLF & _
"                            <tr>" & @CRLF
ProgressSet(1, "Reading the CSV...")
Local $fileLines[1]
If Not _FileReadToArray($chosenOpen, $fileLines) Then
   MsgBox(0, "Error", " Contact the Help Desk:" & @error)
   Exit
EndIf
ProgressSet(3, "Generating table headers...")
; set the headers for the table
$curLineArry = StringSplit($fileLines[1], ",")

$columnNum = $curLineArry[0]
Local $headers[1]
For $i = 1 To $columnNum
   _ArrayAdd($headers, StringRegExpReplace($curLineArry[$i], '\s*', ''))
   $htmlTable &= _
"                                <th class=""" & StringRegExpReplace($curLineArry[$i], '\s*', '') & """>" & $curLineArry[$i] & "</th>" & @CRLF
Next
$htmlTable &= _
"                            </tr>" & @CRLF
$rowNum = $fileLines[0]

; track each column to make sure it is not empty.
; Each column is shown to be empty by default.
; A column must prove it has something in it.
; If it doesn't it remains True
; If it does, it becomes False
Local $colEmptyArry[$columnNum]
Local $colEmptyArry[$columnNum]
For $i = 0 To $columnNum-1
   $colEmptyArry[$i] = True
Next

ProgressSet(5, "Generating table rows...")
; track Call Duration, Queuing TIme, and Number of calls total 

Local $cDqTnC[UBound($uniqueDepts)-1][4];  cD = call duration(seconds), qT = queuing time(seconds), nC = number of calls(integer)
; multidimensional array to contain these values for each department ["dept"][cD][qT][nC]

; Each department starts out with 0 call duration seconds accumulated, 0 queuing time seconds accumulated, and 0 calls
For $i=0 To UBound($uniqueDepts)-2
   $cDqTnC[$i][0] = $uniqueDepts[$i+1] ; department names
   $cDqTnC[$i][1] = 0 ; default call durations
   $cDqTnC[$i][2] = 0 ; default queuing times
   $cDqTnC[$i][3] = 0 ; default number of calls
Next

; add extra for unknown
ReDim $cDqTnc[UBound($cDqTnc)+1][UBound($cDqTnc,2)]
$cDqTnC[UBound($cDqTnc)-1][0] = "Unknown" ; department name
$cDqTnC[UBound($cDqTnc)-1][1] = 0 ; default call durations
$cDqTnC[UBound($cDqTnc)-1][2] = 0 ; default queuing times
$cDqTnC[UBound($cDqTnc)-1][3] = 0 ; default number of calls

; Each phone number starts out with 0 call duration seconds accumulated, 0 queuing time seconds accumulated, and 0 calls
Local $numsCdQtNc[1][4] ; individual phone number calls with no department

For $i = 2 To $rowNum
   ; For each row
   ; strip double quotes, and unneccesary spaces
   $fileLines[$i] = StringRegExpReplace($fileLines[$i], "\s*,\s*", ",")
   $fileLines[$i] = StringRegExpReplace($fileLines[$i], "\n\s*", "\n")
   $fileLines[$i] = StringRegExpReplace($fileLines[$i], """", "")
   ; split the current Lines text into an array
   $curLineArry = StringSplit($fileLines[$i], ",")
   $percentage = Ceiling(((($i - 1) / ($rowNum - 1))*85) +5)
   ProgressSet($percentage, "Checking row #" & ($i-1) & "...")
   ; get the department of the current line
   $lineDept = $curLineArry[6]
   ; get the phone number of the current line
   $linePhoneNum = $curLineArry[4]
   ; get the call type of the current line
   $lineCallType = $curLineArry[3]
   If Not($lineCallType=="Terminating") Then
	  $fileLines[$i] = StringRegExpReplace($fileLines[$i], "\s*,\s*", ",")
	  If $lineDept == "" Then $lineDept="Unknown" ;  If $lineDept isn't in the list of departments contained in $cDqtnC, the department is unkown, so set$lienDept to "Unknown"
	  $dti = _ArraySearch($cDqtnC, $lineDept) ; index of the row within $cDqtnC that corresponds with the department of the current line
	  
	  ; store time strings
	  $duration = $curLineArry[11]
	  $queuingTime = $curLineArry[12]
	  
	  ; format time string into seconds for call duration
	  If Not(StringLen($duration) > 0) Then
		 $duration = 0
	  Else
		 $duration = StringRegExp($duration, "(\s*\d+\s*):(\s*\d+\s*):(\s*\d+\s*)",3)
		 $duration = Int((Int($duration[0])*60*60)+(Int($duration[1])*60)+(Int($duration[2])))
	  EndIf
	  ; and now for queuing time as well
	  If Not(StringLen($queuingTime) > 0) Then
		 $queuingTime = 0
	  Else
		 $queuingTime = StringRegExp($queuingTime, "(\s*\d+\s*):(\s*\d+\s*):(\s*\d+\s*)",3)
		 $queuingTime = Int((Int($queuingTime[0])*60*60)+(Int($queuingTime[1])*60)+(Int($queuingTime[2])))
	  EndIf
	  
	  ; add these times to their previous values, and increment the call number(for individual phone numbers)
	  ; But only if the current call has no calling department.
	  If $lineDept == "Unknown" Then
		 $pni = _ArraySearch($numsCdQtNc, $linePhoneNum)
		 If $pni == -1 Or Not($linePhoneNum == $numsCdQtNc[$pni][0]) Then
			ReDim $numsCdQtNc[UBound($numsCdQtNc)+1][UBound($numsCdQtNc,2)]
			$pni = UBound($numsCdQtNc)-1
			$numsCdQtNc[UBound($numsCdQtNc)-1][0] = $linePhoneNum ; phone number of the current line
			$numsCdQtNc[UBound($numsCdQtNc)-1][1] = 0 ; default call durations
			$numsCdQtNc[UBound($numsCdQtNc)-1][2] = 0 ; default queuing times
			$numsCdQtNc[UBound($numsCdQtNc)-1][3] = 0 ; default number of calls
			
		 EndIf
		 $numsCdQtNc[$pni][1] = $numsCdQtNc[$pni][1] + $duration
		 $numsCdQtNc[$pni][2] = $numsCdQtNc[$pni][2] + $queuingTime
		 $numsCdQtNc[$pni][3] = $numsCdQtNc[$pni][3] + 1
	  EndIf
	  
	  ; add these times to their previous values, and increment the call number(for individual departments)
	  $cDqTnC[$dti][1] = $cDqTnC[$dti][1] + $duration
	  $cDqTnC[$dti][2] = $cDqTnC[$dti][2] + $queuingTime
	  $cDqTnC[$dti][3] = $cDqTnC[$dti][3] + 1
	  
	  $htmlTable &= _
   "                        <tr>" & @CRLF
	  For $j = 1 To $columnNum
		 ; for each column
		 ; if the column shows a value, make sure $colEmptyArry is updated
		 If StringLen($curLineArry[$j]) > 0 AND $colEmptyArry[$j-1] == True Then 
			$colEmptyArry[$j-1] = False 
		 EndIf
		 $htmlTable &= _
   "                            <td class=""" & $headers[$j] & """>" & $curLineArry[$j] & "</td>" & @CRLF
	  Next
	  $htmlTable &= _
   "                        </tr>" & @CRLF
	  $percentage = Ceiling(((($i - 1) / ($rowNum - 1))*85) +5)
	  ProgressSet($percentage, "Using row #" & ($i-1) & "...")
   EndIf
Next

_ArrayDelete($numsCdQtNc,0)
; Clear unwanted columns from the $html string's table
ProgressSet(87, "Clearing Empty columns...")

; Clear empty columns
$voidPresent = False
Local $removedColumns[1]
For $i = $columnNum-1 To 0 Step -1
   If $colEmptyArry[$i] == True Then
	  $voidPresent = True
	  $displacement = $columnNum-1 - $i
	  $removePattern = '(\n.*<t[hd][^>]*>[^<]*</t[hd]>)(?=(\s*'
	  $headerPattern = '<th[^>]*>([^<]*)</th>\s*'
	  For $j = 1 To $displacement
		 $removePattern &= '<t[hd][^>]*>[^<]*</t[hd]>\s*'
		 $headerPattern &= '<th[^>]*>[^<]*</th>\s*'
	  Next
	  $removePattern &= '</tr))'
	  $headerPattern &= '</tr'
	  $result = StringRegExp($htmlTable, $headerPattern, 3)
	  _ArrayAdd($removedColumns, $result[0])
	  $htmlTable = StringRegExpReplace($htmlTable, $removePattern, '')
	  $columnNum = $columnNum - 1
   EndIf
Next


; if any are empty create 'caveats' div. if not, remove the <REPLACEME> line
$replace = ""
If $voidPresent == True Then
   $replace &= _
"                <div class=""caveats"">" & @CRLF & _
"                    <div class=""cavTitle"">" & @CRLF & _
"                        <p onclick=""growDiv(event)"" ><span class=""showHide"">Show </span>removed columns</p>" & @CRLF & _
"                    </div>" & @CRLF & _
"                    <div class=""removedCols grow"">" & @CRLF & _
"                        <div class=""removedCol"">" & @CRLF
   For $i = 1 To UBound($removedColumns)-1
	  $replace &= _
"                            <div>" & @CRLF & _
"                                <p>" & $removedColumns[$i] & "</p>" & @CRLF & _
"                            </div>" & @CRLF
   Next
   $replace &= _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF & _
"                </div>" & @CRLF
EndIf

$html = StringRegExpReplace($html, '<REPLACEME>.*\n', $replace)


ProgressSet(90, "Generating closing tags...")
$htmlTable &= _
"                        </table>" & @CRLF & _
"                    </div>" & @CRLF

; $htmlTable should now be completely filled and dealt with

; create a summary for the logs
$html &= _
"                <div class=""sectionTitle"">" & @CRLF & _
"                    <p>Summary</p>" & @CRLF & _
"                </div>" & @CRLF & _
"                <div class=""summaries"">" & @CRLF

; individual department summaries
Local $exceptions[3][2] = [["COUNCIL","Div1"], _
					 ["Outreach","Div2"], _
					 ["Dept D","Div1"]] ; used to track departments that do not adhere to normal naming schemes, and their organization can not be determined through their name. If they do not start with Div1, or Div2, they are assumed to be part of Corporate, unnless they are in teh exceptions list, with their respective department

Local $htmlTemp ; temporarily store this section, so the organization totals

Local $orgCdQtNc[3][4] = [["Corporate",0,0,0], _
						     ["Div1",0,0,0], _
							 ["Div2",0,0,0]] ; Used to track total call durations, queuing times, and number of calls for each organization.

; format the summaries for each department to raw values, and add their values to their respective organization's totals
Local $excIndex = -1
For $i = 0 To UBound($cDqTnC)-1
   $cdS = $cDqTnC[$i][1]
   $qtS = $cDqTnC[$i][2]
   $orgName = ""
   $excIndex = _ArraySearch($exceptions, $cDqTnC[$i][0])
   If Not($excIndex == -1) Then ; if the $exceptions array has the current departments name in it, change the current organization name to the exception
	  $orgName = $exceptions[$excIndex][1]
   Else ; Else, check the first four characters of the department name
	  Switch StringMid($cDqTnC[$i][0],1,4)
		 Case "Div1" ; If it is "Div1" set $orgName to "Div1"
			$orgName = "Div1"
		 Case "Div2" ; If it is "Div2" set $orgName to "Div2"
			$orgName = "Div2"
		 Case Else ; Else, change it to Corporate
			$orgName = "Corporate"
	  EndSwitch
   EndIf
   ; If the departments name is an exception, $orgName will be set to the exception's corresponding organization.
   ; If not, it htis the Else, which leads to the Switch.
   ; If it hits the switch, it must have either "Div1" or "Div2" as it's first four letters.
   ; If "Div1" it set $orgName to "Div1"
   ; If "Div2" it set $orgName to "Div2"
   ; If not, it is assumed to be "Corporate".
   ; The only way it can be "Corporate" is if it isn't an exception, nor does it start with "Div1" or "Div2"
   ; It also MUST be "Corporate" if it is not an exception, and it does not start with "Div1", and it does not start with "Div2"
   
   $orgIndex = _ArraySearch($orgCdQtNc, $orgName)
   $orgCdQtNc[$orgIndex][1] = $orgCdQtNc[$orgIndex][1] + $cdS ; add current departments call duration to organizations call duration
   $orgCdQtNc[$orgIndex][2] = $orgCdQtNc[$orgIndex][2] + $qtS ; add current departments call duration to organizations call duration
   $orgCdQtNc[$orgIndex][3] = $orgCdQtNc[$orgIndex][3] + $cDqTnC[$i][3] ; add current departments call duration to organizations call duration
   $cdF = (($cdS-Mod($cdS,(60*60)))/(60*60)) & "h " & ((Mod($cdS,(60*60))-Mod($cdS,(60)))/(60)) & "m " & Mod($cdS,(60)) & "s"
   $qtF = (($qtS-Mod($qtS,(60*60)))/(60*60)) & "h " & ((Mod($qtS,(60*60))-Mod($qtS,(60)))/(60)) & "m " & Mod($qtS,(60)) & "s"
   $htmlTemp &= _
"                    <div id=""" & StringRegExpReplace($cDqTnC[$i][0], '\s*', '') & "Sum"" class=""deptSum"">" & @CRLF & _
"                        <div class=""sumTitle"">" & @CRLF & _
"                            <p onclick=""growDiv(event)""><span class=""showHide"">Show </span>" & $cDqTnC[$i][0] & "</p>" & @CRLF & _
"                        </div>" & @CRLF & _
"                        <div class=""sumTableDiv grow"">" & @CRLF & _
"                            <table class=""sumTable"">" & @CRLF & _
"                                <tr><td>Total Calls Duration:</td><td>" & $cdF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Queue Duration:</td><td>" & $qtF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Calls Made:</td><td>" & $cDqTnC[$i][3] & " calls</td></tr>" & @CRLF & _
"                            </table>" & @CRLF & _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF
Next ; will loop through for each department, adding their summaries to $htmlTemp. it also should now have the total for each organization stored in $orgCdQtNc

; the last row in $cDqTnC was the "Unknown" department
$unknownCdF = $cdF
$unknownQtF = $qtF
$unknownNc = $cDqTnC[UBound($cDqTnC)-1][3]

Local $htmlTempPhoneNums
For $i = 0 to UBound($numsCdQtNc)-1 ; Organization values

; store values of each column for the current row for faster access
   $numsCdS = $numsCdQtNc[$i][1]
   $numsQtS = $numsCdQtNc[$i][2]
   
   ; format the seconds of the current organizations times, into a readable format (Xh Ym Zs)
   $numsCdF = (($numsCdS-Mod($numsCdS,(60*60)))/(60*60)) & "h " & ((Mod($numsCdS,(60*60))-Mod($numsCdS,(60)))/(60)) & "m " & Mod($numsCdS,(60)) & "s"
   $numsQtF = (($numsQtS-Mod($numsQtS,(60*60)))/(60*60)) & "h " & ((Mod($numsQtS,(60*60))-Mod($numsQtS,(60)))/(60)) & "m " & Mod($numsQtS,(60)) & "s"
   
   ; append the div to $html for the current organization
   $htmlTempPhoneNums &= _
"                    <div id=""" & StringRegExpReplace($numsCdQtNc[$i][0], '\s*', '') & "Sum"" class=""pnSum"">" & @CRLF & _
"                        <div class=""sumTitle"">" & @CRLF & _
"                            <p onclick=""growDiv(event)""><span class=""showHide"">Show </span>" & $numsCdQtNc[$i][0] & "</p>" & @CRLF & _
"                        </div>" & @CRLF & _
"                        <div class=""sumTableDiv grow"">" & @CRLF & _
"                            <table class=""sumTable"">" & @CRLF & _
"                                <tr><td>Total Calls Duration:</td><td>" & $numsCdF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Queue Duration:</td><td>" & $numsQtF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Calls Made:</td><td>" & $numsCdQtNc[$i][3] & " calls</td></tr>" & @CRLF & _
"                            </table>" & @CRLF & _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF
Next

; add a summary div for each organization.

$tcd = 0
$tqt = 0
$tcn = 0
For $i = 0 to UBound($orgCdQtNc)-1 ; Organization values

; store values of each column for the current row for faster access
   $orgCdS = $orgCdQtNc[$i][1]
   $orgQtS = $orgCdQtNc[$i][2]
   
   ; add the values of the current row to the total sums
   $tcd += $orgCdQtNc[$i][1]
   $tqt += $orgCdQtNc[$i][2]
   $tcn += $orgCdQtNc[$i][3]
   
   ; format the seconds of the current organizations times, into a readable format (Xh Ym Zs)
   $orgCdF = (($orgCdS-Mod($orgCdS,(60*60)))/(60*60)) & "h " & ((Mod($orgCdS,(60*60))-Mod($orgCdS,(60)))/(60)) & "m " & Mod($orgCdS,(60)) & "s"
   $orgQtF = (($orgQtS-Mod($orgQtS,(60*60)))/(60*60)) & "h " & ((Mod($orgQtS,(60*60))-Mod($orgQtS,(60)))/(60)) & "m " & Mod($orgQtS,(60)) & "s"
   
   ; append the div to $html for the current organization
   $html &= _
"                    <div id=""" & StringRegExpReplace($orgCdQtNc[$i][0], '\s*', '') & "Sum"" class=""orgSum"">" & @CRLF & _
"                        <div class=""sumTitle"">" & @CRLF & _
"                            <p>" & $orgCdQtNc[$i][0] & "</p>" & @CRLF & _
"                        </div>" & @CRLF & _
"                        <div class=""sumTableDiv"">" & @CRLF & _
"                            <table class=""sumTable"">" & @CRLF & _
"                                <tr><td>Total Calls Duration:</td><td>" & $orgCdF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Queue Duration:</td><td>" & $orgQtF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Calls Made:</td><td>" & $orgCdQtNc[$i][3] & " calls</td></tr>" & @CRLF & _
"                            </table>" & @CRLF & _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF
Next

; format the seconds of the current total times, into a readable format (Xh Ym Zs)
$tcd = (($tcd-Mod($tcd,(60*60)))/(60*60)) & "h " & ((Mod($tcd,(60*60))-Mod($tcd,(60)))/(60)) & "m " & Mod($tcd,(60)) & "s"
$tqt = (($tqt-Mod($tqt,(60*60)))/(60*60)) & "h " & ((Mod($tqt,(60*60))-Mod($tqt,(60)))/(60)) & "m " & Mod($tqt,(60)) & "s"

; append the div to $html for the total sums
$html &= _
"                    <div id=""ADSum"" class=""orgSum"">" & @CRLF & _
"                        <div class=""sumTitle"">" & @CRLF & _
"                            <p>All Departments</p>" & @CRLF & _
"                        </div>" & @CRLF & _
"                        <div class=""sumTableDiv"">" & @CRLF & _
"                            <table class=""sumTable"">" & @CRLF & _
"                                <tr><td>Total Calls Duration:</td><td>" & $tcd & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Queue Duration:</td><td>" & $tqt & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Calls Made:</td><td>" & $tcn & " calls</td></tr>" & @CRLF & _
"                            </table>" & @CRLF & _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF

   $html &= _
"                    <div id=""UnknownSum"" class=""orgSum"">" & @CRLF & _
"                        <div class=""sumTitle"">" & @CRLF & _
"                            <p>Unassociated Calls</p>" & @CRLF & _
"                        </div>" & @CRLF & _
"                        <div class=""sumTableDiv"">" & @CRLF & _
"                            <table class=""sumTable"">" & @CRLF & _
"                                <tr><td>Total Calls Duration:</td><td>" & $unknownCdF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Queue Duration:</td><td>" & $unknownQtF & "</td></tr>" & @CRLF & _
"                                <tr><td>Total Calls Made:</td><td>" & $unknownNc & " calls</td></tr>" & @CRLF & _
"                            </table>" & @CRLF & _
"                        </div>" & @CRLF & _
"                    </div>" & @CRLF




$html &= _
"                    <div class=""sectionTitleSmall"">" & @CRLF & _
"                        <p>Departments</p>" & @CRLF & _
"                    </div>" & @CRLF

 ; $html should now have the divs for the Organization sums, and the total sums

; append to $html the divs for the sums for the individual departments that were stored in $htmlTemp
$html &= $htmlTemp

; if there were no unassociated phone numbers, don't add that part of the HTMl to it
If UBound($numsCdQtNc) > 0 Then
   $html &= _
"                    <div class=""sectionTitleSmall"">" & @CRLF & _
"                        <p>Unassociated Calls*</p>" & @CRLF & _
"                    </div>" & @CRLF & _
"                    <div class=""disclaimer"">" & @CRLF & _
"                        <p>*Only counts the calls made with numbers that had no associated department at the time of the call.</p>" & @CRLF & _
"                    </div>" & @CRLF

; append to $html the divs for the sums for the individual phone numbers that were stored in $htmlTempPhoneNums
   $html &= $htmlTempPhoneNums
EndIf

$html &= _
"                </div>" & @CRLF
; append to $html the table-related html from $htmlTable
; $html &= $htmlTable

; append the remaining closing tags to $html
$html &= _
"            </div>" & @CRLF & _
"        </div>" & @CRLF & _
"    </body>" & @CRLF & _
"</html>" & @CRLF

; $html should now contain a full, W3C compliant web page,
; starting with the head, including the title, CSS, and JavaScript
; Then the body, which has the title, a disclaimer, the caveats, 
; the summaries for the organizations, their totals, and each individual departments,
; and then the table of all the information

; write it out to the file, flush it, and close it
$saveTo = FileOpen($chosenSave, 2)
ProgressSet(91, "Writing to file...")
FileWrite($saveTo, $html)
ProgressSet(93, "Flushing to disk...")
FileFlush($saveTo)
ProgressSet(99, "Closing file...")
FileClose($saveTo)
ProgressSet(100, "Complete!")
ProgressOff()
MsgBox(0, "Compete", "Processing is complete. Your file is saved at: " & @CRLF & $chosenSave)