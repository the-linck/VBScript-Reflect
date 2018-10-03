<%
Option Explicit
Response.LCID = 1046 ' Brazilian LCID (use your locale code here).
' Cold also be the LCID property of the page declaration or Session.LCID to set it to the entire session.
%>
<!--#include file="jsonObject.class.asp" -->
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>ASPJSON</title>
	
	<style type="text/css">
		body {
			font-family: Arial, Helvetica, sans-serif;
		}
	
		pre {
			border: solid 1px #CCCCCC;
			background-color: #EEE;
			padding: 5px;
			text-indent: 0;
			width: 90%;
			white-space: pre-wrap;
			word-wrap: break-word;
		}
	</style>
</head>
<body>
	<h1>JSON Object and Array Tests</h1>
	<%
	server.ScriptTimeout = 10
	dim jsonObj, jsonString, jsonArr, outputObj
	dim testLoad, testAdd, testRemove, testValue, testChange, testArrayPush, testLoadRecordset
	dim testLoadArray, testChangeDefaultPropertyName, testGetItemAt
	
	testLoad = true
	testLoadArray = true
	testAdd = true
	testRemove = true
	testValue = true
	testChange = true
	
	testArrayPush = true
	
	testLoadRecordset = true
	
	testChangeDefaultPropertyName = true
	
	set jsonObj = new JSONobject
	set jsonArr = new jsonArray
	
	jsonObj.debug = false
	
	if testLoad then
		jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"
		
		if testLoadArray then jsonString = "[" & jsonString & "]"
		
		set outputObj = jsonObj.parse(jsonString)
		%>
		<h3>Parse Input</h3>
		<pre><%= jsonString %></pre>
		<%
	end if
	
	if testAdd then
		dim arr, multArr, nestedObject
		arr = Array(1, "teste", 234.56, "mais teste", "234", now)
		
		redim multArr(2, 3)
		multArr(0, 0) = "0,0"
		multArr(0, 1) = "0,1"
		multArr(0, 2) = "0,2"
		multArr(0, 3) = "0,3"

		multArr(1, 0) = "1,0"
		multArr(1, 1) = "1,1"
		multArr(1, 2) = "1,2"
		multArr(1, 3) = "1,3"
		
		multArr(2, 0) = "2,0"
		multArr(2, 1) = "2,1"
		multArr(2, 2) = "2,2"
		multArr(2, 3) = "2,3"
		
		jsonObj.add "nome", "Jozé"
		jsonObj.add "ficticio", true
		jsonObj.add "idade", 25
		jsonObj.add "saldo", -52
		jsonObj.add "bio", "Nascido em São Paulo\Brasil" & vbcrlf & "Sem filhos" & vbcrlf & vbtab & "Jogador de WoW"
		jsonObj.add "data", now
		jsonObj.add "lista", arr
		jsonObj.add "lista2", multArr
		
		set nestedObject = new JSONobject
		nestedObject.add "sub1", "value of sub1"
		nestedObject.add "sub2", "value of ""sub2"""
		
		jsonObj.add "nested", nestedObject
	end if
	
	if testRemove then
		jsonObj.remove "numbers"
		jsonObj.remove "aNonExistantPropertyName" ' this sould run silently, even to non existant properties
	end if
	
	if testValue then
		%><h3>Get Values</h3><%
		response.write "nome: " & jsonObj.value("nome") & "<br>"
		response.write "idade: " & jsonObj("idade") & "<br>" ' short syntax
		response.write "non existant property:" & jsonObj("aNonExistantPropertyName") & "(" & typeName(jsonObj("aNonExistantPropertyName")) & ")<br>"
	end if
	
	
	if testChange then
		%><h3>Change Values</h3><%
		
		response.write "nome before: " & jsonObj.value("nome") & "<br>"
		
		jsonObj.change "nome", "Mario"
		
		response.write "nome after: " & jsonObj.value("nome") & "<br>"
		
		jsonObj.change "nonExisting", -1
		
		response.write "Non existing property is created with: " & jsonObj.value("nonExisting") & "<br>"
	end if
	
	if testArrayPush then
		jsonArr.Push jsonObj
		jsonArr.Push 1
		jsonArr.Push "strings too"
	end if
	
	if testLoadRecordset then
		%><h3>Load a Recordset</h3>
		<!--
		   METADATA
		   TYPE="TypeLib"
		   NAME="Microsoft ActiveX Data Objects 2.5 Library"
		   UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
		   VERSION="2.5"
		-->
		<%
		dim rs
		set rs = createObject("ADODB.Recordset")
		
		' prepera an in memory recordset 
		' could be, and mostly, loaded from a database
		rs.CursorType = adOpenKeyset
		rs.CursorLocation = adUseClient
		rs.LockType = adLockOptimistic
		
		rs.Fields.Append "ID", adInteger, , adFldKeyColumn
		rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
		rs.Fields.Append "Valor", adDecimal, 14, adFldMayBeNull
		rs.Fields("Valor").NumericScale = 2
		
		rs.Open
		
		rs.AddNew
		rs("ID") = 1
		rs("Nome") = "Nome 1"
		rs("Valor") = 10.99
		rs.Update
		
		rs.AddNew
		rs("ID") = 2
		rs("Nome") = "Nome 2"
		rs("Valor") = 29.90
		rs.Update
		
		rs.moveFirst
		jsonObj.LoadFirstRecord rs
		' or
		rs.moveFirst
		jsonArr.LoadRecordSet rs
		
		rs.Close
		
		set rs = nothing
	end if	
	
	if testLoad then
		%>
		<h3>Parse Output</h3>
		<pre><%= outputObj.write %></pre>
		<%
	end if
	
	if testChangeDefaultPropertyName then
		jsonObj.defaultPropertyName = "CustomName"
		jsonArr.defaultPropertyName = "CustomArrName"
	end if
	%>
	
	<h3>JSON Object Output<% if testLoad then %> (Same as parse output: <% if typeName(jsonObj) = typeName(outputObj) then %>yes<% else %>no<% end if %>)<% end if %></h3>
	<pre><%= jsonObj.write %></pre>	
	
	<h3>Array Output</h3>
	<pre><%= jsonArr.write %></pre>
	
	<h3>Array Loop</h3>
	<pre><%
	dim i, items, item
	

	' more readable loop
	i = 0
	response.write "For Each Loop (readability):<br>==============<br>"
	
	for each item in jsonArr.items
		response.write "Index "
		response.write i
		response.write ": "
	
		if isObject(item) and typeName(item) = "JSONobject" then
			item.write()
		else
			response.write item
		end if
		
		response.write "<br>"
		i = i + 1
	next
	
	response.write "<br><br>For Loop (speed):<br>=========<br>"
	
	' faster but less readable
	for i = 0 to jsonArr.length - 1
		response.write "Index "
		response.write i
		response.write ": "
	
		if isObject(jsonArr(i)) then
			set item = jsonArr(i)
			
			if typeName(item) = "JSONobject" then
				item.write()
			else
				response.write item
			end if
		else
			item = jsonArr(i)
			response.write item
		end if
		
		response.write "<br>"
	next

	
	set outputObj = nothing
	set jsonObj = nothing
	set jsonArr = nothing
	%></pre>
	
	<h3>JSON Script Output</h3>
	
	<%
	
	dim realOutput
	dim expectedOutput
	
	dim javascriptCode
	dim javascriptkey
	
	dim jsonScr

	javascriptCode = "function() { alert('test'); }"
	javascriptKey = "script"
	
	expectedOutput = "{""" & javascriptKey & """:" & javascriptCode & "}"
	
	set jsonScr = new JSONscript
	jsonScr.value = javascriptCode
	
	set jsonObj = new JSONobject
	jsonObj.Add javascriptKey, jsonScr
	
	realOutput = jsonObj.Serialize()
	
	%><h4>Output<% if (realOutput = expectedOutput) then %> (correct)<% else %> (INCORRECT!)<% end if %></h4>
	<pre><%= realOutput %></pre>
	
</body>
</html>
