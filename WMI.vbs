' Header
Option Explicit
On Error Resume Next

Dim WmiComputer, WmiNamespace
Dim WmiAPI
Dim WmiQuery
Dim WmiResult, Record
Dim xmlDoc
Dim XmlNodes, XmlNode
Dim XmlNodeNameAttribute
Const MaxNameLength = 25

' Reference
If wScript.Arguments.Named.Exists("Computer") Then
	WmiComputer = wScript.Arguments.Named("Computer")
Else
	WmiComputer = "."
End If
If wScript.Arguments.Named.Exists("Namespace") Then
	WmiNamespace = wScript.Arguments.Named("Namespace")
Else
	WmiNamespace = "\root\cimv2"
End If
If wScript.Arguments.Named.Exists("Query") Then
	WmiQuery = wScript.Arguments.Named("Query")
Else
	wScript.Echo "Usage: cscript WMI.vbs [/computer:<computer>] [/namespace:<WMI_namespace>] <Query>"
	wScript.Echo "Example Query: Select * From Win32_Process"
	wScript.Quit
End If

Set xmlDoc = CreateObject("Microsoft.XMLDOM")

' Worker
Set WmiAPI = GetObject("winmgmts:\\" &  WmiComputer & WmiNamespace)
Set WmiResult = WmiAPI.ExecQuery(WmiQuery)

' Output
For Each Record In WmiResult
	XmlDoc.LoadXML(Record.GetText_(1))
	Set XmlNodes = XmlDoc.GetElementsByTagName("PROPERTY")
	For Each XmlNode in XmlNodes
		XmlNodeNameAttribute = XmlNode.GetAttribute("NAME")
		If Not Mid(XmlNodeNameAttribute, 1, 2) = "__" Then
			While Len(XmlNodeNameAttribute) < MaxNameLength
				XmlNodeNameAttribute = XmlNodeNameAttribute & " "
			Wend
			WScript.Echo XmlNodeNameAttribute & " = "  & XmlNode.Text
		End If
	Next
	wScript.echo("----------------------------------------------------------------------------")
Next

wScript.echo(vbNewLine&"Done")
