<%@ Language="VBScript" Codepage="65001" %>
<%
Option Explicit
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "text/html"
Response.Charset = "UTF-8"

' Get form data
Dim username, email, message
username = Trim(Request.Form("username"))
email = Trim(Request.Form("email"))
message = Trim(Request.Form("message"))

' Data validation
If username = "" Or email = "" Then
    Response.Write("<h2>Please fill in all required fields.</h2>")
    Response.Write("<a href='form.html'>Go back</a>")
    Response.End()
End If

' File path configuration
Dim filePath, fs
filePath = Server.MapPath("/data/user_data.txt") ' Path to save file

' Create filesystem object
On Error Resume Next
Set fs = Server.CreateObject("Scripting.FileSystemObject")

' Create stream object for UTF-8 encoding
Dim stream
Set stream = Server.CreateObject("ADODB.Stream")
stream.Type = 2 ' Text type
stream.Charset = "UTF-8"
stream.Open

' Read existing file content (if file exists)
If fs.FileExists(filePath) Then
    Dim tempFile
    Set tempFile = fs.OpenTextFile(filePath, 1, False, -1) ' -1 = TristateTrue (Unicode)
    stream.WriteText tempFile.ReadAll
    tempFile.Close
    Set tempFile = Nothing
End If

' Add new data
stream.WriteText "Name: " & username & vbCrLf
stream.WriteText "Email: " & email & vbCrLf
stream.WriteText "Message: " & message & vbCrLf
stream.WriteText "Submission Date: " & Now() & vbCrLf
stream.WriteText "-----------------------------" & vbCrLf

' Save file
stream.SaveToFile filePath, 2 ' 2 = adSaveCreateOverWrite
stream.Close
Set stream = Nothing

' Clean up objects
Set fs = Nothing

' Display results to user
%>
<!DOCTYPE html>
<html>
<head>
    <title>Submission Complete</title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
    <h2>Data has been successfully saved.</h2>
    <p><a href="form.html">Submit another form</a></p>
    
    <h3>Your submitted information:</h3>
    <p>Name: <%= Server.HTMLEncode(username) %></p>
    <p>Email: <%= Server.HTMLEncode(email) %></p>
    <p>Message: <%= Server.HTMLEncode(message) %></p>
</body>
</html>