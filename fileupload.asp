<!--#include file="freeaspupload.asp"-->
<%
Set upload = New FreeASPUpload
upload.Save(Server.MapPath("data"))
Dim fieldName, file
For Each fieldName In upload.Files
  Set file = upload.Files.Item(fieldName)
  Response.Write "<h2>Upload Successful!</h2>"
  Response.Write "Saved File: " & file.FileName & "<br>"
Next
If upload.Files.Count = 0 Then
  Response.Write "<h3>No file uploaded.</h3>"
End If
%>
