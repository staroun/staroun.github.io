<%
' Set variables from the form POST data
Dim username, password, fullname, gender, birthdate, phone, nationality, address, jobTypes, i
username = Request.Form("username")
password = Request.Form("password")
fullname = Request.Form("name")
gender = Request.Form("gender")
birthdate = Request.Form("birthdate")
phone = Request.Form("phone")
nationality = Request.Form("nationality")
address = Request.Form("address")

' job_type[] is an array if multiple checkboxes are checked
jobTypes = Request.Form("job_type")
Dim jobStr
jobStr = ""
If IsArray(jobTypes) Then
  For i = 0 To UBound(jobTypes)
    jobStr = jobStr & jobTypes(i)
    If i < UBound(jobTypes) Then
      jobStr = jobStr & ", "
    End If
  Next
ElseIf Len(Trim(jobTypes)) > 0 Then
  jobStr = jobTypes ' Only one selected
Else
  jobStr = "None Selected"
End If
' Create or open the file for appending
Dim fso, file, filePath
Set fso = Server.CreateObject("Scripting.FileSystemObject")
filePath = Server.MapPath("data/data.txt")
If Not fso.FolderExists(Server.MapPath("data")) Then
  fso.CreateFolder(Server.MapPath("data"))
End If
Set file = fso.OpenTextFile(filePath, 8, True) ' 8 = ForAppending, True = create if not exists
' Write the data
file.WriteLine "Username: " & username
file.WriteLine "Password: " & password
file.WriteLine "Full Name: " & fullname
file.WriteLine "Gender: " & gender
file.WriteLine "Date of Birth: " & birthdate
file.WriteLine "Phone Number: " & phone
file.WriteLine "Nationality: " & nationality
file.WriteLine "Address: " & address
file.WriteLine "Preferred Job Types: " & jobStr
file.WriteLine "-------------------------------------"
file.Close
Set file = Nothing
Set fso = Nothing
' Redirect or give a message
Response.Write "<h3>Registration saved successfully!</h3>"
Response.Write "<a href='registration.html'>Go Back</a>"
%>
