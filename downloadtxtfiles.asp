<%@ LANGUAGE=VBScript %>

<% 

filename = Request.QueryString("filename")
typeoffile = Right(filename, 3)

If typeoffile = "txt" Then

	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	File = "C:\inetpub\wwwroot\Distribution\" + filename
	If FSO.FileExists(File) Then
		Set curfile = FSO.GetFile(File)
		If curfile.Size > 0 Then
			Set st = CreateObject("ADODB.Stream")
			st.Type = 1
			st.Open
			st.LoadFromFile File
			Response.BinaryWrite st.Read
			st.Close
		End If
	End If

End If

%>

