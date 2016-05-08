<%@ LANGUAGE=VBScript %>

<% 

filename = Request.QueryString("filename")
typeoffile = Right(filename, 3)

If typeoffile = "txt" Then

	If Request.TotalBytes > 0 Then

		filecontext = Request.BinaryRead(Request.TotalBytes)

		Set st = CreateObject("ADODB.Stream")
		st.Type = 1
		st.Open
		st.Write filecontext
		st.SaveToFile "C:\inetpub\wwwroot\Distribution\" + filename, 2
		st.Close

	End If

End If

%>

