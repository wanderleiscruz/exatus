Function UploadFile( url, file, fileType )
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
	Dim boundary
	boundary = "12345678"
	Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.Type = 1 ' binary
		.Open
		.LoadFromFile(file)
	End With
	Set sOut = CreateObject("ADODB.Stream")
	With sOut
		.Charset = 	"us-ascii"
		.Type = 2 ' Text!
		.Open
		.WriteText "--" & boundary & vbCrLf
		.WriteText "Content-Disposition: form-data; name=caminho"& vbCrLf  & vbCrLf
		.WriteText ";;wpastaweb;:" & vbCrLf
		.WriteText "--" & boundary & vbCrLf
		.WriteText "Content-Disposition: form-data; name=churrasco; filename=" & file & vbCrLf
		.WriteText "Content-Type: " & fileType & vbCrLf  & vbCrLf
	End With
	Set sOut2 = CreateObject("ADODB.Stream")
	With sOut2
		.Charset = "us-ascii"
		.Type = 2 ' Text
		.Open
		.WriteText vbCrLf  & "--" & boundary & "--" & vbCrLf  & vbCrLf
	End With
	Set sAll = CreateObject("ADODB.Stream")
	sAll.Type = 1 'binary
	sAll.Open
	sOut.Position = 0
	sOut.CopyTo sAll
	objStream.CopyTo sAll
	sOut2.Position = 0
	sOut2.CopyTo sAll
	oXMLHTTP.Open "POST", url, False
	oXMLHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
	oXMLHTTP.setRequestHeader "Content-length", sAll.Size
	sAll.Position = 0
	oXMLHTTP.Send sAll.Read()
	Set objStrem = Nothing
	Set sOut = Nothing
	Set sOut2 = Nothing
	Set sAll = Nothing
	Set oXMLHTTP = Nothing
End Function

Function BinaryToString(Binary)
    With CreateObject("ADODB.Stream")
        .Mode = 3
        .Type = 1
        .Open
        .Write Binary
        .Position = 0
        .Type = 2
        .Charset = "us-ascii"
        BinaryToString = .ReadText
    End With
End Function

UploadFile ";;link;:", ";;nomearq;:", "application/pdf"
