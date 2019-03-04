Set args = WScript.Arguments
If args.Count = 2 Then
    url = args(0)
    target = args(1)
    Set xReq = createObject("Microsoft.XMLHTTP")
    xReq.Open "GET", url, 0
    xReq.Send()
    set fStream = createObject("ADODB.Stream")
    fStream.Mode = 3
    fStream.Type = 1
    fStream.Open()
    fStream.Write xReq.ResponseBody
    fStream.SaveToFile target, 2
End If
