set obj = CreateObject("ODIConverter.ODIConverter")

strRequest = _
	"<REQUEST " & _
		"USERID=""User"" " & _
		"UNIT=""Unit"" " & _
		"MACHINEID=""MACHINE"" " & _
		"CHANNELID=""0"" " & _
		"USERAUTHORITYLEVEL=""100"" " & _
		"ODIENVIRONMENT=""DEV"" " & _
		"OPERATION=""TRANSACTREQUEST"">" & _ 
		"<TRANSACTDATA TRANSACTNUMBER=""1234567890""/>" & _
	"</REQUEST>"


bLoop = true
nCount = 0
tmrStart = Timer
while bLoop and nCount < 1
	strResponse = obj.Request(strRequest)
	if Left(strResponse, 25) <> "<RESPONSE TYPE=""SUCCESS"">" then
		msgbox "Error: " + strResponse
		bLoop = false
	end if
	nCount = nCount + 1
wend
tmrEnd = Timer
tmrDiff = tmrEnd - tmrStart

if bLoop then
	msgbox tmrDiff & " seconds " & CHR(13) & CHR(10) & strRequest & CHR(13) & CHR(10) & strResponse
end if
