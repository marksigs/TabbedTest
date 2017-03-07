set obj = CreateObject("DmsCompression.dmsCompression1")

if true then

	bLoop = true
	nCount = 0
	tmrCOMPAPIDiff = 0
	tmrZLIBDiff = 0
	while bLoop and nCount < 1
		tmrZLIBStart = Timer
		obj.CompressionMethod = "ZLIB"
		obj.FileCompressToFile "C:\DOCS\TAD.DOC", "C:\DOCS\TADZLIB.ZIP"
		obj.FileDecompressToFile "C:\DOCS\TADZLIB.ZIP", "C:\DOCS\TADZLIB.DOC"
		tmrZLIBEnd = Timer
		tmrZLIBDiff = tmrZLIBDiff + (tmrZLIBEnd - tmrZLIBStart)

		tmrCOMPAPIStart = Timer
		obj.CompressionMethod = "COMPAPI"
		obj.FileCompressToFile "C:\DOCS\TAD.DOC", "C:\DOCS\TADCOMPAPI.ZIP"
		obj.FileDecompressToFile "C:\DOCS\TADCOMPAPI.ZIP", "C:\DOCS\TADCOMPAPI.DOC"
		tmrCOMPAPIEnd = Timer
		tmrCOMPAPIDiff = tmrCOMPAPIDiff + (tmrCOMPAPIEnd - tmrCOMPAPIStart)

		nCount = nCount + 1
	wend

	msgbox nCount & " iterations" & CHR(13) & CHR(10) & "COMP: " & CHR(9) & tmrCOMPAPIDiff & " seconds " & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & tmrZLIBDiff & " seconds"

end if

if false then

	obj.CompressionMethod = "COMPAPI"
	strOriginal = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
	strDecompressedCOMPAPI = obj.SafeArrayDecompressToBSTR(obj.BSTRCompressToSafeArray(strOriginal))
	obj.CompressionMethod = "ZLIB"
	strDecompressedZLIB = obj.SafeArrayDecompressToBSTR(obj.BSTRCompressToSafeArray(strOriginal))
	msgbox "Old: " & CHR(9) & strOriginal & CHR(13) & CHR(10) & "COMP: " & CHR(9) & strDecompressedCOMPAPI & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & strDecompressedZLIB

end if

if false then

	strOriginal = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

	set xmlDoc1 = CreateObject("Msxml2.DOMDocument")
	set xmlDoc2 = CreateObject("Msxml2.DOMDocument")

	xmlDoc1.appendChild(xmlDoc1.createNode(1, "ROOT", ""))
	set root1 = xmlDoc1.documentElement
	root1.nodeTypedValue = strOriginal

	xmlDoc2.appendChild(xmlDoc2.createNode(1, "ROOT", ""))
	set root2 = xmlDoc2.documentElement
	
	obj.CompressionMethod = "COMPAPI"
	obj.SafeArrayDecompressToXMLNODE obj.XMLNODECompressToSafeArray(root1), root2
	strDecompressedCOMPAPI = root2.nodeTypedValue

	obj.CompressionMethod = "ZLIB"
	obj.SafeArrayDecompressToXMLNODE obj.XMLNODECompressToSafeArray(root1), root2
	strDecompressedZLIB = root2.nodeTypedValue

	msgbox "Old: " & CHR(9) & strOriginal & CHR(13) & CHR(10) & "COMP: " & CHR(9) & strDecompressedCOMPAPI & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & strDecompressedZLIB
	
end if

if false then
	
	strOriginal = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

	Dim Array1()
	Dim Array2()
	
	ReDim Array1(Len(strOriginal))
	
	for index = 1 to Len(strOriginal) 
		Array1(index) = ASC(Mid(strOriginal, index, 1))
	next
	
	obj.CompressionMethod = "COMPAPI"
	obj.SafeArrayCompressToSafeArray(Array1)
	'Array2 = obj.SafeArrayDecompressToSafeArray(obj.SafeArrayCompressToSafeArray(Array1))
	
	strDecompressedCOMPAPI = ""
	for index = 0 to UBound(Array2) - 1
		strDecompressedCOMPAPI = strDecompressedCOMPAPI & CHR(Array2(index))
	next

	strDecompressedZLIB = ""
	
'	obj.CompressionMethod = "ZLIB"
'	obj.SafeArrayDecompressToXMLNODE obj.SafeArrayCompressToSafeArray(root1), root2
'	strDecompressedZLIB = root2.nodeTypedValue

	msgbox "Old: " & CHR(9) & strOriginal & CHR(13) & CHR(10) & "COMP: " & CHR(9) & strDecompressedCOMPAPI & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & strDecompressedZLIB
	
end if
