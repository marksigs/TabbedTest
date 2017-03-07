'set obj = CreateObject("DmsCompression.dmsCompression1")
set obj = CreateObject("DmsCompression.PkZip")

if false then

	bLoop = true
	nCount = 0
	tmrZLIBDiff = 0
	while bLoop and nCount < 100
		tmrZLIBStart = Timer
		'obj.CompressionMethod = "ZLIB"
		obj.ZipFile "C:\DOCS\TADZLIB.ZIP", "C:\DOCS\TAD.DOC"
		tmrZLIBEnd = Timer
		tmrZLIBDiff = tmrZLIBDiff + (tmrZLIBEnd - tmrZLIBStart)

		nCount = nCount + 1
	wend

	msgbox nCount & " iterations" & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & tmrZLIBDiff & " seconds"

end if

if false then

	bLoop = true
	nCount = 0
	tmrZLIBDiff = 0
	while bLoop and nCount < 1
		tmrZLIBStart = Timer
		'obj.CompressionMethod = "ZLIB"
		'obj.UnzipFile "C:\DOCS\TADZLIB.ZIP", "C:\DOCS\TAD.DOC"
		obj.UnzipFile "C:\DOCS\TADZLIB.ZIP"
		tmrZLIBEnd = Timer
		tmrZLIBDiff = tmrZLIBDiff + (tmrZLIBEnd - tmrZLIBStart)

		nCount = nCount + 1
	wend

	msgbox nCount & " iterations" & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & tmrZLIBDiff & " seconds"

end if

if true then

	bLoop = true
	nCount = 0
	tmrZLIBDiff = 0
	while bLoop and nCount < 100
		tmrZLIBStart = Timer
		obj.ZipFile "C:\DOCS\TADZLIB.ZIP", "C:\DOCS\TAD.DOC"
		obj.UnzipFile "C:\DOCS\TADZLIB.ZIP"
		tmrZLIBEnd = Timer
		tmrZLIBDiff = tmrZLIBDiff + (tmrZLIBEnd - tmrZLIBStart)

		nCount = nCount + 1
	wend

	msgbox nCount & " iterations" & CHR(13) & CHR(10) & "ZLIB: " & CHR(9) & tmrZLIBDiff & " seconds"

end if
