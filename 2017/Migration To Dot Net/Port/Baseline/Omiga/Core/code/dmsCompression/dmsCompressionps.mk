
dmsCompressionps.dll: dlldata.obj dmsCompression_p.obj dmsCompression_i.obj
	link /dll /out:dmsCompressionps.dll /def:dmsCompressionps.def /entry:DllMain dlldata.obj dmsCompression_p.obj dmsCompression_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del dmsCompressionps.dll
	@del dmsCompressionps.lib
	@del dmsCompressionps.exp
	@del dlldata.obj
	@del dmsCompression_p.obj
	@del dmsCompression_i.obj
