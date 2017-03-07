
ODIConverterps.dll: dlldata.obj ODIConverter_p.obj ODIConverter_i.obj
	link /dll /out:ODIConverterps.dll /def:ODIConverterps.def /entry:DllMain dlldata.obj ODIConverter_p.obj ODIConverter_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del ODIConverterps.dll
	@del ODIConverterps.lib
	@del ODIConverterps.exp
	@del dlldata.obj
	@del ODIConverter_p.obj
	@del ODIConverter_i.obj
