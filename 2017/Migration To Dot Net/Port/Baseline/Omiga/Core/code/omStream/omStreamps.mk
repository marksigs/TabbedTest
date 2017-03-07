
omStreamps.dll: dlldata.obj omStream_p.obj omStream_i.obj
	link /dll /out:omStreamps.dll /def:omStreamps.def /entry:DllMain dlldata.obj omStream_p.obj omStream_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del omStreamps.dll
	@del omStreamps.lib
	@del omStreamps.exp
	@del dlldata.obj
	@del omStream_p.obj
	@del omStream_i.obj
