
omMutexps.dll: dlldata.obj omMutex_p.obj omMutex_i.obj
	link /dll /out:omMutexps.dll /def:omMutexps.def /entry:DllMain dlldata.obj omMutex_p.obj omMutex_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del omMutexps.dll
	@del omMutexps.lib
	@del omMutexps.exp
	@del dlldata.obj
	@del omMutex_p.obj
	@del omMutex_i.obj
