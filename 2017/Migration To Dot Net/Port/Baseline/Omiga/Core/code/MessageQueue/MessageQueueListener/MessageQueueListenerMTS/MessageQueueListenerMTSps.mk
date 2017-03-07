
MessageQueueListenerMTSps.dll: dlldata.obj MessageQueueListenerMTS_p.obj MessageQueueListenerMTS_i.obj
	link /dll /out:MessageQueueListenerMTSps.dll /def:MessageQueueListenerMTSps.def /entry:DllMain dlldata.obj MessageQueueListenerMTS_p.obj MessageQueueListenerMTS_i.obj \
		mtxih.lib mtx.lib mtxguid.lib \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \
		ole32.lib advapi32.lib 

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		/MD \
		$<

clean:
	@del MessageQueueListenerMTSps.dll
	@del MessageQueueListenerMTSps.lib
	@del MessageQueueListenerMTSps.exp
	@del dlldata.obj
	@del MessageQueueListenerMTS_p.obj
	@del MessageQueueListenerMTS_i.obj
