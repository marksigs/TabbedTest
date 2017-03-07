
MessageQueueListenerLOGps.dll: dlldata.obj MessageQueueListenerLOG_p.obj MessageQueueListenerLOG_i.obj
	link /dll /out:MessageQueueListenerLOGps.dll /def:MessageQueueListenerLOGps.def /entry:DllMain dlldata.obj MessageQueueListenerLOG_p.obj MessageQueueListenerLOG_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del MessageQueueListenerLOGps.dll
	@del MessageQueueListenerLOGps.lib
	@del MessageQueueListenerLOGps.exp
	@del dlldata.obj
	@del MessageQueueListenerLOG_p.obj
	@del MessageQueueListenerLOG_i.obj
