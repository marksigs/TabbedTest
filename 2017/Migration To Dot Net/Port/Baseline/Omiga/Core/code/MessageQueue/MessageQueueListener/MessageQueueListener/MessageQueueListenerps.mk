
MessageQueueListenerps.dll: dlldata.obj MessageQueueListener_p.obj MessageQueueListener_i.obj
	link /dll /out:MessageQueueListenerps.dll /def:MessageQueueListenerps.def /entry:DllMain dlldata.obj MessageQueueListener_p.obj MessageQueueListener_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del MessageQueueListenerps.dll
	@del MessageQueueListenerps.lib
	@del MessageQueueListenerps.exp
	@del dlldata.obj
	@del MessageQueueListener_p.obj
	@del MessageQueueListener_i.obj
