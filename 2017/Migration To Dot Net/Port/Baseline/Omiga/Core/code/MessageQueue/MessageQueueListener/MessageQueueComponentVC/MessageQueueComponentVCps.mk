
MessageQueueComponentVCps.dll: dlldata.obj MessageQueueComponentVC_p.obj MessageQueueComponentVC_i.obj
	link /dll /out:MessageQueueComponentVCps.dll /def:MessageQueueComponentVCps.def /entry:DllMain dlldata.obj MessageQueueComponentVC_p.obj MessageQueueComponentVC_i.obj \
		mtxih.lib mtx.lib mtxguid.lib \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \
		ole32.lib advapi32.lib 

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		/MD \
		$<

clean:
	@del MessageQueueComponentVCps.dll
	@del MessageQueueComponentVCps.lib
	@del MessageQueueComponentVCps.exp
	@del dlldata.obj
	@del MessageQueueComponentVC_p.obj
	@del MessageQueueComponentVC_i.obj
