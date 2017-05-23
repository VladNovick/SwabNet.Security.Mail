
SwabNetSMailps.dll: dlldata.obj SwabNetSMail_p.obj SwabNetSMail_i.obj
	link /dll /out:SwabNetSMailps.dll /def:SwabNetSMailps.def /entry:DllMain dlldata.obj SwabNetSMail_p.obj SwabNetSMail_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del SwabNetSMailps.dll
	@del SwabNetSMailps.lib
	@del SwabNetSMailps.exp
	@del dlldata.obj
	@del SwabNetSMail_p.obj
	@del SwabNetSMail_i.obj
