
GroupWareps.dll: dlldata.obj GroupWare_p.obj GroupWare_i.obj
	link /dll /out:GroupWareps.dll /def:GroupWareps.def /entry:DllMain dlldata.obj GroupWare_p.obj GroupWare_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del GroupWareps.dll
	@del GroupWareps.lib
	@del GroupWareps.exp
	@del dlldata.obj
	@del GroupWare_p.obj
	@del GroupWare_i.obj
