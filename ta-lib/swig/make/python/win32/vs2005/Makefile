# MSVC makefile for building a TaLib Python module (.py and .dll)
# CMR (commandline multitheraded release only)
# Handmade (experimental)
#
# Tested with SWIG v1.3.33, Python v2.5.1, MSVS2005

PYTHONDIR  = c:\Python25 # !!! maybe you must change it !!!
TOPDIR     = ..\..\..\..
TEMPDIR    = $(TOPDIR)\temp\python\wrap
TARGETDIR  = $(TOPDIR)\lib\python
INSTDIR    = $(PYTHONDIR)\Lib\site-packages
#INTERFACE  = $(TOPDIR)\src\interface\ta_libc.python2.swg # not standard indterface: use exceptions, not return value
INTERFACE  = $(TOPDIR)\src\interface\ta_libc.swg


CC         = cl
LINK       = link
INCPATH    = -I$(TOPDIR)\..\c\include -I"$(PYTHONDIR)\include"
LIBPATH	   = -LIBPATH:"$(PYTHONDIR)\libs" -LIBPATH:"$(TOPDIR)\..\c\lib"	
CFLAGS     = -nologo -DWIN32 -O2
LFLAGS	   = /NOLOGO

default: install

##############################################################################
all: $(TARGETDIR)\_TaLib.dll $(TARGETDIR)\TaLib.py

$(TEMPDIR)\TaLib.py $(TEMPDIR)\ta_libc_wrap.c: \
		$(TOPDIR)\src\interface\ta_libc.swg \
		$(TOPDIR)\src\interface\ta_libc.python.swg \
		$(TOPDIR)\src\interface\ta_func.swg
	swig -python -module TaLib -modern -Fstandard \
		-I$(TOPDIR)\src\interface -I$(TOPDIR)\..\c\include \
		-outdir $(TEMPDIR) -o $(TEMPDIR)\ta_libc_wrap.c \
		$(INTERFACE)

$(TEMPDIR)\ta_libc_wrap.obj: $(TEMPDIR)\ta_libc_wrap.c
	$(CC) -c $(CFLAGS) $(INCPATH) -Fo$@ $(TEMPDIR)\ta_libc_wrap.c

$(TARGETDIR)\_TaLib.dll: $(TEMPDIR)\ta_libc_wrap.obj
	$(LINK) $(LFLAGS) $(LIBPATH) /DLL /OUT:$@ \
		$(TEMPDIR)\ta_libc_wrap.obj ta_func_cmr.lib ta_abstract_cmr.lib ta_common_cmr.lib 

$(TARGETDIR)\TaLib.py: $(TEMPDIR)\TaLib.py $(TOPDIR)\src\interface\python.py 
	copy /b $(TEMPDIR)\TaLib.py + $(TOPDIR)\src\interface\python.py $@

##############################################################################
clean:
	del /q $(TEMPDIR)\*.c
	del /q $(TEMPDIR)\*.py
	del /q $(TEMPDIR)\*.obj
	del /q $(TARGETDIR)\*.lib
	del /q $(TARGETDIR)\*.exp
	del /q $(TARGETDIR)\_TaLib.dll
	del /q $(TARGETDIR)\TaLib.py

##############################################################################
install: "$(INSTDIR)\_TaLib.pyd" "$(INSTDIR)\TaLib.py"

"$(INSTDIR)\_TaLib.pyd": "$(TARGETDIR)\_TaLib.dll"
	copy "$(TARGETDIR)\_TaLib.dll" $@

"$(INSTDIR)\TaLib.py": "$(TARGETDIR)\TaLib.py"
	copy "$(TARGETDIR)\TaLib.py" $@
	cd $(INSTDIR)
	python -c "import TaLib"
	python -O -c "import TaLib"
