* Quick Reference *

How to use the ta-lib Excel Add-Ins?
====================================
You must first build the "Add-Ins" (See next section).

In Excel, select "Add-Ins", browse and select
either "ta-lib.xll" or "ta-lib-reverse.xll".

All the TA functions will become visible in 
the "Function Wizard".


How to rebuild the Excel add-in?
================================
You need Visual Studio 2008 to re-build.

Step #1 : Build ta_libc_cdr.lib
-----------------------------------
(1) Open workspace: ta-lib\c\ide\vs2005\lib_proj\ta_lib.sln
(2) Select "ta_libc" as the active project (will appear in bold)
(3) Select the configuration "Win32 CDR Multithread DLL Release".
(4) Build "ta_libc". The file "ta_libc_cdr.lib" will be created in ta-lib\c\lib
(5) Close this workspace.

Step #2 : Build ta-lib.xll or ta-lib-reverse.xll
-------------------------------------------------
(1) Open workspace: ta-lib\excel\src\ta-lib.sln
(2) Select "xlw_for_talib" as the active project.
(3) Select either the configuration "Release" or "Release Reverse"
(4) Build "xlw_for_talib". The file "ta-lib.xll" or "ta-lib-reverse.xll" 
    will be created in the directory ta-lib\excel




What is the different between "ta-lib.xll" and "ta-lib-reverse.xll"?
====================================================================
They are the same except for the order used to process the input.

"ta-lib.xll" performs the calculation from top to down or left to right.
"ta-lib-reverse.xll" performs the calculation from down to top or right to left.

License
=======
You can freely re-distribute ta-lib.xll and ta-lib-reverse.xll, even with 
closed source commercial products. You must respect the LICENSE_EXCEL.TXT 
included in this package.

It is illegal to redistribute the .xll files without a copy
of LICENSE_EXCEL.TXT. Read the License for details.


Documentation/Comments/Help
===========================
  http://ta-lib.org
  http://tadoc.org
