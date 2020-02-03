from cx_Freeze import setup, Executable 
import sys
import os

os.environ['TCL_LIBRARY'] = r'tcl8.6'
os.environ['TK_LIBRARY'] = r'tk8.6'


build_exe_options = {
                     "include_files":["_strptime.py","tcl86t.dll", "tk86t.dll"],                     
                     'packages': ['pandas', 'numpy','tkinter','os'],
                     'include_msvcr': True,
                     }
  
base = None
if sys.platform == "win32":
    base = "Win32GUI"



setup(name = "GeeksforGeeks" , 
      version = "0.1" , 
      description = "" , 
      executables = [Executable("cli.py",base=base)]) 