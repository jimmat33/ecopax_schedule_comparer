from distutils.core import setup
import py2exe

setup(windows=[{'script' : 'ecopax_schedule_comparison.py', "icon_resources": [(1, "gui_icon.ico")], "dest_base" : "Ecopax Schedule Comparison Tool"}])
