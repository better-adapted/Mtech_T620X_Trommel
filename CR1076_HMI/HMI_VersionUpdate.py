# encoding:utf-8
from __future__ import print_function
import datetime

# see
# https://content.helpme-codesys.com/en/CODESYS%20Scripting/_cds_access_cds_func_in_python_scripts.html

proj = projects.primary

info = proj.get_project_info()

# Set some values
info.company = "MTech"
info.title = "MAIN_HMI"
info.version = (3,0,2,0)
info.default_namespace = ""
info.author = "AW"
info.description = "T620X/Radial"

# now we set a custom / vendor specific value.
info.values["CompileDateTime"] = datetime.datetime.now().strftime("%H:%M:%S,%d-%m-%Y")

# Enable generation of Accessor functions, so the IEC
# application can display the version in an info screen.
info.change_accessor_generation(True)

# And set the project to released
res = system.ui.prompt("Mark For Release?", PromptChoice.YesNo, PromptResult.Yes)
print("The user selected '%s'" % res)

if res == 'Yes':
    info.released = True
else:
    info.released = False