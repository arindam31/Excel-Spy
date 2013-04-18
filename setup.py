#Full working with icon working too.


from distutils.core import setup
import py2exe


#Find details in py2exe\build_exe.py and __init__.py
setup(
    # The first three parameters are not required, if at least a
    # 'version' is given, then a versioninfo resource is built from
    # them and added to the executables.
    version = "1.0.2",
    description = "An application to search within Excel Files (only .xls support)",
    name = "Xcel Spy",

    # targets to build
    windows = [
        {
            "script":"Main.py",
            "icon_resources":[(0, "desktop.ico")]
        }
              ],
    options = {"py2exe":
                          {
                              "dll_excludes":["MSVCP90.dll"]
                              
                          }
               }

    )
