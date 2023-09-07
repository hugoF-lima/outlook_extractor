# Considering how the temp/gen_py folder when created tends to cause errors inside win32com.client
# such as ' has no attribute 'CLSIDToPackageMap' ', the below solution seeks to counter that
# Each time the temporary folder gets deleted when the routine is used.


def custom_dispatch(app_name: str):
    try:
        from win32com import client

        # app = client.gencache.EnsureDispatch(app_name)
        # For outlook, bellow:
        app = client.Dispatch(app_name).GetNamespace("MAPI")
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil

        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r"win32com\.gen_py\..+", module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
        from win32com import client

        app = client.Dispatch(app_name).GetNamespace("MAPI")
        # app = client.gencache.EnsureDispatch(app_name)
    return app
