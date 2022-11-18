import pip

packages = ['openpyxl', 'requests', 'pywin32', 'pandas']

try:
    for pkg in packages:
        pip.main(['install', pkg])
    print("Packages Installed Successfully!")
except Exception as ex:
    print(ex)

