import subprocess
import os
for filename in os.listdir(os.getcwd()):
    if filename.endswith('.doc'):
        print filename
        subprocess.call(['soffice', '--headless', '--convert-to', 'docx', filename])
        os.remove(filename)