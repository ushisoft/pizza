import glob
import subprocess

for doc in glob.iglob("*.doc"):
    subprocess.call(['soffice', '--headless', '--convert-to', 'docx', doc])
