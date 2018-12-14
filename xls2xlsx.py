import glob
import subprocess

for doc in glob.iglob("*.xls"):
    subprocess.call(['soffice', '--headless', '--convert-to', 'xlsx', doc])
