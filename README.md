# nmap-converter
Python script for converting nmap reports into XLS
# Requirements
```bash 
sudo pip install python-libnmap
sudo pip install XlsxWriter
```
or 
```bash 
sudo pip install -r requirements.txt
```
# Usage
```bash
usage: nmap-converter.py [-h] [-o XLS] XML [XML ...]

positional arguments:
  XML                   path to nmap xml report

optional arguments:
  -h, --help            show this help message and exit
  -o XLS, --output XLS  path to xlsx output
```
