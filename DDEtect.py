import zipfile
from sys import argv
import xmltodict
from nested_lookup import nested_lookup
import re

try:
    #Defining a regex for DDE/DDEAUTO
    regex = re.compile("DDE.*")
    print "Very simple DDE detector for word documents\r\n Written by Amit Serper, @0xAmit\n\n"

    #First we open the docx file as a zip
    document = zipfile.ZipFile(argv[1], 'r')
    #Parsing the XML
    #Flattening the dict because we're lazy
    xmldata = document.read('word/document.xml')
    d = xmltodict.parse(xmldata)
    #Looking up for the DDE object in this position
    DDE = nested_lookup('w:instrText', d)

    if DDE:
        print "Malicious DDE objects found: \n{0}".format(regex.findall(str(DDE)))
    else:
        print "No DDE objects were found"
except IndexError:
    print "Error: No path to docx given! Bailing"