import zipfile
import xmltodict
from nested_lookup import nested_lookup
import re
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-x", "--docx", help="target is a modern docx file")
parser.add_argument("-d", "--doc", help="Target is a a regular-old doc file")
args = parser.parse_args()

if args.docx is True:
    #Parsing as a docx document:

    #Defining a regex for DDE/DDEAUTO
    regex = re.compile("DDE.*")
    print "Very simple DDE detector for word documents\r\n Written by Amit Serper, @0xAmit\n\n"

    #First we open the docx file as a zip
    document = zipfile.ZipFile(args.docx, 'r')
    #Parsing the XML
    xmldata = document.read('word/document.xml')
    d = xmltodict.parse(xmldata)
    #Looking up for the DDE object in this position, flattening the xml because we're lazy.
    DDE = nested_lookup('w:instrText', d)
    if DDE:
        print "Malicious DDE objects found: \n{0}".format(regex.findall(str(DDE)))
    else:
        print "No DDE objects were found"

else:
    #Parsing as an old DOC file
    with open(args.doc, 'rb') as doc:
        docstr = doc.read()
    pos = docstr.find('DDE') # We're looking for the string 'DDE', this will return its position.
    pos = pos-1 # Deducting 1 so we're now 1 char before the string starts
    doc_regex = re.compile('^[^"]+') # Regex is looking for the final char to be '"' since that's our delimiter.
    res = doc_regex.findall(docstr[pos:]) # Matching from regex to '"'
    if "DDE" in str(res):
        print "Malicious DDE objects found:\n{0}".format(res)
    else:
        print "No DDE objects were found"






