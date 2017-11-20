# DDEtect
Written by Amit Serper, [@0xAmit](http://twitter.com/0xAmit)
DDEtector is a simple DDE object detector written in python

  - Currently supports only word DOCX and legacy DOC files
  - Prints the contents of the DDE payloads (Note: In some cases DDEtect won't print the entire DDE payload. I'm working on writing a better matching algorithm)
  - More features coming soon...

### Running DDEtector

Execute the python file and supply a path to a docx file as an argument. Use the -d argument for a regular doc file or -x for a docx file:
![alt text](https://github.com/aserper/DDEtect/blob/master/ddetect.jpg?raw=true)
DDEtector requires the following python modules:
- zipfile
- xmltodict
- nested_lookup
- re
- argparse



### Todos
- Format autodetection
- Support other office formats (ie. excel)
