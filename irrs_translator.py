sample_txt = "Exac^^te*ch!!!,Humeral,Li#ner$s,Pr%ojec()t"
print(sample_txt)

def processIRRS(txt):
    
    # These would be the collection of incorrectly displayed symbols
    # In the future the real symbols will come from a table
    wrong_GDT_symbols = "!#$%^&*()"
    
    for symbol in wrong_GDT_symbols:
        # This is a proof of concept where I am replacing each
        # wrong symbol with nothing, basically removing them
        # in the future they will be replaced by the correct symbol
        txt = txt.replace(symbol, '')
    # Replacing all the commas with empty spaces for demonstration
    txt = txt.replace(',', ' ')
    print(txt)

processIRRS(sample_txt)

# TODO:
    # [] add a function to read an excel table with the codes
    #   translation
    # [] modify the processIRRS function to replace the symbols based
    #   on the translation table
    # [] add a function to read each IRRS to be translated
    # [] add a function to export each translated IRRS