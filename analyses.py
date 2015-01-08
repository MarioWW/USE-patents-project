__author__ = 'williewonka'

from textmining import TextMining

analyzer = TextMining('patentdata.xlsx')

analyzer.Parse_Categories('categorien.xlsx')
analyzer.Parse_Word_Counting('words.xlsx')
analyzer.Parse_Categories('categorien.xlsx')
analyzer.Parse_Company_Counting('Companycounting.xlsx')
