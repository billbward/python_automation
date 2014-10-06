import win32com.client
import os

#word_doc
root = os.path.join('C:\\Users\\bgtx.bward\\Desktop\\BGTX\\Projects\\Polling Locations\\')
file_name = 'PollingLocationTemplateTest_2_20141006.docx'
file = os.path.join(root,file_name)


#Create an instance of Word.Application
wordApp = win32com.client.Dispatch('Word.Application')

#Show the application
wordApp.Visible = True
wordApp.DisplayAlerts = 0

#Open the doc
wordDoc = wordApp.Documents.Open(file) 

#Set the text of the first paragraph
#wordDoc.Paragraphs(1).Range.Text = "Hello, World! \n"


#replace the test text
def search_replace_all(word_file, find_str, replace_str):
    ''' replace all occurrences of `find_str` w/ `replace_str` in `word_file` '''
    wdFindContinue = 1
    wdReplaceAll = 2

	

    # expression.Execute(FindText, MatchCase, MatchWholeWord,
    #   MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, 
    #   Wrap, Format, ReplaceWith, Replace)
    wordApp.Selection.Find.Execute(find_str, False, False, False, False, False, \
        True, wdFindContinue, False, replace_str, wdReplaceAll)
    #wordApp.ActiveDocument.Close(SaveChanges=True)
    #wordApp.Quit()

#f = 'c:/path/to/my/word.doc'
search_replace_all(wordDoc, 'Precinct Name', 'Precinct 4006')
search_replace_all(wordDoc, 'Polling Location Name', 'Wendy Davis HQ')
search_replace_all(wordDoc, 'Polling Location Street', '219 S Main St')
search_replace_all(wordDoc, 'Polling Location City', 'Fort Worth')
search_replace_all(wordDoc, 'Polling Location Zip', '76104')

save_pdf = 'C:\\Users\\bgtx.bward\\Desktop\\BGTX\\Projects\\Polling Locations\\PollingLocationQuarterSheetTest' + '.pdf'
wdFormatPDF = 17
wordDoc.SaveAs(save_pdf, FileFormat = wdFormatPDF)

search_replace_all(wordDoc, 'Precinct 4006', 'Precinct Name')
search_replace_all(wordDoc, 'Wendy Davis HQ', 'Polling Location Name')
search_replace_all(wordDoc, '219 S Main St', 'Polling Location Street')
search_replace_all(wordDoc, 'Fort Worth', 'Polling Location City')
search_replace_all(wordDoc, '76104', 'Polling Location Zip')

wordDoc.Close()
wordApp.Quit()
