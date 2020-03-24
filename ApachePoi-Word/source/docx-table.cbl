       >>SOURCE FORMAT FREE
       program-id. docx-table.

       configuration section.
       repository.
           class jFile as "java.io.File"
           class jFileOutputStream as "java.io.FileOutputStream"
           class XWPFDocument as "org.apache.poi.xwpf.usermodel.XWPFDocument"
           class XWPFParagraph as "org.apache.poi.xwpf.usermodel.XWPFParagraph"
           class XWPFRun as "org.apache.poi.xwpf.usermodel.XWPFRun"
           class Borders as "org.apache.poi.xwpf.usermodel.Borders"
           class XWPFTable as "org.apache.poi.xwpf.usermodel.XWPFTable"
           class XWPFTableRow as "org.apache.poi.xwpf.usermodel.XWPFTableRow"
           .
           
       working-storage section.
       77 w-jFile             object reference jFile.
       77 w-jFileOutputStream object reference jFileOutputStream.
       77 w-XWPFDocument      object reference XWPFDocument.
       77 w-XWPFParagraph     object reference XWPFParagraph.
       77 w-XWPFRun           object reference XWPFRun.
       77 w-Borders           object reference Borders.
       77 w-XWPFTable         object reference XWPFTable.
       77 w-XWPFTableRow1     object reference XWPFTableRow.
       77 w-XWPFTableRow2     object reference XWPFTableRow.
       77 w-XWPFTableRow3     object reference XWPFTableRow. 
       77 mytext              pic x any length. 

       procedure division.
       main.
           move "Docx written with isCOBOL using ApachePOI interface"
           to mytext.
       
       try
         *>Blank Document
         set w-XWPFDocument to XWPFDocument:>new()       
         *>Write the Document in file system
         set w-jFileOutputStream to jFileOutputStream:>new(jFile:>new("sample-table.docx"))  
         *>create Paragraph
         set w-XWPFParagraph to w-XWPFDocument:>createParagraph()         
         *>Set bottom border to paragraph     
         w-XWPFParagraph:>setBorderBottom(Borders:>BASIC_BLACK_DASHES)   
         *>Set left border to paragraph   
         w-XWPFParagraph:>setBorderLeft(Borders:>BASIC_BLACK_DASHES)         
         *>Set right border to paragraph
         w-XWPFParagraph:>setBorderRight(Borders:>BASIC_BLACK_DASHES)                
         *>Set top border to paragraph
         w-XWPFParagraph:>setBorderTop(Borders:>BASIC_BLACK_DASHES)
         
         set w-XWPFRun to w-XWPFParagraph:>createRun()         
         w-XWPFRun:>setText(mytext)         
         
         *>create table       
         set w-XWPFTable to w-XWPFDocument:>createTable()  

         *>create first row
         set w-XWPFTableRow1 to w-XWPFTable:>getRow(0)
         w-XWPFTableRow1:>getCell(0):>setText("col one, row one")                   
         w-XWPFTableRow1:>addNewTableCell():>setText("col two, row one")         
         w-XWPFTableRow1:>addNewTableCell():>setText("col three, row one")
         *>create second row
         set w-XWPFTableRow2 to w-XWPFTable:>createRow()
         w-XWPFTableRow2:>getCell(0):>setText("col one, row two")
         w-XWPFTableRow2:>getCell(1):>setText("col two, row two")
         w-XWPFTableRow2:>getCell(2):>setText("col three, row two")
         *>create third row
         set w-XWPFTableRow3 to w-XWPFTable:>createRow()
         w-XWPFTableRow3:>getCell(0):>setText("col one, row three")
         w-XWPFTableRow3:>getCell(1):>setText("col two, row three")
         w-XWPFTableRow3:>getCell(2):>setText("col three, row three")         
         
         w-XWPFDocument:>write(w-jFileOutputStream)
         
         w-jFileOutputStream:>close()
         
         display message "docx created"
         
         display message "salca"
            
       catch exception
         display message exception-object:>getMessage()
         
       end-try.
       goback.