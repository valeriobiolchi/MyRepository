       >>SOURCE FORMAT FREE
       program-id. doc-open-print.

       configuration section.
       repository.
           class jFile as "java.io.File"
           class jFileOutputStream as "java.io.FileOutputStream"
           class XWPFDocument as "org.apache.poi.xwpf.usermodel.XWPFDocument"
           class OPCPackage as "org.apache.poi.openxml4j.opc.OPCPackage"
           class Desktop as "java.awt.Desktop"
           .
           
       working-storage section.
       77 w-jFile             object reference jFile.
       77 w-Desktop           object reference Desktop.
       77 w-jFileOutputStream object reference jFileOutputStream.
       77 w-XWPFDocument      object reference XWPFDocument.
       77 mytext              pic x any length. 

       procedure division.
       main.
 
       try
         set w-JFile to jFile:>new("C:\temp\OpenOffice.odt")
         Desktop:>getDesktop:>print(w-JFile)
       catch exception
         display message exception-object:>getMessage()
         
       end-try.
       goback.