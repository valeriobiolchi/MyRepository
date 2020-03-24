       >>SOURCE FORMAT FREE 
       program-id. excel-sample.

       configuration section.
       repository.
           class j-cell as "org.apache.poi.ss.usermodel.Cell"
           class j-cellstyle as "org.apache.poi.ss.usermodel.CellStyle"
           class j-creationhelper as "org.apache.poi.ss.usermodel.CreationHelper"
           class j-row as "org.apache.poi.ss.usermodel.Row"
           class j-sheet as "org.apache.poi.ss.usermodel.Sheet"
           class j-workbook as "org.apache.poi.ss.usermodel.Workbook"
           class j-xssfworkbook as "org.apache.poi.xssf.usermodel.XSSFWorkbook"
           class j-hssfworkbook as "org.apache.poi.hssf.usermodel.HSSFWorkbook"                                    
           class j-builtinformats as "org.apache.poi.ss.usermodel.BuiltinFormats"
           class j-xssffont as "org.apache.poi.xssf.usermodel.XSSFFont"
           class j-hssffont as "org.apache.poi.hssf.usermodel.HSSFFont"
           class j-color as "org.apache.poi.ss.usermodel.IndexedColors"
           class j-font2 as "org.apache.poi.hssf.usermodel.HSSFFont"
           class j-calendar as "java.util.Calendar"
           class j-gregorian as "java.util.GregorianCalendar"
           class j-output as "java.io.FileOutputStream"
           class j-input as "java.io.FileInputStream"
           class j-file as "java.io.File"                     
           class j-hssfdateutil as "org.apache.poi.hssf.usermodel.HSSFDateUtil"
           class j-simpledateformat as "java.text.SimpleDateFormat"                                    
           .

       working-storage section.
       77 w-cell object reference j-cell.
       77 w-cellstyle1 object reference j-cellstyle.
       77 w-cellstyle2 object reference j-cellstyle.
       77 w-cellstyle3 object reference j-cellstyle.
       77 w-cellstyle4 object reference j-cellstyle.
       77 w-creationhelper object reference j-creationhelper.
       77 w-row object reference j-row.
       77 w-sheet object reference j-sheet.
       77 w-workbook object reference j-workbook.
       77 w-xssfworkbook object reference j-xssfworkbook.
       77 w-builtinformats object reference j-builtinformats.
       77 w-xssffont object reference j-xssffont.
       77 w-hssffont object reference j-hssffont.
       77 w-color object reference j-color.
       77 w-calendar object reference j-calendar.
       77 w-calendar2 object reference j-calendar.
       77 w-output object reference j-output. 
       77 w-file object reference j-file.
       77 w-simpledateformat object reference j-simpledateformat.
       77 w-hssfworkbook object reference j-hssfworkbook.
       01 type-xls pic 9.
          88 xlsx value 1.
          88 xls value 2.
       77 idx int.
       77 counter int.

       77 w-num pic 9(6)v999.
       
       77 num-sheets int.
       77 num-rows int.
       77 row int.
       77 num-cells int.
       77 cell-content pic x any length.

       procedure division.
       main.
       
           perform create-xlsx.
           perform write-excel.
           display "execel-file.xlsx created".
           display omitted.
           accept omitted.
           
           perform create-xls.
           perform write-excel.
           display "execel-file.xls created".
           display omitted.
                                 
           perform set-xlsx.
           perform read-excel.      
           display omitted.     
           
           perform set-xls.
           perform read-excel.
           
           stop run.
           
       create-xlsx.    
           set w-xssfworkbook to j-xssfworkbook:>new().
           set w-workbook to w-xssfworkbook.
           set xlsx to true.

       create-xls.    
           set w-hssfworkbook to j-hssfworkbook:>new().
           set w-workbook to w-hssfworkbook.
           set xls to true.
           
       write-excel.    
           try             
             set w-sheet to w-workbook:>createSheet("sheet1")
             set w-cellstyle1 to w-workbook:>createCellStyle()
             set w-cellstyle2 to w-workbook:>createCellStyle()
             set w-cellstyle3 to w-workbook:>createCellStyle()
             w-cellstyle1:>setDataFormat(j-builtinformats:>getBuiltinFormat("m/d/yy") as short)
             w-cellstyle2:>setDataFormat(j-builtinformats:>getBuiltinFormat("d-mmm-yy") as short)
             set w-calendar to j-gregorian:>new()             
             w-calendar:>set(1990,j-calendar:>APRIL,2)
             set w-row to w-sheet:>createRow(0)
             set w-cell to w-row:>createCell(0)
             w-cell:>setCellValue(w-calendar)
             w-cell:>setCellStyle(w-cellstyle1)
      *> create the font and set the color on it applied on a particular cell
             evaluate type-xls
             when 1    
                set w-xssffont to w-xssfworkbook:>createFont()
                w-xssffont:>setFontHeightInPoints(18)
                w-xssffont:>setFontName("Comic Sans Serif")
                w-xssffont:>setItalic(true)
                w-xssffont:>setColor(j-color:>GREEN:>getIndex())
                w-cellstyle3:>setFont (w-xssffont)
             when 2   
                set w-hssffont to w-hssfworkbook:>createFont()
                w-hssffont:>setFontHeightInPoints(18)
                w-hssffont:>setFontName("Comic Sans Serif")
                w-hssffont:>setItalic(true)
                w-hssffont:>setColor(j-color:>GREEN:>getIndex())
                w-cellstyle3:>setFont (w-hssffont)                
             end-evaluate
             set w-cell to w-row:>createCell(1)
             w-cell:>setCellValue("string1")
             w-cell:>setCellStyle(w-cellstyle3)
             
             set w-cell to w-row:>createCell(2)
             move 1234.567 to w-num;;
             w-cell:>setCellValue(w-num as double)
                                       
             set w-row to w-sheet:>createRow(1)
             set w-cell to w-row:>createCell(0)
             w-cell:>setCellValue("SALCA")
             set w-cell to w-row:>createCell(1)

      *> highlight all the cell of the second row                    
             set w-cellstyle4 to w-workbook:>createCellStyle()
             evaluate type-xls
             when 1                 
                set w-xssffont to null                          
                set w-xssffont to w-xssfworkbook:>createFont()                                                                                      
                w-xssffont:>setBold(true)
                w-cellstyle4:>setFont(w-xssffont)
                w-cellstyle2:>setFont(w-xssffont)
             when 2
                set w-hssffont to null    
                set w-hssffont to w-hssfworkbook:>createFont()                      
                w-hssffont:>setBold(true)
                w-cellstyle4:>setFont(w-hssffont)  
                w-cellstyle2:>setFont(w-hssffont)                        
             end-evaluate   
             set idx to w-row:>getLastCellNum()
             perform varying counter from 0 by 1 until counter = idx
                w-row:>getCell(counter):>setCellStyle(w-cellstyle4)
             end-perform   
             w-cell:>setCellValue(w-calendar)             
             w-cell:>setCellStyle(w-cellstyle2)

             set w-cell to w-row:>createCell(2)
             move 9876.543 to w-num;;
             w-cell:>setCellValue(w-num as double)

                                                                 
             evaluate type-xls
             when 1
                set w-output to j-output:>new(j-file:>new("excel-file.xlsx"))
             when 2   
                set w-output to j-output:>new(j-file:>new("excel-file.xls"))
             end-evaluate
                
             w-workbook:>write(w-output)
             w-output:>close()

             

           catch exception
             display message exception-object:>getMessage()
           end-try.
           
       set-xlsx.    
           try
              set w-workbook to j-xssfworkbook:>new(j-file:>new("excel-file.xlsx"))
              display "READING XLSX"
           catch exception
              display message exception-object:>getMessage()           
           end-try.   
           
       set-xls.    
           try
              set w-workbook to j-hssfworkbook:>new(j-input:>new("excel-file.xls"))
              display "READING XLS"
           catch exception
              display message exception-object:>getMessage()           
           end-try.   
                      
       read-excel.  
           try
              |set w-workbook to j-xssfworkbook:>new(j-file:>new("excel-file.xlsx"))
              |set w-workbook to j-hssfworkbook:>new(j-input:>new("excel-file.xls"))
              display "Data dump:"
              set num-sheets to w-workbook:>getNumberOfSheets()
              perform varying idx from 0 by 1 until idx >= num-sheets
                 set w-sheet to w-workbook:>getSheetAt(idx)
                 set num-rows to w-sheet:>getPhysicalNumberOfRows
                 display "Sheet " idx ' "' w-workbook:>getSheetName(idx) '"' 
                        " has " num-rows " row(s)."
                 perform varying row from 0 by 1 until row >= num-rows       
                    set w-row to w-sheet:>getRow(row)
                    if w-row not = null
                       set num-cells to w-row:>getPhysicalNumberOfCells()
                       display omitted
                       display "ROW " w-row:>getRowNum() " has " 
                               num-cells " cell(s)."
                       perform varying counter from 0 by 1 until counter >= num-cells
                          set w-cell to w-row:>getCell(counter)
                          display "CELL col=" w-cell:>getColumnIndex()
                          evaluate w-cell:>getCellType()
                          when w-cell:>CELL_TYPE_FORMULA
                               set cell-content to w-cell:>getCellFormula()                                
                          when w-cell:>CELL_TYPE_NUMERIC                    
                               if j-hssfdateutil:>isCellDateFormatted(w-cell)
                                  set w-simpledateformat to j-simpledateformat:>new("MM/dd/yyyy")
                                  set cell-content to w-simpledateformat:>format(w-cell:>getDateCellValue())                                  
                               else   
                                  set cell-content to w-cell:>getNumericCellValue()                                  
                               end-if   
                          when w-cell:>CELL_TYPE_STRING
                               set cell-content to w-cell:>getStringCellValue()                               
                           end-evaluate
                           display "CELL col=" w-cell:>getColumnIndex() " VALUE=" cell-content
                       end-perform
                    end-if
                 end-perform
              end-perform                                                                                           
           catch exception
              display message exception-object:>getMessage()           
           end-try.
         
           continue.
          
