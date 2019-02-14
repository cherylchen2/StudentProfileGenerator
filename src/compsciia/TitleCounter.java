/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package compsciia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author lenovo
 */
public class TitleCounter {
    
    private boolean done1 = false;
    private boolean done2 = false;
    NeededDocu excelFile;
    TermsLocator terms;
    XSSFSheet sheet;
    XSSFWorkbook grades;
    FormulaEvaluator evaluator;
    private int colCheckCount;
    private int count1;
    static int startTitleNum;
    private String testText;
    private String cellValue;
    private String comparisonValue;
    
    public TitleCounter() {
        System.out.println("RUNNING TITLECOUNTER");
        excelFile = new NeededDocu();
        terms = new TermsLocator();
        sheet = excelFile.getSheet();
        evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
        testText = "test";
        comparisonValue = "final";
    }
    
    public TitleCounter(String searchTrimester) {
        System.out.println("RUNNING TITLECOUNTER FOR "+searchTrimester);
        excelFile = new NeededDocu();
        terms = new TermsLocator(searchTrimester);
        terms.findTermLocation(searchTrimester);
        sheet = excelFile.getSheet();
        evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
        testText = "test";
        comparisonValue = "final";
    }
    
    public int CountTitle() {
        
    int row = terms.rowNum;
    int col = terms.colNum;
           while (done1 == false ) {
               Row searchTestRow = excelFile.sheet.getRow(row+1);
                    if (searchTestRow == null) {searchTestRow = excelFile.sheet.createRow(row+1);} // null test
                    
               Cell searchTest = searchTestRow.getCell(col+count1);
                    if (searchTest == null) {searchTest = searchTestRow.createCell(col+count1);}   // null test
                    
               cellValue = searchTest.getStringCellValue();         // cellValue is a String variable declared earlier in the class's constructor
               if (cellValue.toLowerCase().contains(testText)) {  // testText has string value of 'test'
                   startTitleNum = col+count1;
                   done1=true;
               } else {
                   count1++;
               }
               
           }
          while (done2==false) {
                Row colCheckRow =excelFile. sheet.getRow(row+1);
                    if (colCheckRow==null){ colCheckRow = excelFile.sheet.createRow(row+1); } // null test
                Cell colCheckCol = colCheckRow.getCell(startTitleNum+colCheckCount);        // startTitleNum = First column of data to be printed
                    if (colCheckCol==null){ colCheckCol = colCheckRow.createCell(startTitleNum+colCheckCount);} //null test
                                
                String valueInCell = colCheckCol.getStringCellValue();  // 2 String varaibles that will be compared to each other
                                
                if ((valueInCell.toLowerCase().contains(comparisonValue.toLowerCase()))){
                    boolean finalLast = false;
                     while (finalLast = false) {
                        Cell checkNextCol = colCheckRow.getCell(startTitleNum+colCheckCount+1);
                                if (checkNextCol==null){ checkNextCol = colCheckRow.createCell(startTitleNum+colCheckCount+1);}
                       if ((evaluator.evaluateInCell(checkNextCol).getCellType())!=Cell.CELL_TYPE_BLANK){ 
                        if (checkNextCol.getStringCellValue().toLowerCase().contains(comparisonValue.toLowerCase())) {
                            colCheckCount++;
                        } else { finalLast = true; } }
                     }
                    colCheckCount = colCheckCount +2;
                    done2 = true;
                } else {
                    colCheckCount++;} 
                                                } 
    return colCheckCount;
                            
}

        public  int CountTitle(int row, int col) {
            System.out.println( "In CountTitle for paramenters and the 2 parameters are:  "+row+" and "+col+"                      \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\");
           while (done1 == false ) {
               Row searchTestRow = excelFile.sheet.getRow(row+1);               if (searchTestRow == null) {searchTestRow = excelFile.sheet.createRow(row+1);}
               Cell searchTest = searchTestRow.getCell(col+count1);               if (searchTest == null) {searchTest = searchTestRow.createCell(col+count1);}
               cellValue = searchTest.getStringCellValue();
               if (cellValue.toLowerCase().contains(testText)) {
                   startTitleNum = col+count1;
                   done1=true;
               } else {
                   count1++;
               }
               
           }
          while (done2==false) {
                                Row colCheckRow =excelFile. sheet.getRow(row+1);   if (colCheckRow==null){ colCheckRow = excelFile.sheet.createRow(row+1); }
                                Cell colCheckCol = colCheckRow.getCell(startTitleNum+colCheckCount);       if (colCheckCol==null){ colCheckCol = colCheckRow.createCell(startTitleNum+colCheckCount);}
                                
                                String valueInCell = colCheckCol.getStringCellValue();
                                
                                
                if ((valueInCell.toLowerCase().contains(comparisonValue.toLowerCase()))){
                    boolean finalLast = false;
                     while (finalLast = false) {
                        Cell checkNextCol = colCheckRow.getCell(startTitleNum+colCheckCount+1);
                                if (checkNextCol==null){ checkNextCol = colCheckRow.createCell(startTitleNum+colCheckCount+1);}
                      if ((evaluator.evaluateInCell(checkNextCol).getCellType())!=Cell.CELL_TYPE_BLANK){   
                        if (checkNextCol.getStringCellValue().toLowerCase().contains(comparisonValue.toLowerCase())) {
                            colCheckCount++; }
                        else { finalLast = true; } }
                        } 
                     colCheckCount = colCheckCount +2;
                    done2 = true;
                     } else {
                    colCheckCount++;}
                            
        }
                           return colCheckCount; 
    }
}

