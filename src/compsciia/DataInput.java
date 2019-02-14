/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package compsciia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class DataInput {
    NeededDocu excelFile;
    FormulaEvaluator evaluator;
    String value;
    String toReturn;
    TermsLocator terms;
    private int titleRowNum; 
    
    public String getData(int row, int col) {
        toReturn = evaluatingData(row,col);
        System.out.println("returned valye in DataInput                                                       :      "+row+"                "+col+"                    "+toReturn );
        return toReturn;
    }
    
    public DataInput() {
        System.out.println("RUNNING DATAINPUT");
        excelFile = new NeededDocu();
        evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
        titleRowNum = (terms.rowNum-1);
    }
    
    private String evaluatingData(int row, int col){
          Row studentRow = excelFile.sheet.getRow(row);
            if (studentRow==null){ studentRow = excelFile.sheet.createRow(row); }   // null test
          Cell studentCol = studentRow.getCell(col);
            if (studentCol==null){ studentCol = studentRow.createCell(col);}        // null test
            
          if (evaluator.evaluateInCell(studentCol).getCellType()==Cell.CELL_TYPE_STRING){
                 value = studentCol.getStringCellValue();
          } else if(evaluator.evaluateInCell(studentCol).getCellType()==Cell.CELL_TYPE_NUMERIC){
                 double tempSave;
                  if (studentCol.getCellStyle().getDataFormatString().contains("%")){
                    tempSave = studentCol.getNumericCellValue()*100;
              } else {
                    tempSave = studentCol.getNumericCellValue(); }
                    value = Integer.toString((int)Math.round(tempSave));
                    } else {
                            value = " null ";
                             }
          return value;                              

                                    }
                                }
                        
/*
         Row titleRow = excelFile.sheet.getRow(titleRowNum);                              if (titleRow==null){ titleRow = excelFile.sheet.createRow(titleRowNum); }
          Cell title = titleRow.getCell(col);                                               if (title==null){ title = titleRow.createCell(col);}
          
          while (evaluator.evaluateInCell(title).getCellType()!=Cell.CELL_TYPE_BLANK) {
*/
