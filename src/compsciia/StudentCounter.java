package compsciia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StudentCounter {
    
/*
    TermsLocator terms = new TermsLocator();
    NeededDocu excelFile = new NeededDocu();
    XSSFWorkbook grades = excelFile.getWorkbook();
    FormulaEvaluator evaluator = grades.getCreationHelper().createFormulaEvaluator();
*/
    TermsLocator terms;
    NeededDocu excelFile;
    XSSFWorkbook grades;
    FormulaEvaluator evaluator;
    
    public StudentCounter(){
     System.out.println("RUNNING STUDENTCOUNTER");
     excelFile  = new NeededDocu(); System.out.println("NeededDocu object created");
     evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
     terms = new TermsLocator(); 
     
    }
    
    public StudentCounter(String searchTrimester){
     System.out.println("RUNNING STUDENTCOUNTER");
     excelFile  = new NeededDocu(); System.out.println("NeededDocu object created");
     evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
     terms = new TermsLocator(searchTrimester);
     terms.findTermLocation(searchTrimester);
     
    }
    
    private int studentNumber = 0;
    

    public int countStudents() {
        int count =2;
        boolean done = false;
        int row = terms.rowNum; // the row that the cell containing trimester 
        int col = terms.colNum; // the column that the cell containing trimester
        XSSFSheet sheet = excelFile.getSheet();
            while (done==false) {
                
                Row studentRow = sheet.getRow(row+count);
                    if (studentRow==null){ studentRow = sheet.createRow(row+count); } // null test
                Cell studentCol = studentRow.getCell(col);
                    if (studentCol==null){ studentCol = studentRow.createCell(col);} // null test
                    
                if (((evaluator.evaluateInCell(studentCol).getCellType())!=Cell.CELL_TYPE_BLANK ) && !(studentCol.getStringCellValue().toLowerCase().contains("average"))){
                    studentNumber++;
                    count++;
                } else {    done = true; count =2; }
            } 
            return studentNumber;
        }
        
        public int countStudents(int row, int col) {
        int count =2;
        boolean done = false;
        XSSFSheet sheet = excelFile.getSheet();
            while (done==false) {
                Row studentRow = sheet.getRow(row+count);          if (studentRow==null){ studentRow = sheet.createRow(row+count); }
                Cell studentCol = studentRow.getCell(col);                if (studentCol==null){ studentCol = studentRow.createCell(col);}
                if (((evaluator.evaluateInCell(studentCol).getCellType())!=Cell.CELL_TYPE_BLANK ) && !(studentCol.getStringCellValue().toLowerCase().contains("average"))){
                    studentNumber++;
                    count++;
                } else {    done = true; count =2; }
                                                } 
            return studentNumber;
        }
    }


