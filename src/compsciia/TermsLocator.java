
package compsciia;

import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.FormulaEvaluator;                  
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class TermsLocator {
    static int rowNum;
    static int colNum;
    static int spreadSheetNum;
    private boolean sheetNumFound;
    FormulaEvaluator evaluator;
    GeneratorGUI detailsInput;
    NeededDocu excelFile = new NeededDocu();
    JOptionPane panel = new JOptionPane();
     
    public TermsLocator(){
   System.out.println("RUNNING TERMSLOCATOR");
    sheetNumFound = false;
    excelFile = new NeededDocu();        // NeededDocu is the class that have managed the import of excel file
    detailsInput = new GeneratorGUI();  // GeneratorGUI is the class that have managed to user input
    evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
    }
    
    public TermsLocator(String searchTrimester){
   System.out.println("RUNNING TERMSLOCATOR FOR"+searchTrimester);
    sheetNumFound = false;
    excelFile = new NeededDocu();           // NeededDocu is the class that have managed the import of excel file
    evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
    }
     
     public void findSheetNum() {
     while (sheetNumFound == false) {
         if (excelFile.grades.getSheetName(spreadSheetNum).equalsIgnoreCase(detailsInput.sheetName)) { 
             sheetNumFound = true;      // Through 'detailsInput', the variable sheetName that was in GeneratorGUI can be used here
             excelFile.setSheetNum(spreadSheetNum); // setSheetNum is a method from NeededDocu that imports the excel sheet that can be used
         } else {spreadSheetNum++; }
     }
     }
     
     public void findTermLocation() {  
     XSSFSheet sheet = excelFile.sheet;      // Needed spreadsheet from NeededDocu class
            for(Row row:sheet){
                for(Cell cell : row){
                if (evaluator.evaluateInCell(cell).getCellType()==Cell.CELL_TYPE_STRING) {
                        if (cell.getStringCellValue().equals(detailsInput.searchTrimester)) {
                            rowNum = cell.getRow().getRowNum();
                            colNum = cell.getColumnIndex();
                            break;
                        }
                  }
            }
      }   
}
     public void findTermLocation(String searchTrimester) {  
     XSSFSheet sheet = excelFile.sheet;
            for(Row row:sheet){
                for(Cell cell : row){
                if (evaluator.evaluateInCell(cell).getCellType()==Cell.CELL_TYPE_STRING) {
                        if (cell.getStringCellValue().equals(searchTrimester)) {
                            rowNum = cell.getRow().getRowNum();
                            colNum = cell.getColumnIndex();  
                            break;
                        }
                  }
            }
      }
}
     
    
}