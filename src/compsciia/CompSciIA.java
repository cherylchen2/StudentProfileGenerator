package compsciia;

import java.io.IOException;
import javax.swing.JOptionPane;

public class CompSciIA {
    
    public void run() throws IOException{
        JOptionPane panel = new JOptionPane();
        NeededDocu inputFiles = new NeededDocu();
        GeneratorGUI inputValues = new GeneratorGUI();

            System.out.println("here");
            try {
            inputFiles.RetrieveExcel();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(panel, "Oops! An error has occured! Please check if you have entered file address, file name with extension and spreadsheet name correct and checked a trimester correctly!");
                System.exit(0);
            }
            System.out.println("retrieved excel");
           
            try {
            TermsLocator terms = new TermsLocator();
            terms.findSheetNum();
            terms.findTermLocation(); }
            catch (Exception e) {
                JOptionPane.showMessageDialog(panel, "Oops! An error has occured while trying to find and locate required term! Please make sure to have all terms entered in!");
                System.exit(0);
            }

        try {
        GenerateWord generate = new GenerateWord();
        generate.generateWord();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(panel, "Oops! An error has occured! Please check that your excel format is correct with rows aligned and same student number for each term!");
            System.exit(0);
        }
        System.out.println("line before generateword");
        
        System.out.println("generateword");
        System.exit(0);
    }
}
/*
import java.io.File;                                                                    import java.io.FileInputStream;                                                        import java.io.IOException;                                               
import org.apache.poi.ss.usermodel.Cell;                                import org.apache.poi.ss.usermodel.FormulaEvaluator;                import org.apache.poi.ss.usermodel.Row;                   
import org.apache.poi.xssf.usermodel.XSSFSheet;                 import org.apache.poi.xssf.usermodel.XSSFWorkbook;                  import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.FileOutputStream;                                            import org.apache.poi.xwpf.usermodel.ParagraphAlignment;      import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;     import org.apache.poi.xwpf.usermodel.XWPFRun;                          import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;  
public class CompSciIA {

    public static void main(String[] args) throws IOException{
        
        FileInputStream gradebook = new FileInputStream(new File("ForCherylGrades.xlsx"));
        XSSFWorkbook grades = new XSSFWorkbook(gradebook);
        XSSFSheet sheet = grades.getSheetAt(0);
        FormulaEvaluator forlulaEvaluator = grades.getCreationHelper().createFormulaEvaluator();
        final int tNameGap = 2;
        
        for(Row row:sheet){
            for(Cell cell : row){
                if (forlulaEvaluator.evaluateInCell(cell).getCellType()==Cell.CELL_TYPE_STRING) {
                        if (cell.getStringCellValue().equals("Trimester 1")) {
                            final int t1RowNum = cell.getRow().getRowNum();
                            final int t1ColNum = cell.getColumnIndex();
                            int studentNumber = 0;
                            boolean done = false;
                            int count = 2;
                            System.out.println(t1RowNum+" and "+t1ColNum);
                            
                            //finding number of students
                            Row r = sheet.getRow(t1RowNum);   if (r==null){ r = sheet.createRow(t1RowNum); }
                            Cell c = r.getCell(t1ColNum);       if (c==null){ c = r.createCell(t1ColNum);}
                            while (done==false) {
                                Row studentRow = sheet.getRow(t1RowNum+count);   if (studentRow==null){ studentRow = sheet.createRow(t1RowNum+count); }
                                Cell studentCol = studentRow.getCell(t1ColNum);       if (studentCol==null){ studentCol = row.createCell(t1ColNum);}
                                if (((forlulaEvaluator.evaluateInCell(studentCol).getCellType())!=Cell.CELL_TYPE_BLANK ) && !(studentCol.getStringCellValue().equals("Average"))){
                                    System.out.println(studentCol);
                                    studentNumber++;
                                    count++;
                                } else {    done = true; count =2; }
                                                } 
                            System.out.println(studentNumber);
                            
                            //finding number of titles
                            done = false;
                            int colCheckCount = 0;
                             while (done==false) {
                                Row colCheckRow = sheet.getRow(t1RowNum+1);   if (colCheckRow==null){ colCheckRow = sheet.createRow(t1RowNum+1); }
                                Cell colCheckCol = colCheckRow.getCell(t1ColNum+colCheckCount);       if (colCheckCol==null){ colCheckCol = colCheckRow.createCell(t1ColNum+colCheckCount);}
                                if ((forlulaEvaluator.evaluateInCell(colCheckCol).getCellType())!=Cell.CELL_TYPE_BLANK){
                                    System.out.println(colCheckCol);
                                    colCheckCount++;
                                } else {    done = true;}
                                                } 
                            System.out.println(colCheckCount);
                            
                            
                            //creating arrays and putting info into array
                            done = false;

                            String studentInfoT1[][] = new String[studentNumber+1][colCheckCount];

                            while (done==false){
                                for (int arrayRow = 0; arrayRow <studentNumber+1 ; arrayRow++){
                                    int searchTest = t1ColNum;
                                    for (int arrayCol = 0 ; arrayCol <colCheckCount ; arrayCol++){
                                        System.out.println(arrayRow+" "+arrayCol);
                                        Row studentRow = sheet.getRow(t1RowNum+arrayRow+tNameGap-1);                              if (studentRow==null){ studentRow = sheet.createRow(t1RowNum+arrayRow+tNameGap-1); }
                                        Cell studentCol = studentRow.getCell(searchTest);                                               if (studentCol==null){ studentCol = studentRow.createCell(searchTest);}
                                        if (forlulaEvaluator.evaluateInCell(studentCol).getCellType()==Cell.CELL_TYPE_STRING){
                                                studentInfoT1[arrayRow][arrayCol] = studentCol.getStringCellValue();
                                                searchTest++;
                                        } else if(forlulaEvaluator.evaluateInCell(studentCol).getCellType()==Cell.CELL_TYPE_NUMERIC){
                                            double tempSave;
                                            if (studentCol.getCellStyle().getDataFormatString().contains("%")){
                                                System.out.print("im here");
                                                tempSave = studentCol.getNumericCellValue()*100;
                                            } else {
                                               tempSave = studentCol.getNumericCellValue(); }
                                                studentInfoT1[arrayRow][arrayCol] = Double.toString(tempSave);
                                                searchTest++;
                                        } else {
                                            studentInfoT1[arrayRow][arrayCol] = " null ";
                                            searchTest++;
                                        }
                                        

                                    }
                                }
                                done = true;
                            }
                            
                            for (int rowprint = 0; rowprint < studentNumber ; rowprint++) {
                                System.out.println();
                                for (int colprint =0 ; colprint < colCheckCount ; colprint++){
                                    System.out.print(studentInfoT1[rowprint][colprint]+" ");
                                }
                            }
                            
           //Making word document                 
            for (int studentPrintCount = 1; studentPrintCount < studentNumber ; studentPrintCount++){
                XWPFDocument studentProfile = new XWPFDocument();
                
                XWPFParagraph paragraph = studentProfile.createParagraph();
                XWPFRun run = paragraph.createRun();
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                
                run.setText("Grade 8 Term 1 ");
                run.setText(" Report for "+studentInfoT1[studentPrintCount][0]);
                run.addBreak();
                run.setText("Term 1 Grade:");
                
                XWPFParagraph paragraph2 = studentProfile.createParagraph();
                XWPFRun run2 = paragraph2.createRun();
                run2.setText("Here is the table: ");
                
                XWPFTable studentGrades = studentProfile.createTable();
                XWPFTableRow titles = studentGrades.getRow(0);
                XWPFTableRow row2Grades = studentGrades.createRow();
                //printing tables
                for (int rowNum = 0; rowNum < 2; rowNum++){
                         for (int colNum = 0; colNum < colCheckCount; colNum++){
                             if ((rowNum == 0)&&(colNum == 0)){
                            titles.getCell(0).setText(studentInfoT1[0][0]);
                             } else if(rowNum==0){
                                    titles.addNewTableCell().setText(studentInfoT1[0][colNum]); }
                             else {
                                 if ((rowNum==1)&&(colNum==0)){
                                     row2Grades.getCell(0).setText(studentInfoT1[studentPrintCount][colNum]);
                                 } else {
                                    row2Grades.addNewTableCell().setText(studentInfoT1[studentPrintCount][colNum]); }
                             }
                             }
                }
                
                
                    try {
                FileOutputStream output = new FileOutputStream((studentInfoT1[studentPrintCount][0])+".docx");
                studentProfile.write(output);
                studentProfile.close();
            } catch (Exception e){
                e.printStackTrace();
                        break;
                        
            }
                }
            }
        }
        
    }
        
        }
    }
}
*/

    
    