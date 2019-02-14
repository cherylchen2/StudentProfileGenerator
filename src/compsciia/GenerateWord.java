/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package compsciia;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import javax.swing.JOptionPane;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class GenerateWord {
        JOptionPane panel = new JOptionPane();
        StudentCounter getNumber;
        private int studentNumber;
        StudentInfo getArray;
        String studentInfo[][];
        TitleCounter titles;
        private int colCheckCount;
        GeneratorGUI getName;
        private String sheetName;
        private String trimester;

        private String[][] classAvgGrades;
        GenerateTables getTable;
        XWPFDocument studentProfile;
       
        public GenerateWord() {
            System.out.println("RUNNING GENERATEWORD");
            StudentCounter getNumber = new StudentCounter();
            System.out.println("RUNNING GENERATEWORD 2");
            studentNumber = getNumber.countStudents();
            System.out.println("RUNNING GENERATEWORD 3");
            getArray = new StudentInfo();
            System.out.println("RUNNING GENERATEWORD 4");
            studentInfo = getArray.getStudentInfo();
            System.out.println("instantiated objects");
            titles = new TitleCounter();
            colCheckCount = titles.CountTitle();
            getName = new GeneratorGUI();
            sheetName = getName.sheetName;
            trimester = getName.searchTrimester;
            //ClassAverage classAvg = new ClassAverage();
            //classAvgGrades = classAvg.classAverage();
            classAvgGrades = getArray.classAvg;
            System.out.println("instantiated objects 2");
            getTable = new GenerateTables();

        }
        
    public void generateWord() {
        for (int studentPrintCount = 1; studentPrintCount <= studentNumber ; studentPrintCount++){
                System.out.println("created studentProfile in GenerateWord");
                System.out.println("Studentprintcount: "+studentPrintCount+"                       studentNumber : "+studentNumber);
                System.out.println("------------------------------------------------------>"+colCheckCount);
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy");
                LocalDate localDate = LocalDate.now();
  
                studentProfile = new XWPFDocument();
                XWPFParagraph paragraph = studentProfile.createParagraph();
                XWPFRun centerText = paragraph.createRun();
                paragraph.setAlignment(ParagraphAlignment.CENTER);
                
                centerText.setBold(true);
                centerText.setText(dtf.format(localDate)+" "+sheetName+" "+trimester);
                centerText.addBreak();
                centerText.setText(trimester+" Math Grade:");
                
                XWPFParagraph paragraph2 = studentProfile.createParagraph();
                XWPFRun leftText = paragraph2.createRun();
                paragraph2.setAlignment(ParagraphAlignment.LEFT);
                leftText.setText("Report for "+studentInfo[studentPrintCount][0]);
                leftText.addBreak();
                leftText.setText(trimester+" Grades: ");
                getTable.makeTables(studentProfile,studentPrintCount);
                
                XWPFParagraph belowTable = studentProfile.createParagraph();
                XWPFRun belowTableRun = belowTable.createRun();
                belowTableRun.addCarriageReturn();
                
                if (getName.searchTrimester.equals("Trimester 2")) {
                    XWPFParagraph secondParagraph = studentProfile.createParagraph();
                    XWPFRun tableTitle = secondParagraph.createRun();
                    secondParagraph.setAlignment(ParagraphAlignment.LEFT);
                    System.out.println("============Running for for Trimester 2 pringint Trimester 1 table");
                    tableTitle.setText("Trimester 1 Grades: ");
                    
                    GenerateTables trimester1table = new GenerateTables("Trimester 1");
                    trimester1table.makeTables(studentProfile,studentPrintCount);
                    tableTitle.addCarriageReturn();
                }
                
                if (getName.searchTrimester.equals("Trimester 3")) {
                    XWPFParagraph secondParagraph = studentProfile.createParagraph();
                    XWPFRun tableTitle = secondParagraph.createRun();
                    secondParagraph.setAlignment(ParagraphAlignment.LEFT);
                    System.out.println("===============Running for Trimester 3 printing Trimester 2 and 3 table");
                    tableTitle.setText("Trimester 1 Grades: ");
                    
                    GenerateTables trimester1table = new GenerateTables("Trimester 1");
                    trimester1table.makeTables(studentProfile, studentPrintCount);
                    
                    XWPFParagraph thirdParagraph = studentProfile.createParagraph();
                    XWPFRun tableTitle2 = thirdParagraph.createRun();
                    thirdParagraph.setAlignment(ParagraphAlignment.LEFT);
                    
                    tableTitle2.addCarriageReturn();
                    tableTitle2.setText("Trimester 2 Grades: ");
                    GenerateTables trimester2table = new GenerateTables("Trimester 2");
                    trimester2table.makeTables(studentProfile, studentPrintCount);                    
                }
                
                XWPFParagraph commentPar = studentProfile.createParagraph();
                XWPFRun commentRun = commentPar.createRun();
                commentPar.setAlignment(ParagraphAlignment.LEFT);
                commentRun.addCarriageReturn();
                commentRun.setText("Comment:");
                
                XWPFParagraph teachername = studentProfile.createParagraph();
                XWPFRun teacherRun = teachername.createRun();
                teachername.setAlignment(ParagraphAlignment.LEFT);
                teacherRun.addCarriageReturn();
                teacherRun.setText("Teacher: Ms. Chia");
                 try {
                FileOutputStream output = new FileOutputStream(getName.excelLocation+"/"+(studentInfo[studentPrintCount][0])+".docx");
                studentProfile.write(output);
                studentProfile.close();
            } catch (Exception e){
                e.printStackTrace();
                        break;
                        
            }
                //getTable.makeTables(studentPrintCount);
                /*
             XWPFTable studentGrades = studentProfile.createTable();
                studentGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow titles = studentGrades.getRow(0);
                XWPFTableRow row2Grades = studentGrades.createRow();
                
                //printing tables
               for (int rowNum = 0; rowNum < 2; rowNum++){
                         for (int colNum = 0; colNum < colCheckCount; colNum++){
                             if ((rowNum == 0)&&(colNum == 0)){
                            titles.getCell(0).setText(studentInfo[0][0]);
                            titles.setHeight(200);
                            System.out.print(studentInfo[0][0]+" ");
                             } else if(rowNum==0){
                                    titles.addNewTableCell().setText(studentInfo[0][colNum]); 
                                    System.out.print(studentInfo[0][colNum]+" ");}
                             else {
                                 if ((rowNum==1)&&(colNum==0)){
                                     row2Grades.getCell(0).setText(studentInfo[studentPrintCount][colNum]);
                                     row2Grades.setHeight(200);
                                     System.out.print("|n"+studentInfo[studentPrintCount][colNum]+" ");
                                 } else {
                                    row2Grades.addNewTableCell().setText(studentInfo[studentPrintCount][colNum]);
                                    System.out.print(studentInfo[studentPrintCount][colNum]+" ");}
                             }
                             }
                }
                XWPFTable averageGrades = studentProfile.createTable();
                averageGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow avgGrades = averageGrades.getRow(0);
                XWPFTableRow gradeRow = averageGrades.createRow();
               for (int colNum = 0; colNum < colCheckCount; colNum++){
                   if (colNum==0){
                   avgGrades.getCell(0).setText(getArray.classAvg[0][colNum]);
                   } else {
                       avgGrades.addNewTableCell().setText(getArray.classAvg[0][colNum]);
                   }
               }
                    try {
                FileOutputStream output = new FileOutputStream((studentInfo[studentPrintCount][0])+".docx");
                studentProfile.write(output);
                studentProfile.close();
            } catch (Exception e){
                e.printStackTrace();
                        break;
                        
            }
                }
            */}
        }
}
       

        /*
     for (int studentPrintCount = 1; studentPrintCount <= studentNumber ; studentPrintCount++){
                System.out.println("created studentProfile in GenerateWord");
                System.out.println("Studentprintcount: "+studentPrintCount+"                       studentNumber : "+studentNumber);
                System.out.println("------------------------------------------------------>"+colCheckCount);
                
                XWPFParagraph paragraph = studentProfile.createParagraph();
                XWPFRun centerText = paragraph.createRun();
                paragraph.setAlignment(ParagraphAlignment.CENTER);
                
                centerText.setText(sheetName+trimester);
                centerText.setText(trimester+" Grade:");
                
                XWPFParagraph paragraph2 = studentProfile.createParagraph();
                XWPFRun leftText = paragraph2.createRun();
                paragraph2.setAlignment(ParagraphAlignment.LEFT);
                leftText.setText(" Report for "+studentInfo[studentPrintCount][0]);
                leftText.addBreak();
                leftText.setText("Here is the table: ");
                //getTable.makeTables(studentPrintCount);
                
             XWPFTable studentGrades = studentProfile.createTable();
                studentGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow titles = studentGrades.getRow(0);
                XWPFTableRow row2Grades = studentGrades.createRow();
                
                //printing tables
               for (int rowNum = 0; rowNum < 2; rowNum++){
                         for (int colNum = 0; colNum < colCheckCount; colNum++){
                             if ((rowNum == 0)&&(colNum == 0)){
                            titles.getCell(0).setText(studentInfo[0][0]);
                            titles.setHeight(200);
                            System.out.print(studentInfo[0][0]+" ");
                             } else if(rowNum==0){
                                    titles.addNewTableCell().setText(studentInfo[0][colNum]); 
                                    System.out.print(studentInfo[0][colNum]+" ");}
                             else {
                                 if ((rowNum==1)&&(colNum==0)){
                                     row2Grades.getCell(0).setText(studentInfo[studentPrintCount][colNum]);
                                     row2Grades.setHeight(200);
                                     System.out.print("|n"+studentInfo[studentPrintCount][colNum]+" ");
                                 } else {
                                    row2Grades.addNewTableCell().setText(studentInfo[studentPrintCount][colNum]);
                                    System.out.print(studentInfo[studentPrintCount][colNum]+" ");}
                             }
                             }
                }
                XWPFTable averageGrades = studentProfile.createTable();
                averageGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow avgGrades = averageGrades.getRow(0);
                XWPFTableRow gradeRow = averageGrades.createRow();
               for (int colNum = 0; colNum < colCheckCount; colNum++){
                   if (colNum==0){
                   avgGrades.getCell(0).setText(getArray.classAvg[0][colNum]);
                   } else {
                       avgGrades.addNewTableCell().setText(getArray.classAvg[0][colNum]);
                   }
               }
                    try {
                FileOutputStream output = new FileOutputStream((studentInfo[studentPrintCount][0])+".docx");
                studentProfile.write(output);
                studentProfile.close();
            } catch (Exception e){
                e.printStackTrace();
                        break;
                        
            }
                } */
