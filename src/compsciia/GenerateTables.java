/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package compsciia;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author lenovo
 */
public class GenerateTables {
    
    TitleCounter titles;
    private int colCheckCount;
    StudentInfo getArray;
    String[][] studentInfo;
    String[][] classAvg;
    TermsLocator terms;
    
    public GenerateTables() {
        System.out.println("-------------------------------------------------RUN GENERATE TABLES-------------------------------------------------");
        titles = new TitleCounter();
        colCheckCount = titles.CountTitle();
        getArray = new StudentInfo();
        studentInfo = getArray.studentArray;
        classAvg = getArray.classAvg;
        
    }
    
    public GenerateTables(String searchTrimester) {
        System.out.println("-------------------------------------------------RUN GENERATE TABLES FOR PARAMETERS-------------------------------------------------         "+searchTrimester);
        terms = new TermsLocator(searchTrimester);
        terms.findTermLocation(searchTrimester);
        System.out.println("//////////////////////////                                                         terms found location at "+terms.rowNum+" and "+terms.colNum+"                                     ///////////////////////////////////////");
        titles = new TitleCounter(searchTrimester);
        colCheckCount = titles.CountTitle(terms.rowNum, terms.colNum);
        getArray = new StudentInfo(searchTrimester);
        studentInfo = getArray.getStudentInfo(terms.rowNum,terms.colNum,titles.startTitleNum, searchTrimester);
        classAvg = getArray.classAvg;
    }
        
        public void makeTables(XWPFDocument studentProfile, int studentPrintCount) {
                XWPFTable studentGrades = studentProfile.createTable();
                studentGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow titles = studentGrades.getRow(0);
                XWPFTableRow row2Grades = studentGrades.createRow();
                
                //printing tables
                for (int rowNum = 0; rowNum < 2; rowNum++){                         // rowNum for the table
                         for (int colNum = 0; colNum < colCheckCount; colNum++){   // colNum for the table
                             if ((rowNum == 0)&&(colNum == 0)){
                            titles.getCell(0).setText(studentInfo[0][0]);          // titles is the first row of the table
                            titles.setHeight(200);                                  // studentInfo is the 2D array. [0][0] always contains 'Name'
                             } else if(rowNum==0){
                                    titles.addNewTableCell().setText(studentInfo[0][colNum]); }
                             else {
                                 if ((rowNum==1)&&(colNum==0)){
                                     row2Grades.getCell(0).setText(studentInfo[studentPrintCount][colNum]);
                                     row2Grades.setHeight(200);
                                 } else {
                                    row2Grades.addNewTableCell().setText(studentInfo[studentPrintCount][colNum]);}
                             }
                             }
                }
                XWPFTable averageGrades = studentProfile.createTable();
                averageGrades.setCellMargins(200,50,50, 200);
                XWPFTableRow avgGrades = averageGrades.getRow(0);
               for (int colNum = 0; colNum < colCheckCount; colNum++){
                   if (colNum==0){
                   avgGrades.getCell(0).setText("Class average");
                   } else {
                       avgGrades.addNewTableCell().setText(classAvg[0][colNum]);
                   }
               } 
    }
}

