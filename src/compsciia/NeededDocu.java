/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package compsciia;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NeededDocu extends GeneratorGUI{
    
    static FileInputStream gradebook;
    static XSSFWorkbook grades;
    static XSSFSheet sheet;
    private int sheetNum;

    public FileInputStream getGradebook() {return gradebook;}
    public XSSFWorkbook getWorkbook() {return grades;}
    
    public void setSheetNum(int num) {                                           // Method with a parameter to set the spreadsheet
        sheetNum = num;
        sheet = grades.getSheetAt(sheetNum);                                    // Spreadsheet specificied
            }
    public XSSFSheet getSheet() {return sheet;}
    
    public void RetrieveExcel() throws IOException{
        gradebook = new FileInputStream(new File(excelLocation+"\\"+file));
        grades = new XSSFWorkbook(gradebook);                                   // Gradebook created  
    }
 }


