/* compsciia;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ClassAverage {
    TitleCounter titles;
    int titleNumber;
    String classAverage[][];
    StudentCounter students;
    TermsLocator locator;
    int avgLocationRow;
    int avgLocationCol;
    NeededDocu excelFile;
     DataInput inputData;
     
     public ClassAverage() {
         System.out.println("RUNNING CLASSAVERAGE");
         titles = new TitleCounter();
        titleNumber = titles.CountTitle();
        System.out.println("CLASS AVERAGE AND THE TITLENUMBER IS:   "+titleNumber);
         classAverage = new String[1][titleNumber];
        students = new StudentCounter();
        locator = new TermsLocator();
        avgLocationRow = students.countStudents()+locator.rowNum+2;
        avgLocationCol = locator.colNum;
        excelFile = new NeededDocu();
        inputData = new DataInput();
     }
     
     public ClassAverage(String searchTrimester) {
         System.out.println("RUNNING CLASSAVERAGE  -- Search Trimester");
         titles = new TitleCounter(searchTrimester);
        titleNumber = titles.CountTitle();
         classAverage = new String[1][titleNumber];
         locator = new TermsLocator();
         locator.findTermLocation(searchTrimester);
        students = new StudentCounter( );
        avgLocationRow = students.countStudents(locator.rowNum, locator.colNum)+locator.rowNum+2;
        avgLocationCol = locator.colNum;
        excelFile = new NeededDocu();
        inputData = new DataInput();
     }
     
    public String[][] classAverage(){
        for (int count = 0; count < titleNumber; count++){
             Row studentRow = excelFile.sheet.getRow(avgLocationRow);                              if (studentRow==null){ studentRow = excelFile.sheet.createRow(avgLocationRow); }
             Cell studentCol = studentRow.getCell(avgLocationCol+count);             if (studentCol==null){ studentCol = studentRow.createCell(avgLocationCol+count);}
            classAverage[0][count] = inputData.getData(avgLocationRow,(avgLocationCol+count));
        }
        return classAverage;
        }
    
    public String[][] classAverage(String searchTrimester){
        for (int count = 0; count < titleNumber; count++){
             Row studentRow = excelFile.sheet.getRow(avgLocationRow);                              if (studentRow==null){ studentRow = excelFile.sheet.createRow(avgLocationRow); }
             Cell studentCol = studentRow.getCell(avgLocationCol+count);             if (studentCol==null){ studentCol = studentRow.createCell(avgLocationCol+count);}
            classAverage[0][count] = inputData.getData(avgLocationRow,(avgLocationCol+count));
        }
        return classAverage;
        }
    }

*/