package compsciia;

import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class StudentInfo{
    private boolean done = false;
    private boolean done2 = false;
    DataInput inputData;
    TermsLocator terms;
    TitleCounter titles;
    NeededDocu excelFile;
    private int titleRowNum;
    FormulaEvaluator evaluator;
    String cellValue;
    String testText;
    static String[][] classAvg;
    static String[][] studentArray;
    
    public StudentInfo() {
        inputData = new DataInput();
        titles = new TitleCounter();
        titleRowNum = terms.rowNum-1;
        excelFile = new NeededDocu();
        evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
        testText = "test";
    }
    
    public StudentInfo(String searchTrimester) {
        inputData = new DataInput();
        titles = new TitleCounter(searchTrimester);
        terms = new TermsLocator(searchTrimester);
        terms.findTermLocation(searchTrimester);
        titleRowNum = terms.rowNum-1;
        excelFile = new NeededDocu();
        evaluator = excelFile.grades.getCreationHelper().createFormulaEvaluator();
        testText = "test";
    }
    
    public String[][] getStudentInfo() {
    String studentArray[][];
    studentArray = makeStudentInfo();
    return studentArray; }
    
    public String[][] getStudentInfo(int row, int nameCol, int startCol, String searchTrimester){
    String studentArray[][];
    studentArray = makeStudentInfo(row, nameCol, startCol, searchTrimester);
    return studentArray; }
    
    private String[][] makeStudentInfo() {
        StudentCounter studentCounter = new StudentCounter();
        int studentNumber = studentCounter.countStudents();
          terms = new TermsLocator();
         int row = terms.rowNum;
        int colCheckCount = titles.CountTitle();
        String studentInfo[][] = new String[studentNumber+1][colCheckCount];
        classAvg = new String[1][colCheckCount];
        final int tNameGap = 2;
        
        while (done == false){
            for (int arrayRow = 0; arrayRow < studentNumber+1; arrayRow++){
                int nameCol = terms.colNum;
                int startCol = titles.startTitleNum;
                done2 = false;
                for (int arrayCol = 0; arrayCol< colCheckCount; arrayCol++){
                    int rowNum = row+arrayRow+tNameGap-1;
                    
                    if (done2 == false) {
                        studentInfo[arrayRow][arrayCol] = inputData.getData(rowNum, nameCol);
                        classAvg[0][arrayCol] = inputData.getData((row+tNameGap+studentNumber), nameCol);
                        done2 = true;
                    } else {
                        studentInfo[arrayRow][arrayCol] = inputData.getData(rowNum, startCol);
                        classAvg[0][arrayCol] = inputData.getData((row+tNameGap+studentNumber),startCol);
                        startCol++;
                    } 
                }
            }
            done = true;
        }
        System.out.println("RUNNING STUDENT INFO WHERE DATA INPUT IS USED");
        
        for (int arow = 0; arow < studentNumber+1 ; arow++) {
            System.out.print("/n");
            for (int acol = 0; acol < colCheckCount ; acol++){
                System.out.print(studentInfo[arow][acol]);
            }
        }
        studentArray = studentInfo;
        return studentInfo;
   }
    
        private String[][] makeStudentInfo(int row, int nameCol, int startCol, String searchTrimester) {
            int actualStartCol = startCol;
            System.out.println("++++++++++++++++++++++++++++++++++++++                  row: "+row);
        StudentCounter studentCounter = new StudentCounter(searchTrimester);
        int studentNumber = studentCounter.countStudents();
        int colCheckCount = titles.CountTitle(row, terms.colNum);
        String studentInfo[][] = new String[studentNumber+1][colCheckCount];
        classAvg = new String[1][colCheckCount];
        System.out.println("studentInfo[][] ------------------------------------> [][] is     ["+(studentNumber+1)+"]["+colCheckCount);
        final int tNameGap = 2;
        
        while (done == false){
            for (int arrayRow = 0; arrayRow < studentNumber+1; arrayRow++){
                done2 = false;
                startCol = actualStartCol;
                for (int arrayCol = 0; arrayCol< colCheckCount; arrayCol++){
                    int rowNum = row+arrayRow+tNameGap-1;
                    
                    if (done2 == false) {
                        studentInfo[arrayRow][arrayCol] = inputData.getData(rowNum, nameCol);
                        System.out.println("OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"+studentInfo[0][0]+"                                                                                       "+rowNum+"                               "+nameCol);
                        classAvg[0][0] = inputData.getData((row+tNameGap+studentNumber), nameCol);
                        done2 = true;
                    } else {
                        studentInfo[arrayRow][arrayCol] = inputData.getData(rowNum, startCol);
                        classAvg[0][arrayCol] = inputData.getData((row+tNameGap+studentNumber),startCol);
                        startCol++;
                    } 
                }
            }
            done = true;
        }
        System.out.println("RUNNING STUDENT INFO WHERE DATA INPUT IS USED");
        
        for (int arow = 0; arow < studentNumber+1 ; arow++) {
            System.out.print("\n");
            for (int acol = 0; acol < colCheckCount ; acol++){
                System.out.print(studentInfo[arow][acol]);
            }
        }
        studentArray = studentInfo;
        return studentInfo;
   }
}
        
/*
Row titleRow = excelFile.sheet.getRow(titleRowNum);                              if (titleRow==null){ titleRow = excelFile.sheet.createRow(titleRowNum); }
                     Cell title = titleRow.getCell(colNum);                                                           if (title==null){ title = titleRow.createCell(colNum);}
          
                    if (evaluator.evaluateInCell(title).getCellType()!=Cell.CELL_TYPE_BLANK) 
*/