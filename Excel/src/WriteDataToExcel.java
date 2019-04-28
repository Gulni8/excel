/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author gulni
 */
import java.io.File;
import jxl.Workbook;

import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class WriteDataToExcel {

    static Workbook wbook;
    static WritableWorkbook wwbCopy;
    static String ExecutedTestCasesSheet;
    static WritableSheet shSheet;

    public void readExcel() {
        try {
            wbook = Workbook.getWorkbook(new File("C:\\Users\\gulni\\Deskt\"C:\\\\Usersop\\test1.xls")); //real file
            wwbCopy = Workbook.createWorkbook(new File("C:\\Users\\gulni\\Desktop\\testWrite.xls"), wbook); // copy of the input file
            shSheet = wwbCopy.getSheet(0);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void setStudentId(String strSheetName, String stId) throws WriteException {
        
        WritableSheet wshTemp = wwbCopy.getSheet(strSheetName);
        int iRowNumber = wshTemp.getRows();
        int iColumnNumber = wshTemp.getColumns();
            

            Label labTemp = new Label(iColumnNumber-4, iRowNumber, stId);

            try {
                wshTemp.addCell(labTemp);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        public void setStudentName(String strSheetName, String stName) throws WriteException {
            
        WritableSheet wshTemp = wwbCopy.getSheet(strSheetName);
        int iRowNumber = wshTemp.getRows();
        int iColumnNumber = wshTemp.getColumns();
            

            Label labTemp = new Label(iColumnNumber-3, iRowNumber-1, stName);

            try {
                wshTemp.addCell(labTemp);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        public void setProgId(String strSheetName, String progId) throws WriteException {
            
        WritableSheet wshTemp = wwbCopy.getSheet(strSheetName);
        int iRowNumber = wshTemp.getRows();
        int iColumnNumber = wshTemp.getColumns();
        

            Label labTemp = new Label(iColumnNumber-2, iRowNumber-1, progId);

            try {
                wshTemp.addCell(labTemp);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        public void setProgName(String strSheetName, String progName) throws WriteException {
            
        WritableSheet wshTemp = wwbCopy.getSheet(strSheetName);
        int iRowNumber = wshTemp.getRows();
        int iColumnNumber = wshTemp.getColumns();
            

            Label labTemp = new Label(iColumnNumber-1, iRowNumber-1, progName);

            try {
                wshTemp.addCell(labTemp);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    

    public void closeFile() {
        try {
            // Closing the writable work book
            wwbCopy.write();
            wwbCopy.close();

            // Closing the original work book
            wbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws WriteException {
        WriteDataToExcel ds = new WriteDataToExcel();
        ds.readExcel();
        ds.setStudentId("Report", "615002");
        ds.setStudentName("Report", "art");
        ds.setProgId("Report", "13");
        ds.setProgName("Report", "pp");
        ds.closeFile();
    }
}
