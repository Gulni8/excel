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
import java.io.IOException;
import java.util.Locale;
import java.util.Scanner;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import writeexcel.WriteExcel;

public class WriteWscanner {
    
    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;

   
    
public void setOutputFile(String inputFile) {
    this.inputFile = inputFile;
    }

    public void write() throws IOException, WriteException {
        File file = new File(inputFile);
        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));
        
        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        createContent(excelSheet);

        workbook.write();
        workbook.close();
    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);
        // Lets automatically wrap the cells
        times.setWrap(true);

        // create create a bold font with unterlines
        WritableFont times10ptBoldUnderline = new WritableFont(
                WritableFont.TIMES, 10, WritableFont.BOLD, false,
                UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        // Lets automatically wrap the cells
        timesBoldUnderline.setWrap(true);


        // Write a few headers
        addCaption(sheet, 0, 0, "StudentId");
        addCaption(sheet, 1, 0, "StudentName");
        addCaption(sheet, 2, 0, "ProgramId");
        addCaption(sheet, 3, 0, "ProgramName");

    }

    private void createContent(WritableSheet sheet) throws WriteException,
            RowsExceededException {
        
        int lastRow = sheet.getRows();
        Scanner add = new Scanner(System.in);
        System.out.println("student id?");
        String x1 = add.nextLine();
        addLabel(sheet, 0, lastRow, x1); //sheet column row input
        
        System.out.println("student name?");
        String x2 = add.nextLine();
        addLabel(sheet, 1, lastRow, x2);
        
        System.out.println("program id?");
        String x3 = add.nextLine();
        addLabel(sheet, 2, lastRow, x3);
        
        System.out.println("program name?");
        String x4 = add.nextLine();
        addLabel(sheet, 3, lastRow, x4);
        
        
        
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
            Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    public static void main(String[] args) throws WriteException, IOException {
        WriteExcel test = new WriteExcel();
        test.setOutputFile("C:\\Users\\gulni\\Desktop\\testWrite1.xls");
        test.write();
        System.out
                .println("Please check the result file under C:\\Users\\gulni\\Desktop\\testWrite.xls ");
    }
    
}
