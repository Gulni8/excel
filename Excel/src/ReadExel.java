/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author gulni
 */
import java.util.Scanner;
import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadExel {

    private String inputFile;

    public void setInputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    public void read() throws IOException {
        File inputWorkbook = new File(inputFile);
        Workbook w;
       // Scanner search = new Scanner(System.in);
       // System.out.println("What do you want to search?");
       // String x = search.nextLine();
        try {
            w = Workbook.getWorkbook(inputWorkbook);
            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            

             //show all sheet by rows
                for (int i = 0; i < sheet.getRows(); i++) {
                    Cell cell = sheet.getCell(0, i);
                        for (int k = 0; k < sheet.getColumns(); k++) {
                            Cell cel = sheet.getCell(k, i);
                            System.out.print(cel.getContents()+" ");
                        }
                        System.out.println("\n");
                    }
              
                //search by wanted value
        /*   for (int j = 0; j < sheet.getColumns(); j++) {
                for (int i = 0; i < sheet.getRows(); i++) {
                    Cell cell = sheet.getCell(j, i);
                    if (cell.getContents().equalsIgnoreCase(x)) { //finds the cell who has the value and list the row's members
                        for (int k = 0; k < sheet.getColumns(); k++) {
                            Cell cel = sheet.getCell(k, i);
                            System.out.println(cel.getContents());
                        }
                        System.out.println("\n");
                    }
                }
            }*/
            
            //reads all by columns 
          /*  for (int j = 0; j < sheet.getColumns(); j++) {
                for (int i = 0; i < sheet.getRows(); i++) {
                    Cell cell = sheet.getCell(j, i);
                    CellType type = cell.getType();
                    if (type == CellType.LABEL) {
                        System.out.println(cell.getContents());
                    }

                    if (type == CellType.NUMBER) {
                        System.out.println(cell.getContents());
                    }

                } 
            } */ 
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        
        ReadExel test = new ReadExel();
        test.setInputFile("C:\\Users\\gulni\\Desktop\\test.xls");
        test.read();
    }

}
