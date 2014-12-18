package Hulu;
import java.io.*;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLS {

    static void convertToXlsx(File inputFile, File outputFile)
    {
        // For storing data into CSV files
        StringBuffer cellValue = new StringBuffer();
        try
        {
            FileOutputStream fos = new FileOutputStream(outputFile);

            // Get the workbook instance for XLSX file
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(inputFile));

            // Get first sheet from the workbook
            XSSFSheet sheet = wb.getSheetAt(0);

            Row row;
            Cell cell;

            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext())
            {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext())
                {
                    cell = cellIterator.next();

                    switch (cell.getCellType())
                    {

                        case Cell.CELL_TYPE_BOOLEAN:
                            cellValue.append(cell.getBooleanCellValue() + ",");
                            break;

                        case Cell.CELL_TYPE_NUMERIC:
                            cellValue.append(cell.getNumericCellValue() + ",");
                            break;

                        case Cell.CELL_TYPE_STRING:
                            cellValue.append(cell.getStringCellValue() + ",");
                            break;

                        case Cell.CELL_TYPE_BLANK:
                            cellValue.append("" + ",");
                            break;

                        default:
                            cellValue.append(cell + ",");

                    }
                }
            }

            fos.write(cellValue.toString().getBytes());
            fos.close();

        }
        catch (Exception e)
        {
            System.err.println("Exception :" + e.getMessage());
        }
    }

    static void convertToXls(File inputFile, File outputFile)
    {
// For storing data into CSV files
        StringBuffer cellDData = new StringBuffer();
        try
        {
            FileOutputStream fos = new FileOutputStream(outputFile);

            // Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
            // Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            Cell cell;
            Row row;
            //int j = sheet.getLastRowNum();
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();
            //int count =1;
            while (rowIterator.hasNext())
            {
                row = rowIterator.next();
                if(row.getCell(0).getStringCellValue()=="")break;
                String first = row.getCell(0).getStringCellValue();
                int sec = (int)row.getCell(7).getNumericCellValue();
                String val = first+"-"+sec;
                cellDData.append(val + ";");
                if(row.getCell(33).getStringCellValue()!=""){
                    String ah = row.getCell(33).getStringCellValue();
                    String t[] = ah.split(",");
                    int sum=1;
                    for(String i:t){
                        String out [] = i.split("-");
                        int diff = Integer.parseInt(out[1])-Integer.parseInt(out[0]);
                        sum+=diff;
                    }
                    cellDData.append(sum + ";");
                    //sum=1;
                }


                // For each row, iterate through each columns
                int colval=0;
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext())
                {
                    cell = cellIterator.next();
                    int n = cell.getColumnIndex();

                    if(n == 4 ||n==18 || n==24 ||n==30) {
                        switch (cell.getCellType()) {

                            case Cell.CELL_TYPE_BOOLEAN:
                                cellDData.append(cell.getBooleanCellValue() + ";");
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                                cellDData.append((int)cell.getNumericCellValue() + ";");
                                colval = (int)cell.getNumericCellValue();
                                break;

                            case Cell.CELL_TYPE_STRING:
                                cellDData.append(cell.getStringCellValue() + ";");
                                break;


                            case Cell.CELL_TYPE_BLANK:
                                //cellDData.append("" + ",");
                                break;

                            default:
                                //cellDData.append(cell + ",");
                        }
                    }
                    if(n == 34 ||n==35|| n==36) {
                        switch (cell.getCellType()) {

                            case Cell.CELL_TYPE_BOOLEAN:
                                cellDData.append(cell.getBooleanCellValue() + ";");
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                                cellDData.append((int) cell.getNumericCellValue() + ";");
                                break;

                            case Cell.CELL_TYPE_STRING:
                                if(n==36) {
                                    String [] data = cell.getStringCellValue().split(",");

                                    int value = getMax(data,colval);
                                    cellDData.append(cell.getStringCellValue()+ ";");
                                    cellDData.append(value);
                                }
                                break;


                            case Cell.CELL_TYPE_BLANK:
                                //cellDData.append("" + ",");
                                break;

                            default:
                                //cellDData.append(cell + ",");
                        }
                    }


                }

                    cellDData.append(System.getProperty("line.separator"));

            }

            fos.write(cellDData.toString().getBytes());
            fos.close();

        }
        catch (FileNotFoundException e)
        {
            System.err.println("Exception" + e.getMessage());
        }
        catch (IOException e)
        {
            System.err.println("Exception" + e.getMessage());
        }
    }

    public static int getMax(String []d,int col){
        int diff = 1000000;
        int value =0;
        int curr=0;
        for(String s:d){
            int first = Integer.parseInt(s);
            curr = first-col;
            //value = first;
            if(curr>0 && curr<diff){
                diff = curr;
                value = first;
            }
        }
        return value;
    }
    public static void main(String[] args)
    {
        Scanner sc = new Scanner(System.in);
        System.out.println("Please enter Excel File Path");
        String path1 = sc.nextLine();
        System.out.println("Please enter Output File Path");
        String path2 = sc.nextLine();
        File inputFile = new File(path1);
        File outputFile = new File(path2);
        System.out.println("Please enter Date");
        String date = sc.nextLine();
        System.out.println("Please enter Date");
        String time = sc.nextLine();
        //File inputFile = new File("C://Users/praghav.scslt111/Desktop/Code/Excel.xls");
        //File outputFile = new File("C://Users/praghav.scslt111/Desktop/Code/output1.csv");

        convertToXls(inputFile, outputFile);

    }
}
