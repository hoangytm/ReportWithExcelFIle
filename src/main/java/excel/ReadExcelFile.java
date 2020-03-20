package excel;

import java.io.*;
import java.util.*;


import net.sf.jxls.transformer.XLSTransformer;;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * @author PhanHoang
 * 3/10/2020
 */
public class ReadExcelFile {
    private static final String FILE_PATH = "C:\\Users\\phann\\Desktop\\testExelFile\\templateFile.xlsx";
    private static final String SOURCE_FILE = "C:\\Users\\phann\\Desktop";
    private static final Integer TOTAL_PAGE = 1000;

    public static void main(String[] args) throws Exception {
        export();
    }

    //    public static void exportReportFile()  {
//        try {
//            Customer customer1 = new Customer();
//            customer1.setName("hoang");
//            List<Customer>lstCustomer = new ArrayList<Customer>();
//            lstCustomer.add(customer1);
//          Integer begin, end;
//          String sheetName, tempName;
//          List sheetNames = new ArrayList();
//          List tempNames = new ArrayList();
//          List maps = new ArrayList();
//          // tinh ra so sheet can thiet trong 1 file
//            if(lstCustomer!=null){
//                int sheetNum= (int) Math.ceil(lstCustomer.size()/(double) TOTAL_PAGE);
//                for(int i=0; i<sheetNum;i++){
//                    begin= i*TOTAL_PAGE;
//                    end =(i+1)*TOTAL_PAGE;
//                    if(begin >= lstCustomer.size()) break;
//                    if(end>=lstCustomer.size()) end = lstCustomer.size();
//                    sheetName = "sheet "+ i;
//                    tempName = "sheet";
//                    sheetNames.add(sheetName);
//                    tempNames.add(tempName);
//                    Map beans = new HashMap();
//                    beans.put("name", "hoang");
//                    beans.put("age","25");
//                    maps.add(beans);
//                }
//
//            }
//
//            XLSTransformer transformer = new XLSTransformer();
//            File file = new File(FILE_PATH);
//            InputStream is = new DataInputStream(new FileInputStream(file));
//
//            String time = String.valueOf(System.currentTimeMillis());
//            String dest = SOURCE_FILE + "\\testExelFile\\" + time + "_JavaBooks.xlsx";
////            Workbook workbook = transformer.transformXLS(FILE_PATH,maps,dest);
//
//            OutputStream outputStream = new BufferedOutputStream(new FileOutputStream(SOURCE_FILE + "\\testExelFile\\" + time + "_JavaBooks.xlsx"));
////            workbook.write(outputStream);
//        }
//        catch (Exception e){
//            e.printStackTrace();
//        }
//    }

    // xuat file excel tu bao cao
    /*
    http://jxls.sourceforge.net/1.x/samples/collectionsample.html
     */
    public static void export() throws Exception {
        try {
            String templateFileName = "src/main/resources/excel/Employees.xls";
            String destFileName = "src/main/resources/excel/test.xls";
            Collection staff = new HashSet();
            staff.add(new Customer(1, "3000", "0.30"));
            staff.add(new Customer(2, "Elsa", ""));
            staff.add(new Customer(3, "Oleg", "s"));
            staff.add(new Customer(4, "Neil", "@"));

            Map beans = new HashMap();
            beans.put("customer", staff);
            XLSTransformer transformer = new XLSTransformer();
            transformer.transformXLS(templateFileName, beans, destFileName);
        } catch (Exception e) {
            e.printStackTrace();
//            throw e;
        }
        System.out.println("success");
    }

    //read excel file
    public static void readExcelFile() {

        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_PATH));
            InputStream content = excelFile;
            Workbook workbook = new HSSFWorkbook(content);
            Sheet datatypeSheet = workbook.getSheetAt(0);
//            DataFormatter fmt = new DataFormatter();
            // iterator will have all rows and We have to loop iterator for reading data
            Iterator<Row> iterator = datatypeSheet.iterator();
            Row firstRow = iterator.next();
            Cell firstCell = firstRow.getCell(0);
            System.out.println(firstCell.getStringCellValue());
            List<Customer> listOfCustomer = new ArrayList<Customer>();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Customer customer = new Customer();
//                customer.setId(Integer.parseInt(fmt.formatCellValue(currentRow.getCell(0))));
                customer.setName(currentRow.getCell(1).getStringCellValue());
                customer.setEmail(currentRow.getCell(2).getStringCellValue());
                listOfCustomer.add(customer);
            }
            for (Customer customer : listOfCustomer) {
                System.out.println(customer);
            }
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    //write execl file
    public static void writeExcelFile() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Customer_Info");
        int rowNum = 0;
        //first row start at 1
        Row firstRow = sheet.createRow(rowNum++);
        //first cell start at o
        Cell firstCell = firstRow.createCell(0);

        firstCell.setCellValue("List of Customer");
        Cell secondCell = firstRow.createCell(1);
        secondCell.setCellValue("name");
        Cell thirdCell = firstRow.createCell(2);
        thirdCell.setCellValue("email");

        List<Customer> listOfCustomer = new ArrayList<Customer>();
        listOfCustomer.add(new Customer(1, "Sylvester Stallone", "abc@gmail.com"));
        listOfCustomer.add(new Customer(2, "Tom Cruise", "xyz@yahoo.com"));
        listOfCustomer.add(new Customer(3, "Vin Diesel", "abc@hotmail.com"));
        for (Customer customer : listOfCustomer) {
            Row row = sheet.createRow(rowNum++);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(customer.getId());
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(customer.getName());
            Cell cell3 = row.createCell(2);
            cell3.setCellValue(customer.getEmail());
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_PATH);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Done");
    }
}

