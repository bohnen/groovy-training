package bohnen;

import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class Main {

    public static void main(String[] args) throws Exception{
        for(String s : args) {
            System.out.println(s);
        }

        InputStream in = new FileInputStream(args[0]);
        XSSFWorkbook ws = new XSSFWorkbook(in);
        // debug
        System.out.println(ws.getNumberOfSheets());

        XSSFSheet sheet = ws.getSheetAt(ws.getNumberOfSheets() -1);
        sheet.setAutobreaks(true);

        XSSFPrintSetup setup = sheet.getPrintSetup();
        setup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
        setup.setLandscape(true);
        setup.setFitWidth((short)1);
        setup.setFitHeight((short)0); // これが無いとエラーになる

        int lastRow = sheet.getPhysicalNumberOfRows();
        System.out.println(lastRow);
        System.out.println(sheet.getLastRowNum());

        int lastCol = 0;
        for(int i = 0 ; i <= sheet.getLastRowNum(); i++){
            XSSFRow row = sheet.getRow(i);
            if(row != null && row.getLastCellNum() > lastCol) {
                lastCol = row.getLastCellNum();
                System.out.println(row.getLastCellNum());
            }
        }

        System.out.println(lastCol);

        OutputStream os = new FileOutputStream("out.xlsx");
        ws.write(os);
        os.close();
        in.close();
    }
}
