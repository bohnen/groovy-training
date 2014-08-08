/**
 * Created by bohnen on 2014/08/08.
 */
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
                @Grab(group='dom4j', module='dom4j', version='1.6.1')
        ]
)
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*

XSSFWorkbook wb;

new File("sample.xlsx").withInputStream {
    wb = new XSSFWorkbook(it)
    println wb.getNumberOfSheets()
    sheet = wb.getSheetAt(wb.getNumberOfSheets() - 1)
//    sheet.setAutobreaks(true)

    // 印刷設定
    ps = sheet.getPrintSetup()
    ps.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE)
    ps.setLandscape(true)
//    ps.setFitWidth((short)1)

    new File("out.xlsx").withOutputStream { out ->
        wb.write(out)
    }
}