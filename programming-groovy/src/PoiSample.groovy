/**
 * Created by bohnen on 2014/08/05.
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
import org.xml.sax.helpers.DefaultHandler

Workbook wb = new XSSFWorkbook()
wb.createSheet("testdesu")
wb.createSheet "secondsheet"
new File("workbook.xlsx").withOutputStream {
    wb.write(it)
}
