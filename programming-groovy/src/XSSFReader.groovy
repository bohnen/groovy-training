/**
 * Created by bohnen on 2014/08/08.
 */
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL')
        ]
)
import org.apache.poi.xssf.usermodel.*


FileInputStream inst = new FileInputStream("./data/sample.xlsx")
XSSFWorkbook wb = new XSSFWorkbook(inst)
println wb.getNumberOfSheets()
XSSFSheet sheet = wb.getSheetAt(wb.getNumberOfSheets() - 1)
sheet.autobreaks = true

sheet.each{ XSSFRow row ->
    println row.lastCellNum
}

println sheet.inject(0,{acc, row ->
  if(acc < row.lastCellNum)
      row.lastCellNum
  else
      acc
})

// 印刷設定
XSSFPrintSetup ps = sheet.getPrintSetup()
ps.paperSize = XSSFPrintSetup.A4_PAPERSIZE
ps.landscape = true
ps.fitWidth = (short)1
ps.fitHeight = (short)0

FileOutputStream oust = new FileOutputStream("out.xlsx")
wb.write(oust);
oust.close()
inst.close()
