/**
 * Download smartsheet project file as xls and format to print.
 */
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
                @Grab(group='commons-io', module='commons-io', version='2.4'),
                @Grab(group='org.codehaus.groovy.modules.http-builder', module='http-builder', version='0.7.1' )

        ]
)
import org.apache.poi.ss.usermodel.*
import groovyx.net.http.HTTPBuilder
import groovyx.net.http.Method
import org.apache.poi.hssf.usermodel.HSSFBorderFormatting
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.util.HSSFColor

/**
 * Sheetクラスへの拡張
 */
@Category(Sheet)
class WBSSheetCategory{
    /**
     * マージンを設定。左右 約1cm, 上下約1.5cm
     * @return
     */
    Sheet applyCustomMargin(){
        this.with {
            setMargin(Sheet.LeftMargin,0.4)
            setMargin(Sheet.RightMargin,0.4)
            setMargin(Sheet.TopMargin,0.6)
            setMargin(Sheet.BottomMargin,0.6)
            setMargin(Sheet.FooterMargin,0.3)
            setMargin(Sheet.HeaderMargin,0.3)
        }
        this
    }

    /**
     * 用紙サイズ、A4 1ページフィットを設定
     * @return
     */
    Sheet applyCustomPageSetup(){
        this.with {
            autobreaks = true
            fitToPage = true
        }
        org.apache.poi.ss.usermodel.PrintSetup ps = this.getPrintSetup()
        ps.setPaperSize(org.apache.poi.ss.usermodel.PrintSetup.A4_PAPERSIZE)
        ps.setLandscape(false)
        ps.setFitWidth((short)1)
        ps.setFitHeight((short)0)
        this
    }

    /**
     * 印刷エリアと罫線、ヘッダ、フォントの設定
     *
     * @return
     */
    Sheet applyCustomPrintArea(){
        Workbook wb = this.workbook
        int nr = this.lastRowNum
        int nc = 0
        this.each { row ->
            Cell c = row.getCell(0)
            if(c != null && c.getCellType() == Cell.CELL_TYPE_STRING && c.getStringCellValue() == "PB") {
                this.setRowBreak(row.getRowNum())
                c.setCellValue("")
            }
            if(nc < row.lastCellNum)
                nc = row.lastCellNum
        }
        wb.setPrintArea(wb.getSheetIndex(this),0,nc -1,0,nr) // なんでか知らんが列が1ずれる
        println "lastRow: ${nr}, lastCell: ${nc}"

        // ヘッダ、フォント、罫線の設定
        Font f = wb.createFont()
        f.setFontName("メイリオ")

        // 全体設定
        for(int i=0;i<=nr;i++){
            Row r = this.getRow(i)
            for(int k=0;k<nc;k++){
                Cell c = r.getCell(k)
                CellStyle s = c.getCellStyle()
                s.setFont(f)
                s.setBorderTop(HSSFBorderFormatting.BORDER_THIN)
                s.setBorderBottom(HSSFBorderFormatting.BORDER_THIN)
                s.setBorderLeft(HSSFBorderFormatting.BORDER_THIN)
                s.setBorderRight(HSSFBorderFormatting.BORDER_THIN)
                s.setWrapText(true);
                c.setCellStyle(s);
            }
            r.setHeight((short)(20*20))
        }

        // ヘッダ
        Row hRow = this.getRow(0)
        f = wb.createFont()
        f.setFontName("メイリオ")
        f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD)
        for(int i=0;i<nc; i++){
            Cell c = hRow.getCell(i)
            CellStyle s = c.getCellStyle()
            s.setFillForegroundColor(HSSFColor.GREEN.index)
            s.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND)
            s.setFont(f)
            c.setCellStyle(s)
        }

        this
    }
    /**
     * ヘッダフッタを設定。
     * @return
     */
    Sheet applyCustomHeaderFooter(){
        Header header = this.getHeader()
        header.setLeft("&\"メイリオ,標準\"ファイル名：&F")
        header.setRight("&\"メイリオ,標準\"シート名：&A")
        Footer footer = this.getFooter()
        footer.setRight("&P / &N")
        this
    }
}

def formatSheet(Workbook wb){
    // Discussion シートを削除
    idx = wb.getSheetIndex("Discussions")
    wb.removeSheetAt(idx)

    // 印刷設定
    sheet = wb.getSheetAt(0)

    // 不必要な列削除
    sheet.setColumnHidden(8,true)
    sheet.setColumnHidden(9,true)

    use(WBSSheetCategory){
        sheet.applyCustomMargin()
                .applyCustomPageSetup()
                .applyCustomPrintArea()
                .applyCustomHeaderFooter()
    }
}


// input properties
def prop = new ConfigSlurper().parse(new File(".smartsheet.groovy").toURI().toURL())

// output directory and file
new File("out").mkdir();
def file = new File("./out/out.xls")

def http = new HTTPBuilder('https://api.smartsheet.com')
http.request(Method.GET){ req ->
    uri.path = prop.path
    headers.Authorization = "Bearer ${prop.apikey}"
    headers.Accept = 'application/vnd.ms-excel'

    response.success = {resp, reader ->
        Workbook wb = WorkbookFactory.create(reader)
        formatSheet(wb)

        file.withOutputStream {out ->
            wb.write(out)
        }
    }
}
