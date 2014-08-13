
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
                @Grab(group='commons-io', module='commons-io', version='2.4')
        ]
)
import org.apache.poi.ss.usermodel.*
import org.apache.commons.io.FilenameUtils

/**
 * Sheetクラスへの拡張
 */
@Category(Sheet)
class SheetCategory{
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
        ps.setLandscape(true)
        ps.setFitWidth((short)1)
        ps.setFitHeight((short)0)
        this
    }

    /**
     * 印刷エリアを設定。印刷エリアは何かしら入力のあるセルを含むように、最大の行 x 最大の列となるように設定
     * 行の最初のセルに"PB"という文字列があると、改ページを挿入する
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
        wb.setPrintArea(wb.getSheetIndex(this),0,nc,0,nr)
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

/**
 * 対象のブックに対して印刷設定を行う。実際の印刷設定はSheetCategoryクラスにて。
 * @param wb
 */
def printSetup(Workbook wb){
    for(i = 0; i < wb.numberOfSheets ; i++){
        sheet = wb.getSheetAt(i)
        use(SheetCategory) {
            sheet.applyCustomMargin()
                 .applyCustomPageSetup()
                 .applyCustomPrintArea()
                 .applyCustomHeaderFooter()
        }
    }
}

// main
// 全ての引数はxlsxファイルである前提. never check errors.

new File("out").mkdir();

args.each{ arg ->
    Workbook wb = WorkbookFactory.create(new File(arg))
    printSetup(wb)

    new File("./out/${FilenameUtils.getName(arg)}").withOutputStream { out ->
        wb.write(out)
    }
}