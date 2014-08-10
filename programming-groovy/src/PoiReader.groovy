
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
                @Grab(group='commons-io', module='commons-io', version='2.4')
        ]
)
import org.apache.poi.ss.usermodel.*
import org.apache.commons.io.FilenameUtils;

/**
 * Sheetクラスへの拡張
 */
@Category(Sheet)
class SheetCategory{
    /**
     * マージンを設定
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
     * 用紙サイズ、1ページフィットを設定
     * @return
     */
    Sheet applyCustomPageSetup(){
        this.with {
            autobreaks = true
            fitToPage = true
        }
        PrintSetup ps = this.getPrintSetup()
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE)
        ps.setLandscape(true)
        ps.setFitWidth((short)1)
        ps.setFitHeight((short)0)
        this
    }

    /**
     * 印刷エリアを設定。
     * @return
     */
    Sheet applyCustomPrintArea(){
        Workbook wb = this.workbook
        int nr = this.lastRowNum
        int nc = this.inject(0, {acc, row ->
            if(acc < row.lastCellNum)
                row.lastCellNum
            else
                acc
        })
        wb.setPrintArea(wb.getSheetIndex(this),0,nc,0,nr)
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
            sheet.applyCustomMargin().applyCustomPageSetup().applyCustomPrintArea()
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