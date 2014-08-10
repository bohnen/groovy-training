/**
 * Created by bohnen on 2014/08/08.
 */
@Grapes(
        [
                @Grab(group='org.apache.poi', module='poi', version='3.10-FINAL'),
                @Grab(group='org.apache.poi', module='poi-ooxml', version='3.10-FINAL'),
        ]
)
import org.apache.poi.ss.usermodel.*

Sheet.metaClass.setCustomMargin = { ->
    delegate.with {
        setMargin(Sheet.LeftMargin,0.25)
        setMargin(Sheet.RightMargin,0.25)
        setMargin(Sheet.TopMargin,0.75)
        setMargin(Sheet.BottomMargin,0.75)
        setMargin(Sheet.FooterMargin,0.3)
        setMargin(Sheet.HeaderMargin,0.3)
    }
    delegate
}

Sheet.metaClass.customPageSetup = { ->
    delegate.with {
        autobreaks = true
        fitToPage = true
    }
    ps = delegate.getPrintSetup()
    ps.setPaperSize(PrintSetup.A4_PAPERSIZE)
    ps.setLandscape(true)
    ps.setFitWidth((short)1)
    ps.setFitHeight((short)0)
    delegate
}

Sheet.metaClass.customPrintAreaSetup = { ->
    wb = delegate.workbook
    nr = delegate.lastRowNum
    nc = delegate.inject(0, {acc, row ->
        if(acc < row.lastCellNum)
            row.lastCellNum
        else
            acc
    })
    wb.setPrintArea(wb.getSheetIndex(delegate),0,nc,0,nr)
    delegate
}

// main

wb = WorkbookFactory.create(new File("./data/sample.xlsx"))
for(i = 0; i < wb.numberOfSheets ; i++){
    sheet = wb.getSheetAt(i)
    sheet.setCustomMargin().customPageSetup().customPrintAreaSetup()
}

new File("out.xlsx").withOutputStream { out ->
    wb.write(out)
}