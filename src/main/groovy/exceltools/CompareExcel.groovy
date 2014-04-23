package exceltools

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.util.CellReference

class CompareExcel {

    HSSFWorkbook getWorkbook(fileName) {
        HSSFWorkbook workbook
        new File(fileName).withInputStream { is ->
            workbook = new HSSFWorkbook(is)
        }
        workbook
    }

    public void compare(String firstFile, String secondFile) {

        HSSFWorkbook wb1 = getWorkbook(firstFile)
        HSSFWorkbook wb2 = getWorkbook(secondFile)


        int sheetCnt1 = wb1.getNumberOfSheets();
        int sheetCnt2 = wb2.getNumberOfSheets();

        if (sheetCnt1 != sheetCnt2) {
            println "Number of sheets doesn't match :" + sheetCnt1 + " vs " + sheetCnt2
        }

        for(int i = 0; i<sheetCnt1 ; i++) {
            HSSFSheet sheet1 = wb1.getSheetAt(i)
            String sheetName = sheet1.sheetName
            HSSFSheet sheet2 = wb2.getSheet(sheetName)

            if (sheet1 == null) {
                println "No sheet found with name '" + sheetName + ' in ' +secondFile
            } else {
                println "Sheet:" + sheetName
                compareSheets(sheet1, sheet2, firstFile, secondFile)
            }
        }
    }

    private void compareSheets(HSSFSheet sheet1, HSSFSheet sheet2, String firstFile, String secondFile) {
        Iterator rows1 = sheet1.rowIterator()
        Iterator rows2 = sheet2.rowIterator()

        println  "Row count : " + sheet1.getPhysicalNumberOfRows() +" vs "+   sheet2.getPhysicalNumberOfRows()

        while (rows1.hasNext()) {
            if (rows2.hasNext()) {
                HSSFRow row1 = rows1.next()
                HSSFRow row2 = rows2.next()
                compareRows(row1, row2, firstFile, secondFile)
            } else {
                println "Number of rows doesn't match."
            }

        }
    }

    def compareRows(HSSFRow row1, HSSFRow row2, String firstFile, String secondFile) {
        def numberOfCells = row1.getPhysicalNumberOfCells()
        if (numberOfCells != row2.getPhysicalNumberOfCells()) {
            println("Row " + row1.getRowNum() + " doesn't match")
        } else {
            for(cellIndex in 0 .. numberOfCells-1) {
                HSSFCell c1 = row1.getCell(cellIndex)
                String cell1 = row1.getCell(cellIndex).getNumericCellValue()
                String cell2 = row2.getCell(cellIndex).getNumericCellValue()

                if (cell1 != cell2){
                   CellReference cf = new CellReference(c1.getRowIndex(), cellIndex)
                   println cf.formatAsString() + " : " + cell1 +" vs " +  cell2
                }

            }

        }

    }

    public double getColValue(HSSFRow row) {

        Iterator cells = row.cellIterator()
        HSSFCell cell = (HSSFCell) cells.next()
        return cell.getNumericCellValue()
    }
}
