package com.kondasamy.soapui.plugin.utilities

//Dummy standalone test file created to test Apache POI functionality
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell

def today = new Date().format('yyyyMMdd')
def projectName = "SamyTest"
def userDir = System.getProperty('user.home')
def SoapUIDir = new File(userDir,"\\SoapUI Test Report\\")
def fileName = "$projectName Test report_$today"+".xlsx"

//Directory existence check
if (SoapUIDir.exists())
{
    file = new File(SoapUIDir,fileName)
    println "File exists"+file.absolutePath
}
else
{
    SoapUIDir.mkdirs()
    file = new File(SoapUIDir,fileName)
    println "File doesn't exists; but created -> "+file.absolutePath
}

XSSFWorkbook workBookWrite = new XSSFWorkbook()
XSSFSheet sheetWrite = workBookWrite.createSheet("SoapUITestReport")
println "Sheet created : count -> "+workBookWrite.numberOfSheets

//Style setting for 1-3 rows
def style = workBookWrite.createCellStyle()
style.setFillForegroundColor(IndexedColors.GOLD.index)
style.setFillPattern(CellStyle.SOLID_FOREGROUND)
style.setBorderBottom(CellStyle.BORDER_THIN)
style.setBottomBorderColor(IndexedColors.BLACK.index)
style.setBorderLeft(CellStyle.BORDER_THIN)
style.setLeftBorderColor(IndexedColors.GREEN.index)
style.setBorderRight(CellStyle.BORDER_THIN)
style.setRightBorderColor(IndexedColors.BLUE.index)
style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED)
style.setTopBorderColor(IndexedColors.BLACK.index)
Font font = workBookWrite.createFont()
font.setBold(true)
font.setColor(IndexedColors.WHITE.index)
font.setBoldweight(Font.BOLDWEIGHT_BOLD)
style.setFont(font)

//iterating r number of rows
for (int r=0;r < 5; r++ )
{
    XSSFRow row = sheetWrite.createRow(r)
    //iterating c number of columns
    for (int c=0;c < 5; c++ )
    {
/*        XSSFCell cell = row.createCell(c)
        cell.setCellValue("Cell "+r+" "+c)
        if (r<3)
            cell.setCellStyle(style)*/
        row.createCell(c).each
                {
                    cell->
                        cell.setCellValue("Cell "+r+" "+c)
                        cell.setCellStyle(style)
                }
        //println "Created row -> $r column -> $c"
    }
    println "Row Number -> "+sheetWrite.lastRowNum

}
workBookWrite.write(new FileOutputStream(file))
println("Done")