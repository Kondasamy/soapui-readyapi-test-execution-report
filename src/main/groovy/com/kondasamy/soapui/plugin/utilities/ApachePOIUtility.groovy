package com.kondasamy.soapui.plugin.utilities

import org.apache.poi.ss.usermodel.CellStyle
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
style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index)
style.setFillPattern(CellStyle.SOLID_FOREGROUND)
style.setFont(workBookWrite.createFont().setBold(true))

//iterating r number of rows
for (int r=0;r < 5; r++ )
{
    XSSFRow row = sheetWrite.createRow(r)
    //iterating c number of columns
    for (int c=0;c < 5; c++ )
    {
        XSSFCell cell = row.createCell(c)
        cell.setCellValue("Cell "+r+" "+c)
        if (r<3)
            cell.setCellStyle(style)
        //println "Created row -> $r column -> $c"
    }
    println "Row Number -> "+sheetWrite.lastRowNum

}
workBookWrite.write(new FileOutputStream(file))
println("Done")