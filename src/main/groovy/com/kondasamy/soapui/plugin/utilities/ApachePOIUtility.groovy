package com.kondasamy.soapui.plugin.utilities

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell

def today = new Date().format('yyyyMMdd')
def projectName = "SamyTest"
def userDir = System.getProperty('user.home')
def SoapUIDir = new File(userDir,"\\SoapUI Test Report\\")
def fileName = "$projectName Test report - $today"+".xlsx"

//Directory existence check
if (SoapUIDir.exists())
{
    file = new File(SoapUIDir,fileName)
    println "File exists"
}
else
{
    SoapUIDir.mkdirs()
    file = new File(SoapUIDir,fileName)
    println "File doesn't exists; but created"
}

//File existence check
if(file.exists())
{
    FileOutputStream fileOut = new FileOutputStream(file)
    XSSFWorkbook workBook = new XSSFWorkbook(fileOut)
    println "Workbook created : $workBook"
    def sheetCount = workBook.numberOfSheets
    //Sheet existence check
    if(sheetCount < 1)
        XSSFSheet sheet = workBook.createSheet("SoapUITestReport")
    else
        XSSFSheet sheet = workBook.getSheet("SoapUITestReport")
    println "Got Sheet : "+sheet.lastRowNum
    //Data existence check
    if(sheet.lastRowNum < 1)
    {
        //iterating r number of rows
        for (int r=0;r < 5; r++ )
        {
            XSSFRow row = sheet.createRow(r)
            //iterating c number of columns
            for (int c=0;c < 5; c++ )
            {
                XSSFCell cell = row.createCell(c)
                cell.setCellValue("Cell "+r+" "+c)
                println "Created row -> $r column -> $c"
            }
        }
    }
    else
    {
        //iterating r number of rows
        for (int r=sheet.lastRowNum;r < 5; r++ )
        {
            XSSFRow row = sheet.createRow(r)
            //iterating c number of columns
            for (int c=0;c < 5; c++ )
            {
                XSSFCell cell = row.createCell(c)
                cell.setCellValue("Cell "+r+" "+c)
                println "Created row -> $r column -> $c"
            }
        }
    }
    workBook.write(fileOut)
    fileOut.flush()
    fileOut.close()
}


