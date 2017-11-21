package com.kondasamy.soapui.plugin.utilities

import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*

import org.apache.poi.*

FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\kondasamy.j\\SoapUI Test Report\\SamyTest Test report.xlsx"))
//XSSFWorkbook workbook = new XSSFWorkbook()
XSSFWorkbook workbook = new XSSFWorkbook(inputStream)
println "Sheets count :: "+workbook.numberOfSheets
def sheet
if(workbook.numberOfSheets>0)
    sheet=workbook.getSheet(" Pagina_1")
else
    sheet=workbook.createSheet(" Pagina_1")
println "Last row count :: " + sheet.lastRowNum
for(int row=sheet.lastRowNum;row<5;row++){
    Row r = sheet.createRow(row)
    for (int cellnum=0; cellnum <5; cellnum++) {
        Cell c = r.createCell(cellnum);
        c.setCellValue( "Valor "+row+" "+cellnum );
    }
}
//inputStream.close()
FileOutputStream out = new FileOutputStream(new File("C:\\Users\\kondasamy.j\\SoapUI Test Report\\SamyTest Test report.xlsx"))
//write operation workbook using file out object
workbook.write(out)
out.close()
println("Successful")
