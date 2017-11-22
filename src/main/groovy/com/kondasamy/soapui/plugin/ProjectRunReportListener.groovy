package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */
import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.support.ProjectRunListenerAdapter
import com.eviware.soapui.model.testsuite.ProjectRunContext
import com.eviware.soapui.model.testsuite.ProjectRunner
import com.eviware.soapui.plugins.ListenerConfiguration

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow

@ListenerConfiguration
public class ProjectRunReportListener extends ProjectRunListenerAdapter
{
    @Override
    void afterRun(ProjectRunner projectRunner, ProjectRunContext runContext)
    {
        try
        {
            //def currentUser = System.getProperty('user.name')
            //def now = new Date().format('yyyy-MM-dd HH:mm:ss')
            def today = new Date().format('yyyyMMdd')
            def projectName = runContext.modelItem.project.name.replaceAll("[^a-zA-Z0-9.-]", "_")
            def userDir = System.getProperty('user.home')
            def SoapUIDir = new File(userDir,"\\Midun SoapUI Test Report\\")
            def fileName = "$projectName - Test execution report - $today"+".xlsx"
            def file
            //Directory existence check
            if (SoapUIDir.exists())
            {
                file = new File(SoapUIDir,fileName)
                SoapUI.log "RESULT EXPORTER :: File exists -> "+file.absolutePath
            }
            else
            {
                SoapUIDir.mkdirs()
                file = new File(SoapUIDir,fileName)
                SoapUI.log "RESULT EXPORTER :: File doesn't exists; but created -> "+file.absolutePath
            }

            //Initiate XSSF Workbook
            XSSFWorkbook workBookWrite = new XSSFWorkbook()
            XSSFSheet sheetWrite = workBookWrite.createSheet("SoapUITestReport")
            SoapUI.log "RESULT EXPORTER :: Sheet created :: name -> "+workBookWrite.getSheetAt(0).sheetName

            //Header style
            def headerStyle = workBookWrite.createCellStyle()
            headerStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            headerStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            headerStyle.setBorderBottom(CellStyle.BORDER_DASH_DOT_DOT)
            headerStyle.setBottomBorderColor(IndexedColors.BLACK.index)
            headerStyle.setBorderLeft(CellStyle.BORDER_THIN)
            headerStyle.setLeftBorderColor(IndexedColors.GREEN.index)
            headerStyle.setBorderRight(CellStyle.BORDER_THIN)
            headerStyle.setRightBorderColor(IndexedColors.BLUE.index)
            headerStyle.setBorderTop(CellStyle.BORDER_THIN)
            headerStyle.setTopBorderColor(IndexedColors.BLACK.index)
            Font passStyleFont = workBookWrite.createFont()
            passStyleFont.setBold(true)
            passStyleFont.setColor(IndexedColors.WHITE.index)
            passStyleFont.setBoldweight(Font.BOLDWEIGHT_BOLD)
            headerStyle.setFont(passStyleFont)

            //Pass status style
            def passStyle = workBookWrite.createCellStyle()
            passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index)
            passStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            passStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            passStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            passStyle.setBorderBottom(CellStyle.BORDER_THIN)
            passStyle.setBottomBorderColor(IndexedColors.BLACK.index)
            passStyle.setBorderLeft(CellStyle.BORDER_THIN)
            passStyle.setLeftBorderColor(IndexedColors.GREEN.index)
            passStyle.setBorderRight(CellStyle.BORDER_THIN)
            passStyle.setRightBorderColor(IndexedColors.BLUE.index)
            passStyle.setBorderTop(CellStyle.BORDER_THIN)
            passStyle.setTopBorderColor(IndexedColors.BLACK.index)

            //Fail status style
            def failStyle = workBookWrite.createCellStyle()
            failStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index)
            failStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            failStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            failStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            failStyle.setBorderBottom(CellStyle.BORDER_THIN)
            failStyle.setBottomBorderColor(IndexedColors.BLACK.index)
            failStyle.setBorderLeft(CellStyle.BORDER_THIN)
            failStyle.setLeftBorderColor(IndexedColors.GREEN.index)
            failStyle.setBorderRight(CellStyle.BORDER_THIN)
            failStyle.setRightBorderColor(IndexedColors.BLUE.index)
            failStyle.setBorderTop(CellStyle.BORDER_THIN)
            failStyle.setTopBorderColor(IndexedColors.BLACK.index)

            //Default style
            def defaultStyle = workBookWrite.createCellStyle()
            defaultStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            defaultStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            defaultStyle.setFillForegroundColor(IndexedColors.GOLD.index)
            defaultStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
            defaultStyle.setBorderBottom(CellStyle.BORDER_THIN)
            defaultStyle.setBottomBorderColor(IndexedColors.BLACK.index)
            defaultStyle.setBorderLeft(CellStyle.BORDER_THIN)
            defaultStyle.setLeftBorderColor(IndexedColors.GREEN.index)
            defaultStyle.setBorderRight(CellStyle.BORDER_THIN)
            defaultStyle.setRightBorderColor(IndexedColors.BLUE.index)
            defaultStyle.setBorderTop(CellStyle.BORDER_THIN)
            defaultStyle.setTopBorderColor(IndexedColors.BLACK.index)

            //Initialize SoapUI data sets
            def testCaseName, testCaseErrorMessage, testRunStatus, testTimeStamp
            def row = 0, col = 0

            //Initialize the first row of the file with header
            XSSFRow rowData = sheetWrite.createRow(row)
            rowData.createCell(0).setCellValue("TESTCASE NAME")
            rowData.createCell(1).setCellValue("TESTCASE STATUS")
            rowData.createCell(3).setCellValue("TIMESTAMP")
            rowData.createCell(4).setCellValue("REMARKS")
            rowData.setRowStyle(headerStyle)
            row++

            //Initiate project result collection
            projectRunner.results.each
                    {
                        testSuiteResult ->
                            testSuiteResult.results.each
                                    {
                                        testCaseResult ->
                                            //Collect SoapUI data
                                            testCaseName = testCaseResult.testCase.name
                                            testRunStatus = testCaseResult.status.toString()
                                            testTimeStamp = testCaseResult.startTime
                                            /* Check if status falls under - INITIALIZED, RUNNING, CANCELED, FINISHED, FAILED, WARNING;*/
                                            if(testCaseResult.status.toString() == "FAILED")
                                            {
                                                testCaseResult.results.each
                                                        {
                                                            testStepResult ->
                                                                testCaseErrorMessage += testStepResult.messages
                                                        }
                                            }
                                            else
                                                testCaseErrorMessage = "OK"
                                            //Initialize testcase data
                                            rowData = sheetWrite.createRow(row)
                                            rowData.createCell(0).setCellValue(testCaseName)
                                            rowData.createCell(1).setCellValue(testRunStatus)
                                            rowData.createCell(3).setCellValue(testTimeStamp)
                                            rowData.createCell(4).setCellValue(testCaseErrorMessage.toString())
                                            if(testRunStatus == "FAILED")
                                                rowData.setRowStyle(failStyle)
                                            else if (testRunStatus == "FINISHED")
                                                rowData.setRowStyle(passStyle)
                                            else
                                                rowData.setRowStyle(defaultStyle)
                                            row++
                                    }
                    }
            workBookWrite.write(new FileOutputStream(file))
            SoapUI.log "RESULT EXPORTER :: Success. Results exported to -> "+file.absolutePath
        }
        catch(Exception exception)
        {
            SoapUI.log.error "RESULT EXPORTER :: Exception occurred -> $exception"
        }
    }
}
/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */