package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */
import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.support.TestRunListenerAdapter
import com.eviware.soapui.model.testsuite.TestCaseRunContext
import com.eviware.soapui.model.testsuite.TestCaseRunner
import com.eviware.soapui.plugins.ListenerConfiguration

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell


@ListenerConfiguration
public class TestRunReportListener extends TestRunListenerAdapter
{
    @Override
    void afterRun(TestCaseRunner testRunner, TestCaseRunContext runContext)
    {
        def currentUser = System.getProperty('user.name')
        def now = new Date().format('yyyy-MM-dd HH:mm:ss')
        def today = new Date().format('yyyyMMdd')
        def projectName = runContext.modelItem.project.name.replaceAll("[^a-zA-Z0-9.-]", "_")
        def userDir = System.getProperty('user.home')
        def SoapUIDir = new File(userDir,"\\Midun SoapUI Test Report\\")
        def fileName = "$projectName Test report - $today"+".xlsx"
        //Directory existence check
        if (SoapUIDir.exists())
        {
            file = new File(SoapUIDir,fileName)
            SoapUI.log "File exists"+file.absolutePath
        }
        else
        {
            SoapUIDir.mkdirs()
            file = new File(SoapUIDir,fileName)
            SoapUI.log "File doesn't exists; but created -> "+file.absolutePath
        }

        XSSFWorkbook workBookWrite = new XSSFWorkbook()
        XSSFSheet sheetWrite = workBookWrite.createSheet("SoapUITestReport")
        SoapUI.log "Sheet created : count -> "+workBookWrite.numberOfSheets

        //Style setting for 1-3 rows
        def style = workBookWrite.createCellStyle()
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index)
        style.setFillPattern(CellStyle.SOLID_FOREGROUND)

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
                //SoapUI.log "Created row -> $r column -> $c"
            }
            SoapUI.log "Row Number -> "+sheetWrite.lastRowNum

        }
        workBookWrite.write(new FileOutputStream(file))
        SoapUI.log "Done"

    }

    //    @Override
//    public void afterStep(TestCaseRunner testRunner, TestCaseRunContext runContext, TestStepResult result)
//    {
//        SoapUI.log.info "Test step type : "+ result.testStep.getClass()
////        Get test details from SoapUI
//        def tstName = result.testStep.name
//        def tcName = result.testStep.testCase.name
//        def tsName = result.testStep.testCase.testSuite.name
//        def projName = result.testStep.testCase.testSuite.project.name
//        def testID = tcName.split(/-/)[0] //Split just the TestCaseID from the test case name
//        def rawResponse, rawRequest, compiledRequestResponse, requestHeaders, responseHeaders, status, error, messages, timestamp
//
////        Get current date
//        def today= new Date().format("yyyy-MM-dd HH:mm:ss.S")
//
//        if(result.testStep.getClass().toString().equals("class com.eviware.soapui.impl.wsdl.teststeps.WsdlTestRequestStep") || result.testStep.getClass().toString().equals("class com.eviware.soapui.impl.wsdl.teststeps.RestTestRequestStep"))
//        {
//            //Get test response details for each test step result of SOAP and REST request
//            rawResponse = new String(result.getRawResponseData())
//            rawRequest = new String(result.getRawRequestData())
//            compiledRequestResponse = "#RAW REQUEST : " + System.getProperty("line.separator") + rawRequest+ System.getProperty("line.separator")+System.getProperty("line.separator") + "Raw Response :"+ System.getProperty("line.separator")+ rawResponse
//            status = result.status
//            error = result.error
//            messages = result.messages
//            timestamp = result.timeStamp
//        }
//        else
//        {
//            //Get test response details for each test step result of SOAP and REST request
//            rawResponse = "Not a REST or SOAP request"
//            rawRequest = "Not a REST or SOAP request"
//            compiledRequestResponse = "#RAW REQUEST : " + System.getProperty("line.separator") + rawRequest+ System.getProperty("line.separator")+System.getProperty("line.separator") + "Raw Response :"+ System.getProperty("line.separator")+ rawResponse
//            status = result.status
//            error = result.error
//            messages = result.messages
//            timestamp = result.timeStamp
//        }
//
////        Logs all the collected details
//        SoapUI.log.info "***Test step status => " + status
//        SoapUI.log.info "***Compiled Request & Response => " + compiledRequestResponse
//        SoapUI.log.info "***Raw Request => " + rawRequest
//        SoapUI.log.info "***Raw Response => " + rawResponse
//        SoapUI.log.info "***Response Headers => " + responseHeaders
//        SoapUI.log.info "***Errors => " + error + "Messages => " + messages + "Time Taken => " + timestamp
//        SoapUI.log.info "Today's date :" + today
//        SoapUI.log.warn "***KONDASAMY*** Plugin script ends for the step :: " + result.testStep.label
//
//        def compiledError = "Errors => " + error + System.getProperty("line.separator") +"Messages => " + messages
//
//    }
}
/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */