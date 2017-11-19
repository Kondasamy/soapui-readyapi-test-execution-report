package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */
import com.eviware.soapui.SoapUI;
import com.eviware.soapui.model.support.TestRunListenerAdapter
import com.eviware.soapui.model.testsuite.TestCaseRunContext
import com.eviware.soapui.model.testsuite.TestCaseRunner
import com.eviware.soapui.model.testsuite.TestStepResult
import com.eviware.soapui.plugins.ListenerConfiguration
import com.eviware.soapui.support.GroovyUtils
import com.kondasamy.soapui.plugin.settings.DatabasePrefs
import groovy.sql.Sql

@ListenerConfiguration
public class TestRunReportListener extends TestRunListenerAdapter
{
    @Override
    public void afterStep(TestCaseRunner testRunner, TestCaseRunContext runContext, TestStepResult result)
    {
        SoapUI.log.warn "###KONDASAMY***Plugin output*** \nRunning now : " + result.testStep.label
        SoapUI.log.info "Test step type : "+ result.testStep.getClass()
//        Get test details from SoapUI
        def tstName = result.testStep.name
        def tcName = result.testStep.testCase.name
        def tsName = result.testStep.testCase.testSuite.name
        def projName = result.testStep.testCase.testSuite.project.name
        def testID = tcName.split(/-/)[0] //Split just the TestCaseID from the test case name
        def rawResponse, rawRequest, compiledRequestResponse, requestHeaders, responseHeaders, status, error, messages, timestamp

//        Get current date
        def today= new Date().format("yyyy-MM-dd HH:mm:ss.S")

//        Get the database preferences details
        def DB_URL = DatabasePrefs.DEFAULT_URL
        def DB_NAME = DatabasePrefs.DB_NAME
        def DB_LOGIN = DatabasePrefs.LOGIN
        def DB_PWD = DatabasePrefs.PASSWORD

        if(result.testStep.getClass().toString().equals("class com.eviware.soapui.impl.wsdl.teststeps.WsdlTestRequestStep") || result.testStep.getClass().toString().equals("class com.eviware.soapui.impl.wsdl.teststeps.RestTestRequestStep"))
        {
            //Get test response details for each test step result of SOAP and REST request
            rawResponse = new String(result.getRawResponseData())
            rawRequest = new String(result.getRawRequestData())
            compiledRequestResponse = "#RAW REQUEST : " + System.getProperty("line.separator") + rawRequest+ System.getProperty("line.separator")+System.getProperty("line.separator") + "Raw Response :"+ System.getProperty("line.separator")+ rawResponse
            status = result.status
            error = result.error
            messages = result.messages
            timestamp = result.timeStamp
        }
        else
        {
            //Get test response details for each test step result of SOAP and REST request
            rawResponse = "Not a REST or SOAP request"
            rawRequest = "Not a REST or SOAP request"
            compiledRequestResponse = "#RAW REQUEST : " + System.getProperty("line.separator") + rawRequest+ System.getProperty("line.separator")+System.getProperty("line.separator") + "Raw Response :"+ System.getProperty("line.separator")+ rawResponse
            status = result.status
            error = result.error
            messages = result.messages
            timestamp = result.timeStamp
        }

//        Logs all the collected details
        SoapUI.log.info "***Test step status => " + status
        SoapUI.log.info "***Compiled Request & Response => " + compiledRequestResponse
        SoapUI.log.info "***Raw Request => " + rawRequest
        SoapUI.log.info "***Raw Response => " + rawResponse
        SoapUI.log.info "***Response Headers => " + responseHeaders
        SoapUI.log.info "***Errors => " + error + "Messages => " + messages + "Time Taken => " + timestamp
        SoapUI.log.info "Today's date :" + today
        SoapUI.log.warn "***KONDASAMY*** Plugin script ends for the step :: " + result.testStep.label
        SoapUI.log.fatal "*****Database Details : " + DB_NAME + DB_LOGIN + DB_PWD + DB_URL

        def compiledError = "Errors => " + error + System.getProperty("line.separator") +"Messages => " + messages

        /*
         * DB CONNECTION :: Connecting to the database
         */
        GroovyUtils.registerJdbcDriver( "com.microsoft.sqlserver.jdbc.SQLServerDriver" )
        def sql = Sql.newInstance("jdbc:sqlserver://$DB_URL;databaseName=$DB_NAME;user=$DB_LOGIN;password=$DB_PWD","com.microsoft.sqlserver.jdbc.SQLServerDriver")

        /*
         * EXECUTION :: Based on the Test Suite ans Test Case mappings take from the database, Test Runner execues the test cases
         */
        def insertQuery = "INSERT INTO dbo.TestResults (TestCaseRunID, TestCaseID, TestModuleID, Result, ExecutionDateTime, ScreenshotPath, ErrorLog) VALUES ('123', $testID, $tsName, $status, $timestamp, $compiledRequestResponse, $compiledError)"
        sql.executeInsert()

    }
}
/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */