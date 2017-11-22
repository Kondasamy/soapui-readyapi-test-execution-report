package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */
import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.support.ProjectRunListenerAdapter
import com.eviware.soapui.model.testsuite.ProjectRunContext
import com.eviware.soapui.model.testsuite.ProjectRunner
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
public class ProjectRunReportListener extends ProjectRunListenerAdapter
{
    @Override
    void afterRun(ProjectRunner projectRunner, ProjectRunContext runContext)
    {


        //Initiate project result collection
        projectRunner.results.each
                {
                    testSuiteResult ->
                    testSuiteResult.results.each
                            {
                                testCaseResult ->
                                    testCaseName = testCaseResult.testCase.name
                                    testRunStatus = testCaseResult.status
                                    testTimeStamp = testCaseResult.startTime
                                    /* Check if status falls under -
                                    INITIALIZED,
                                    RUNNING,
                                    CANCELED,
                                    FINISHED,
                                    FAILED,
                                    WARNING;*/
                                    def testCaseErrorMessage
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
                            }
                }
    }
}
/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */