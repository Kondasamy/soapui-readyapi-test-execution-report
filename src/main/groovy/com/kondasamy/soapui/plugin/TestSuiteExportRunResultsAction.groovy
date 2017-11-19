package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */
import com.eviware.soapui.SoapUI
import com.eviware.soapui.model.testsuite.TestSuite
import com.eviware.soapui.plugins.ActionConfiguration
import com.eviware.soapui.support.UISupport
import com.eviware.soapui.support.action.support.AbstractSoapUIAction

@ActionConfiguration(actionGroup = ActionGroups.TEST_SUITE_ACTIONS)
class TestSuiteExportRunResultsAction extends AbstractSoapUIAction<TestSuite>
{
    public TestSuiteExportRunResultsAction()
    {
        super("Plugin:Export Request and Response", "Saves recent responses of underlying TestSteps's to a file")
    }
    @Override
    void perform(TestSuite testSuite, Object o)
    {
        testSuite.getTestCaseList().each
                {
                    testCase->
                        testCase.getTestStepsOfType(com.eviware.soapui.impl.wsdl.teststeps.RestTestRequestStep.class).each
                                {
                                    tests->
                                        def tstName = tests.getName()
                                        def tcName = tests.testCase.getName()
                                        def projName = tests.testCase.testSuite.project.name
                                        def testID = tcName..split(/::/)[0] //Extracts the test ID from the test case name
                                        def response = tests.httpRequest.response
                                        def actualResponse = ""
                                        if( response == null )
                                        {
                                            actualResponse = "Missing Response for TestStep : " + tests.testStep.testCase.testSuite.name + "=>" + tests.testStep.testCase.name + "->" + tests.name
                                            SoapUI.log.warn "Missing Response for TestStep : " + tests.testStep.testCase.testSuite.name + "=>" + tests.testStep.testCase.name + "->" + tests.name
                                        }
                                        def data = response.getRawResponseData()
                                        if( data == null || data.length == 0 )
                                        {
                                            SoapUI.log.warn "Empty Response data for TestStep : "+ tests.testStep.testCase.testSuite.name + "=>"+tests.testStep.testCase.name+ "->"+tests.name
                                            actualResponse = "Empty Response data for TestStep : "+ tests.testStep.testCase.testSuite.name + "=>"+tests.testStep.testCase.name+ "->"+tests.name
                                        }
                                        else
                                        {
                                            def rawRequest = new String(response.getRawRequestData(),"UTF-8")
                                            def rawResponse = new String(response.getRawResponseData(),"UTF-8")

                                            def today= new Date().format("yyyy-MM-dd HH:mm:ss.S")
                                            actualResponse = "Raw Request:" + System.getProperty("line.separator") + rawRequest + System.getProperty("line.separator")+System.getProperty("line.separator") + "Raw Response :"+ System.getProperty("line.separator")+ rawResponse
                                            SoapUI.log.info "***Raw Request and Raw Response is exported to a file"
                                        }

                                }
//                        testCase.getTestStepsOfType(com.eviware.soapui.impl.wsdl.teststeps.WsdlTestRequestStep.class).each

                }
        UISupport.showInfoMessage("Artifacts Successfully exported!! Please see the SoapUI log for more information!","File Export Success!!!")
    }
}
/**
 * Created by Kondasamy J
 * Contact: Kondasamy@outlook.com
 */
