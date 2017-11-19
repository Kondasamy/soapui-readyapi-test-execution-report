package com.kondasamy.soapui.plugin

/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */

import com.eviware.soapui.SoapUI
import com.eviware.soapui.impl.wsdl.WsdlTestSuite
import com.eviware.soapui.model.project.Project
import com.eviware.soapui.impl.wsdl.WsdlProject
import com.eviware.soapui.impl.wsdl.testcase.WsdlTestCase
import com.eviware.soapui.plugins.ActionConfiguration
import com.eviware.soapui.support.UISupport
import com.eviware.soapui.support.action.support.AbstractSoapUIAction
import com.eviware.soapui.support.GroovyUtils
import com.kondasamy.soapui.plugin.settings.DatabasePrefs
import groovy.sql.Sql

@ActionConfiguration(actionGroup = ActionGroups.OPEN_PROJECT_ACTIONS)
class ProjectCreateTestRunAction extends AbstractSoapUIAction <WsdlProject>
{
    public ProjectCreateTestRunAction()
    {
        super("Plugin:Create Test Manager", "Creates Test Manager test suite to manage automation runs")
    }

    @Override
    void perform(WsdlProject project, Object o)
    {
        def projectName = project.name

        //Set project properties for initial run
        project.setPropertyValue("Endpoint","") //Endpoint Project Property
        project.setPropertyValue("UnicornProjectID","") //Unicorn project id holding property

        //Create test manager
        if (project.getTestSuiteByName("Test Manager") != null)
        {
            if (project.getTestSuiteByName("Test Manager").getTestCaseByName("Test Driver") != null)
            {
                SoapUI.log.error "The current project is already equipped with the \"Test Manager\" test suite and the \"Test Driver\" test case!"
                UISupport.showErrorMessage("The current project is already equipped with the \"Test Manager\" test suite and the \"Test Driver\" test case!")
            }
            else
            {
                def testSuite = project.getTestSuiteByName("Test Manager")
                createTestDriverTC(testSuite)
            }
            if (project.getTestSuiteByName("Test Manager").getTestCaseByName("Setup Scripts") != null)
            {
                SoapUI.log.error "The current project is already equipped with the \"Test Manager\" test suite and the \"Setup Scripts\" test case!"
                UISupport.showErrorMessage("The current project is already equipped with the \"Test Manager\" test suite and the \"Setup Scripts\" test case!")
            }
            else
            {
                def testSuite = project.getTestSuiteByName("Test Manager")
                createSetupScriptTC(testSuite)
            }
         }
        else
        {
            createTestManagerTS(project)
        }
    }

    void createTestManagerTS(Project project)
    {
        SoapUI.log.info "The test suite \"Test Manager\" is not available. Created by plugin..!"
        WsdlTestSuite testSuite = project.addNewTestSuite("Test Manager")
        createSetupScriptTC(testSuite)
        createTestDriverTC(testSuite)
        //04. Create Project in Database if not created. If already available fetch the project ID using the project name.
        //Get Project ID to be run from the project level property
        def projectName = project.name
        def DB_URL = SoapUI.settings.getString(DatabasePrefs.DEFAULT_URL,"Not_found")
        def DB_LOGIN = SoapUI.settings.getString(DatabasePrefs.LOGIN,"Not_found")
        def DB_PWD = SoapUI.settings.getString(DatabasePrefs.PASSWORD,"Not_found")
        def DB_NAME = SoapUI.settings.getString(DatabasePrefs.DB_NAME,"Not_found")
        GroovyUtils.registerJdbcDriver( "com.microsoft.sqlserver.jdbc.SQLServerDriver" )
        def sql = Sql.newInstance("jdbc:sqlserver://$DB_URL;databaseName=$DB_NAME;user=$DB_LOGIN;password=$DB_PWD","com.microsoft.sqlserver.jdbc.SQLServerDriver")
        def projectID = sql.firstRow("select ID from dbo.M_Projects where ProjectName=$projectName")
        if (projectID == null)
        {
            sql.executeInsert("INSERT INTO dbo.M_Projects (ProjectName,Jurisdiction) VALUES ($projectName,'ProjectJurisdiction')")
            projectID = sql.firstRow("select ID from dbo.M_Projects where ProjectName=$projectName")
            project.setPropertyValue("UnicornProjectID",projectID[0].toString())
        }
        else
        {
            project.setPropertyValue("UnicornProjectID",projectID[0].toString())
        }
        SoapUI.log.info "%%%%%Also set the ProjectID in database!!! %%%%%%"
        UISupport.showInfoMessage("Successfully created the \"Test Manager\" test suite!", "Test Manager Creation Success!!!")
    }

    /*
    Method to create setup script tc that contains features to set endpoints and reset properties
     */
    void createTestDriverTC(WsdlTestSuite testSuite)
    {
        SoapUI.log.info "The test case \"Test Driver\" is not available. Created by plugin..!"
        WsdlTestCase testCase = testSuite.addNewTestCase("Test Driver")
        try
        {
            testCase.addTestStep("groovy", "Plugin_Test_Runner").setScript(fileToString("TestRunner.txt"))
        }
        catch (Exception e)
        {
            SoapUI.log.error e.toString()
        }
        SoapUI.log.info "Successfully created the \"Plugin_Test_Runner\" groovy test step!"
     }

    /*
    Method to create test manager
    */
    void createSetupScriptTC(WsdlTestSuite testSuite)
    {
        SoapUI.log.info "The test case \"Test Driver\" is not available. Created by plugin..!"
        WsdlTestCase testCase = testSuite.addNewTestCase("Setup Scripts")
        try
        {
            testCase.addTestStep("groovy", "01. Reset Properties").setScript(fileToString("ResetProperties.txt"))
            testCase.addTestStep("groovy", "02. Set Endpoint from Project Property").setScript(fileToString("SetEndpoints.txt"))
        }
        catch (Exception e)
        {
            SoapUI.log.error e.toString()
        }
        SoapUI.log.info "Successfully created the \"Set Endpoint from Project Property\" groovy test step!"

    }

    /*
    Method to read files and convert to String
     */
    String fileToString(String fileName)
    {
        InputStreamReader inputStreamReader = new InputStreamReader(this.class.getResourceAsStream("/$fileName"))
        BufferedReader bufferedReader = new BufferedReader(inputStreamReader)
        String text = "", line = null
        while((line=bufferedReader.readLine())!=null)
        {
            text = text + "\n" + line
        }
        return text
    }
}
/**
 * Created by Kondasamy Jayaraman
 * Contact: Kondasamy@outlook.com
 */