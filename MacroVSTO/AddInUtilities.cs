using Microsoft.Office.Interop.MSProject;
using System.Runtime.InteropServices;
using MSProject = Microsoft.Office.Interop.MSProject;
using System.Collections.Generic;
using Microsoft.AspNetCore.WebUtilities;
using System.Net.Http;
using System;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;

namespace MacroVSTO
{
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        void ReimportTestExecution(String key);
        void ImportTestExecution(String key);
        void ImportTestPlan(String key);
        void ImportAllTestExecutions();
        void WriteToTextToPlain(string text10In, string text11In, string text12In, string text13In, string text14In, string text15In);
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        
        // this is a dictionary containing the mapping from the text field of a task to the desired custom field id
        Dictionary<string, string> textToField = new Dictionary<string, string>();

        string token = "Bearer OTY3MDY3MzE3OTQ1OghfelOUY/pXavz/Mfos0VSIZe5f";

        HttpClient client = new HttpClient()
        {
            BaseAddress = new Uri("http://localhost:8080/rest/"),
        };

        Dictionary<string, string> parameters = new Dictionary<string, string>()
        {
            ["testExecKey"] = "",
            ["includeTestFields"] = "customfield_10128,customfield_10129,summary"
        };

        MSProject.Application application = Globals.ThisAddIn.Application;

        // this method goes through every task in the project and updates the tests directory accordingly
        public Dictionary<string, Dictionary<string, int>> scanProject()
        {
            
            Dictionary<string, Dictionary<string, int>> content = new Dictionary<string, Dictionary<string, int>>();
            foreach (MSProject.Task task in application.ActiveProject.Tasks)
            {
                if (!content.ContainsKey(task.Text20))
                {
                    content[task.Text20] = new Dictionary<string, int>();
                    content[task.Text20][task.Text19] = 1;
                    continue;
                }
                if (!content[task.Text20].ContainsKey(task.Text19))
                {
                    content[task.Text20][task.Text19] = 1;
                }
                else
                {
                    content[task.Text20][task.Text19] = content[task.Text20][task.Text19] + 1;
                }
            }
            return content;
        }

        // in order to get the mapping from text field to custom field, all possible custom fields and their corresponding id are stored int plaintofield.json
        // using the texttoplain.json file, the mapping from text field to custom field id can be created
        public void WriteToTextToPlain(string text10In, string text11In, string text12In, string text13In, string text14In, string text15In)
        {
            string text10 = text10In.ToLower().Trim().Replace(" ", "");
            string text11 = text11In.ToLower().Trim().Replace(" ", "");
            string text12 = text11In.ToLower().Trim().Replace(" ", "");
            string text13 = text11In.ToLower().Trim().Replace(" ", "");
            string text14 = text11In.ToLower().Trim().Replace(" ", "");
            string text15 = text11In.ToLower().Trim().Replace(" ", "");
            Dictionary<string, string> textToPlain = new Dictionary<string, string>();
            textToPlain.Add("text10", text10);
            textToPlain.Add("text11", text11);
            textToPlain.Add("text12", text12);
            textToPlain.Add("text13", text13);
            textToPlain.Add("text14", text14);
            textToPlain.Add("text15", text15);

            string jsonString = JsonConvert.SerializeObject(textToPlain);
            File.WriteAllText("..\\texttoplain.json", jsonString);
        }

        // this method creates the actual mapping from text fields to custom field ids
        // it also adjusts the parameters of the http request, to get the necessary custom fields
        // as such, this method has to be called before each http request
        public void updateTextToField()
        {
            string plainToFieldString = File.ReadAllText("..\\plaintofield.json");
            string textToPlainString = File.ReadAllText("..\\texttoplain.json");
            var plainToField = JsonConvert.DeserializeObject<Dictionary<string, string>>(plainToFieldString);
            var textToPlain = JsonConvert.DeserializeObject<Dictionary<string, string>>(textToPlainString);

            parameters["includeTestFields"] = "customfield_10128,customfield_10129,summary";
            textToField = new Dictionary<string, string>();

            foreach (var text in textToPlain)
            {
                if (text.Value != "")
                {
                    textToField.Add(text.Key, plainToField[text.Value]);
                    parameters["includeTestFields"] += "," + plainToField[text.Value];
                }
            }
        }

        // imports a single test execution
        public async void ImportTestExecution(String key)
        {
            //try
            //{
                key = key.ToUpper();
                updateTextToField();
                parameters["testExecKey"] = key;
                string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);
                client.DefaultRequestHeaders.Add("Authorization", token);
                var task = client.GetStringAsync(uri);
                String jsonString = task.GetAwaiter().GetResult();
                client.DefaultRequestHeaders.Remove("Authorization");
                var jsonArray = JArray.Parse(jsonString);

                foreach (var test in jsonArray)
                {
                    AddTest(test, key); // this method call adds the actual test to the project, all the other stuff can probably be refactored away (and should be)
                }
        //}
        //    catch
        //    {
        //        MessageBox.Show("There's connection issues with Jira or the specified TestExecution does not exist");
        //        return;
        //    }

}

        // this method fetches data on a test execution and compares it to the existing representation of that test execution in the project
        // if they don't match, the extra tests are added to the project
        // however, tests cannot be removed from this
        public void ReimportTestExecution(String key)
        {
            application.SelectRow(1, false);
            if (key == null)
            {
                return;
            }
            key = key.ToUpper();
            Dictionary<string, Dictionary<string, int>> tests = scanProject();
            updateTextToField();
            if (!tests.ContainsKey(key))
            {
                MessageBox.Show("Test Execution has not been imported\nThus it cannot be reimported");
                return;
            }

            Dictionary<string, int> testExecCopy = tests[key];

            parameters["testExecKey"] = key;
            string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);
            client.DefaultRequestHeaders.Add("Authorization", token);
            var task = client.GetStringAsync(uri);
            String jsonString = task.GetAwaiter().GetResult();
            client.DefaultRequestHeaders.Remove("Authorization");
            var jsonArray = JArray.Parse(jsonString);
            foreach (var test in jsonArray)
            {
                if (!testExecCopy.ContainsKey(test["testKey"].ToString()))
                {
                    testExecCopy[test["testKey"].ToString()] = 0;
                    AddTest(test, key);
                } else if (testExecCopy[test["testKey"].ToString()] == 0)
                {
                    AddTest(test, key);
                } else
                {
                    testExecCopy[test["testKey"].ToString()] = testExecCopy[test["testKey"].ToString()] - 1;
                }

            }
        }

        // this method adds the actual test to the project
        // here, the data on the test from xray is transformed into the contents of an MS Project Task
        public void AddTest(JToken test, String testExecKey)
        {
            MSProject.Task newTask = application.ActiveProject.Tasks.Add(test["testIssueFields"]["summary"]);

            // this sets the basic fields, that every test hast
            newTask.Start = test["testIssueFields"]["customfield_10128"].ToString();
            newTask.Finish = test["testIssueFields"]["customfield_10129"].ToString();
            newTask.Text19 = test["testKey"].ToString();
            newTask.Text20 = testExecKey;

            // this sets the custom fields, that were configured and translated to the texttofield dictionary
            newTask.Text10 = test["testIssueFields"][textToField["text10"]].ToString();
            newTask.Text11 = test["testIssueFields"][textToField["text11"]].ToString();
            newTask.Text12 = test["testIssueFields"][textToField["text12"]].ToString();
            newTask.Text13 = test["testIssueFields"][textToField["text13"]].ToString();
            newTask.Text14 = test["testIssueFields"][textToField["text14"]].ToString();
            newTask.Text15 = test["testIssueFields"][textToField["text15"]].ToString();

            if (test["assignee"] == null)
                {
                    newTask.ResourceNames = "Herbert";
                }
                else
                {
                    newTask.ResourceNames = test["assignee"].ToString();
                }

                application.SelectRow(newTask.ID, false);
                switch (test["status"].ToString())
                {
                    case "PASS":
                        application.GanttBarFormat(StartColor: PjColor.pjLime, MiddleColor: PjColor.pjLime, EndColor: PjColor.pjLime);
                        break;
                    case "EXECUTING":
                        application.GanttBarFormat(StartColor: PjColor.pjYellow, MiddleColor: PjColor.pjYellow, EndColor: PjColor.pjYellow);
                        break;
                    case "FAIL":
                        application.GanttBarFormat(StartColor: PjColor.pjRed, MiddleColor: PjColor.pjRed, EndColor: PjColor.pjRed);
                        break;
                }

        }

        // this method fetches the key of all test executions and imports all of them
        public async void ImportAllTestExecutions()
        {
            try
            {
                client.DefaultRequestHeaders.Add("Authorization", token);
                var task = client.GetStringAsync("api/2/search?jql=issueType='Test Execution'");
                String testExecutionsString = task.GetAwaiter().GetResult();
                client.DefaultRequestHeaders.Remove("Authorization");
                var testExecutionsRaw = JObject.Parse(testExecutionsString);
                var testExecutions = testExecutionsRaw["issues"];
                foreach (var testExecution in testExecutions)
                {
                    lock (testExecutionsString)
                    {
                        ImportTestExecution(testExecution["key"].ToString());
                    }
                }
            } 
            catch
            {
                MessageBox.Show("An exception occured. This might mean that Jira is unavailable");
                return;
            }
        }

        // this method fetches all tests related to a test plan and then adds those tests to the project
        public async void ImportTestPlan(string key)
        {
            key = key.ToUpper();
            updateTextToField();
            parameters.Remove("testExecKey");
            parameters.Add("testPlanKey", key);
            client.DefaultRequestHeaders.Add("Authorization", token);
            string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);

            var task = client.GetStringAsync(uri);
            String testPlanString = task.GetAwaiter().GetResult();

            client.DefaultRequestHeaders.Remove("Authorization");
            parameters.Remove("testPlanKey");
            parameters.Add("testExecKey", "");
            var testPlan = JArray.Parse(testPlanString);

            foreach (var test in testPlan)
            {
                AddTest(test, key);
            }
        }
    }
}

                //key = key.ToUpper();
                //updateTextToField();
                //parameters["testExecKey"] = key;
                //string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);
                //client.DefaultRequestHeaders.Add("Authorization", token);
                //var task = client.GetStringAsync(uri);
                //String jsonString = task.GetAwaiter().GetResult();
                //client.DefaultRequestHeaders.Remove("Authorization");
                //var jsonArray = JArray.Parse(jsonString);

//Sub ReimportTestExecution()
//    Dim addIn As COMAddIn
//    Dim automationObject As Object
//    Set addIn = Application.COMAddIns("MacroVSTO")
//    Set automationObject = addIn.Object
//    Dim key As String
//    key = InputBox("Please specify the key of the Test Execution you wish to reimport", "Import Test Execution")
//    automationObject.ReimportTestExecution(key)
//End Sub
