﻿using Microsoft.Office.Interop.MSProject;
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
        void ImportAllTestExecutions();
        void WriteToTextToPlain(String text10, String text11);
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        
        Dictionary<string, Dictionary<string, int>> tests = new Dictionary<string, Dictionary<string, int>>();

        Dictionary<string, string> textToField = new Dictionary<string, string>();

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

        public void scanProject()
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
            tests = content;
        }

        public void WriteToTextToPlain(string text10In, string text11In)
        {
            string text10 = text10In.ToLower().Trim().Replace(" ", "");
            string text11 = text11In.ToLower().Trim().Replace(" ", "");
            Dictionary<string, string> textToPlain = new Dictionary<string, string>();
            textToPlain.Add("text10", text10);
            textToPlain.Add("text11", text11);

            string jsonString = JsonConvert.SerializeObject(textToPlain);
            File.WriteAllText("C:\\Users\\anton.scheitler\\source\\repos\\MacroVSTO\\MacroVSTO\\texttoplain.json", jsonString);
        }

        public void updateTextToField()
        {
            string plainToFieldString = File.ReadAllText("C:\\Users\\anton.scheitler\\source\\repos\\MacroVSTO\\MacroVSTO\\plaintofield.json");
            string textToPlainString = File.ReadAllText("C:\\Users\\anton.scheitler\\source\\repos\\MacroVSTO\\MacroVSTO\\texttoplain.json");
            var plainToField = JsonConvert.DeserializeObject<Dictionary<string, string>>(plainToFieldString);
            var textToPlain = JsonConvert.DeserializeObject<Dictionary<string, string>>(textToPlainString);

            foreach (var text in textToPlain)
            {
                if (text.Value != "")
                {
                    textToField.Add(text.Key, plainToField[text.Value]);                
                }
            }
        }

        public async void ImportTestExecution(String key)
        {
            //try
            //{
                key = key.ToUpper();
                scanProject();
                Dictionary<string, int> testExecDictionary = new Dictionary<string, int>();
                parameters["testExecKey"] = key;
                string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);
                client.DefaultRequestHeaders.Add("Authorization", "Bearer MzE5MzM5OTcxMzYwOtThs7BNUYzG2JCRqFkFpiisVmes");
                var task = client.GetStringAsync(uri);
                String jsonString = task.GetAwaiter().GetResult();
                client.DefaultRequestHeaders.Remove("Authorization");
                var jsonArray = JArray.Parse(jsonString);
                if (!tests.ContainsKey(key))
                {
                    tests.Add(key, testExecDictionary);
                }

                foreach (var test in jsonArray)
                {
                    if (!testExecDictionary.ContainsKey(test["testKey"].ToString()))
                    {
                        testExecDictionary[test["testKey"].ToString()] = 1;
                    }
                    else
                    {
                        testExecDictionary[test["testKey"].ToString()] = testExecDictionary[test["testKey"].ToString()] + 1;
                    }
                    AddTest(test, key);
                }
                tests[key] = testExecDictionary;
        //}
        //    catch
        //    {
        //        MessageBox.Show("There's connection issues with Jira or the specified TestExecution does not exist");
        //        return;
        //    }

}

        public void ReimportTestExecution(String key)
        {
            application.SelectRow(1, false);
            if (key == null)
            {
                return;
            }
            key = key.ToUpper();
            scanProject();
            if (!tests.ContainsKey(key))
            {
                MessageBox.Show("Test Execution has not been imported\nThus it cannot be reimported");
                return;
            }

            Dictionary<string, int> testExecCopy = tests[key].ToDictionary(entry => entry.Key, entry => entry.Value);

            parameters["testExecKey"] = key;
            string uri = QueryHelpers.AddQueryString("raven/2.0/api/testruns", parameters);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer MzE5MzM5OTcxMzYwOtThs7BNUYzG2JCRqFkFpiisVmes");
            var task = client.GetStringAsync(uri);
            String jsonString = task.GetAwaiter().GetResult();
            client.DefaultRequestHeaders.Remove("Authorization");
            var jsonArray = JArray.Parse(jsonString);
            foreach (var test in jsonArray)
            {
                if (!testExecCopy.ContainsKey(test["testKey"].ToString()))
                {
                    tests[key][test["testKey"].ToString()] = 1;
                    testExecCopy[test["testKey"].ToString()] = 0;
                    AddTest(test, key);
                } else if (testExecCopy[test["testKey"].ToString()] == 0)
                {
                    tests[key][test["testKey"].ToString()] = tests[key][test["testKey"].ToString()] + 1;
                    AddTest(test, key);
                } else
                {
                    testExecCopy[test["testKey"].ToString()] = testExecCopy[test["testKey"].ToString()] - 1;
                }

            }
        }
        public void AddTest(JToken test, String testExecKey)
        {
            MSProject.Task newTask = application.ActiveProject.Tasks.Add(test["testIssueFields"]["summary"]);

            newTask.Start = test["testIssueFields"]["customfield_10128"].ToString();
            newTask.Finish = test["testIssueFields"]["customfield_10129"].ToString();
            newTask.Text19 = test["testKey"].ToString();
            newTask.Text20 = testExecKey;


            newTask.Text10 = test["testIssueFields"][textToField["text10"]].ToString();
            newTask.Text11 = test["testIssueFields"][textToField["text11"]].ToString();

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

        public async void ImportAllTestExecutions()
        {
            try
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer MzE5MzM5OTcxMzYwOtThs7BNUYzG2JCRqFkFpiisVmes");
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

            scanProject();
        }
    }
}

//Sub ReimportTestExecution()
//    Dim addIn As COMAddIn
//    Dim automationObject As Object
//    Set addIn = Application.COMAddIns("MacroVSTO")
//    Set automationObject = addIn.Object
//    Dim key As String
//    key = InputBox("Please specify the key of the Test Execution you wish to reimport", "Import Test Execution")
//    automationObject.ReimportTestExecution(key)
//End Sub
