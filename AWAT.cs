using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using APA.Systemetix.MouseController;
using APA.Displayer.WinFormer;
using System.IO;
using APA.Systemetix;
using Microsoft.Win32;
using APA.Toolboxit;
using APA.Utilitizer;
using APA.Displayer.TreeViewser;
using System.Threading;
using APA.Network;
using APA.WAT.Common;
using System.Net;
using Gecko.DOM;
using Gecko;
using APA.Web.GeckoControl;
using System.Diagnostics;
using System.Runtime.InteropServices;
using WindowsInput.Native;
using WindowsInput;


namespace APA.WAT.Agent
{
    /// <summary>
    ///  Код формы Агента. 
    ///  Агент запускается контроллером с параметрами, инициализирует GeckoFX и начинает выполнять полученный сценарий.
    /// </summary>
    public partial class FormMain : Form
    {
        APAPluginCore PluginCore = new APAPluginCore();

        Color labelTitleForeColor;
        Color labelTitleBackColor;

        IServer botServer;
        TaskManager taskManager;

        AutomationConfiguration automationCFG;
        Automation automation;
        int automationMaxActions = 100000;
		
        Elementare geckoElementare;

        MouseController mouseController;
        LabelBorders formBorders;

        AutomationPreset targetPreset;

        bool stopTaskManager = false;
        Thread ActionsThread;
        StartActionsParameters startParameters;

        ConfiguratorHelper cfgHelper;

        Watchdog Watchdog;

        string newFirefoxProfileFolderPath;
        string newFirefoxEngineFolderPath;
        string resourceDirectoryPath;

        public string presetName = string.Empty;
        public string agentServerAddress = string.Empty;
        public string agentServerPort = string.Empty;
        public string browserDefaultAddress = string.Empty;
        public string browserProfileName = string.Empty;
        public string browserUserAgent = string.Empty;
        public string browserProxyIP = string.Empty;
        public string browserProxyPort = string.Empty;
        public string browserProxyUser = string.Empty;
        public string browserProxyPass = string.Empty;

        public string formHeight = string.Empty;
        public string formWidth = string.Empty;

        string remoteGuid = string.Empty;
        string automationFilePath = string.Empty;
        string patternGuid = string.Empty;
        string presetGuid = string.Empty;

        string controllerAddress = string.Empty;
        string userName = string.Empty;
        string variables = string.Empty;

        public FormMain(string[] args)
        {
            try
            {
                InitializeArguments(args);
                InitializeComponent();
                InitializeForm();

                LOGG.INFO("Initialize Automation");
                InitializeAutomation();

                LOGG.INFO("Initialize Mouse Controller");
                InitializeMouseController();

                LOGG.INFO("Initialize Server");
                InitializeServer();

                LOGG.INFO("Initialize Browser");
                InitializeBrowser();

                STATUS($"Application started");

                LOGG.INFO("Initialize Watchdog");
                InitializeWatchdog();

                geckoElementare = new Elementare(BrowserMain.Browser);

                SendTextToManager(ObjectInfo.Types.StartAgent, browserProfileName);

                LOGG.INFO("Application started");
            }
            catch(Exception exp)
            {
                LOGG.ERROR(exp.ToString());
                SendTextToManager(ObjectInfo.Types.ExitMessageAction, $"{userName}|Error|{exp.Message}|{browserProfileName}");
            }
        }

        void InitializeWatchdog()
        {
            Watchdog = new Watchdog(browserProfileName, IPAddress.Parse(controllerAddress.Split(':')[0]), Convert.ToInt32(controllerAddress.Split(':')[1]), 1000);
            Watchdog.Start();
        }

        void SendTextToManager(ObjectInfo.Types type, string message)
        {
            var client = TransmitterFactory.CreateTCPClient(IPAddress.Parse(controllerAddress.Split(':')[0]), Convert.ToInt32(controllerAddress.Split(':')[1]));
            ClientServerMessage messageToServer = new ClientServerMessage(type, message);
            client.SendObject(messageToServer);
        }

        object SendMessageToServer(IPAddress ip, int port, ObjectInfo.Types type, string message)
        {
            var client = TransmitterFactory.CreateTCPClient(ip, port);
            ClientServerMessage messageToServer = new ClientServerMessage(type, message);
            return client.SendObject(messageToServer);
        }

        void InitializeAutomation()
        {
            automation = new Automation();
            foreach (var variable in variables.Split(new string[] { "|||" }, StringSplitOptions.None))
            {
                automation.SetVariable(variable.Split(new string[] { "|=|" }, StringSplitOptions.None)[0], variable.Split(new string[] { "|=|" }, StringSplitOptions.None)[1]);
            }
        }


        void InitializeArguments(string[] args)
        {

            // Server 5 parameters
            presetName = args[0];
            agentServerAddress = args[1];
            agentServerPort = args[2];
            browserProfileName = args[3];
            browserUserAgent = args[4];
            browserDefaultAddress = args[5];

            // Proxy 4 parameters
            browserProxyIP =  args[6];
            browserProxyPort = args[7];
            browserProxyUser = args[8];
            browserProxyPass = args[9];

            // Form 2 parameters
            formHeight = args[10];
            formWidth = args[11];

            remoteGuid = args[12];

            automationFilePath = args[13];
            patternGuid = args[14];
            presetGuid = args[15];
            userName = args[16];
            controllerAddress = args[17];
            variables = args[18];

            PluginCore.LoadEntityPlugins(AppDomain.CurrentDomain.BaseDirectory + $"Plugin\\");

            LoadAutomationFile();

        }

        void LoadAutomationFile()
        {
            automationCFG = new AutomationConfiguration(PluginCore.EntityPluginsTypes, Path.GetDirectoryName(automationFilePath) + "\\", Path.GetFileNameWithoutExtension(automationFilePath), Path.GetExtension(automationFilePath), true);
            SetPatternPresetByGuid();
        }

        void InitializeMouseController()
        {
           
            mouseController = new MouseController();
        }

        void InitializeServer()
        {
            startParameters = new StartActionsParameters(null);
            startParameters.ContinueExecution = false;

            botServer = TransmitterFactory.CreateServer(IPAddress.Parse(agentServerAddress), Convert.ToInt32(agentServerPort));
            botServer.ObjectRecievedSuccessEvent += (obj) =>
            {
                try
                {
                    LOGG.INFO($"Agent got a message {((ClientServerMessage)obj).MessageType}");
                    switch (((ClientServerMessage)obj).MessageType)
                    {
                        case ObjectInfo.Types.Actions:
                            if (!startParameters.ContinueExecution)
                            {
                                
                                InitializeTaskManager();
                                var actions = (List<AutomationEntity>)((ClientServerMessage)obj).MessageObj;
                                

                                List<AutomationEntity> entitiestToExecute = new List<AutomationEntity>();
                                foreach(var action in actions)
                                {

                                    if (action is AutomationEntityGroup)
                                    {
                                        foreach (var gAction in ((AutomationEntityGroup)action).Entities)
                                        {
                                            entitiestToExecute.Add(gAction);
                                        }
                                    }
                                    else
                                    {
                                        entitiestToExecute.Add(action);
                                    }
                                }

                                LOGG.INFO($"Start actions: {entitiestToExecute.Count}");
                                StartActions(entitiestToExecute);
                            }
                            break;
                        case ObjectInfo.Types.StopExecution:
                            startParameters.StopExecution();
                            SetStatusBusy(false);
                            STATUS("Actions stopped");
                            LOGG.DEBUG($"Stop execution command");
                            break;
                        case ObjectInfo.Types.StopAgent:
                            this.BeginInvoke(new Action(() =>
                            {
                                LOGG.DEBUG($"Stop Agent command");
                                BrowserMain.Shutdown();
                               

                            }));
                            Thread.Sleep(1000);
                            Directory.Delete(newFirefoxProfileFolderPath, true);
                            Directory.Delete(newFirefoxEngineFolderPath, true);
                            Application.Exit();
                            break;

                    }
                }
                catch(Exception exp)
                {
                    LOGG.ERROR(exp.ToString());
                }

                return true;
            };

            botServer.Start();
        }

		void STATUS(string message, params object[] args)
		{
			lblStatus.Text = String.Format(message, args);
		}
		
        void StartActions(List<AutomationEntity> actions)
        {
            startParameters = new StartActionsParameters(actions);
            ActionsThread = new Thread(new ParameterizedThreadStart(ExecuteActions));
            ActionsThread.Start(startParameters);
        }

        void ExecuteActions(Object startParameters)
        {
            try
            {
                SetStatusBusy(true);

                STATUS("Start actions.");

                StartActionsParameters parameters = (StartActionsParameters)startParameters;
                LOGG.DEBUG($"Execute actions: {parameters.Actions.Count}");
                int currentActionIndex = 0;


                while (currentActionIndex < parameters.Actions.Count)
                {
                    if (!parameters.ContinueExecution)
                        break;

                    LOGG.DEBUG($"Current Action: {currentActionIndex + 1} of {parameters.Actions.Count}");
                    STATUS($"Current Action: {currentActionIndex + 1} {parameters.Actions[currentActionIndex].Name}");
                    currentActionIndex = ExecuteAction(parameters.Actions[currentActionIndex], currentActionIndex);
                }

                STATUS($"Done: with {parameters.Actions.Count} actions.");

                SendTextToManager(ObjectInfo.Types.ExitMessageAction, userName + $"|{automation.GetVariable("ExitMessage")}");
                SetStatusBusy(false);
            }
            catch (Exception exp)
            {
                STATUS($"Execution Error: {exp.Message}");
                LOGG.ERROR($"{exp.ToString()}");
                SetStatusBusy(false);
                SendTextToManager(ObjectInfo.Types.ExitMessageAction, $"{userName}|Error|{exp.Message}|{browserProfileName}");
            }
            
        }

        int ExecuteAction(AutomationEntity action, int currentActionIndex)
        {
            LOGG.DEBUG($"ExecuteAction  {action.Type}");
            currentActionIndex++;

            switch (action.Type)
            {
                case ObjectInfo.Types.EmptyAction:

                    break;
                case ObjectInfo.Types.SetProxyByVariableAction:
                    var spbvAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var proxyData = automation.GetVariable(((SetProxyByVariableAction)action).VariableName);
                            LOGG.DEBUG($"Set proxy to {proxyData}");
                            BrowserMain.SetProxy(proxyData.Split(':')[0], proxyData.Split(':')[1], proxyData.Split(':')[2], proxyData.Split(':')[3]);
                            LOGG.DEBUG($"Proxy applied successfully");
                            
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(spbvAction);
                    break;
                case ObjectInfo.Types.SetUAgentByVariableAction:
                    var saubvAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var uaData = automation.GetVariable(((SetUAgentByVariableAction)action).VariableName);
                            LOGG.DEBUG($"Set uagent to {uaData}");
                            BrowserMain.SetUserAgent(uaData);
                            LOGG.DEBUG($"Uagent applied successfully");
                            
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(saubvAction);
                    break;
                case ObjectInfo.Types.BrowserNavigateAction:
                    var bnTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        var result = BrowserMain.Navigate(((BrowserNavigateAction)action).TargetURL, 30);
                        result.Wait();

                        return result.Result;
                    });

                    TaskManager.WaitTaskComplete(bnTask);
                    break;
                case ObjectInfo.Types.BrowserNavigateVariableAction:
                    var bnvTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        var sourceAction = (BrowserNavigateVariableAction)action;
                        var urlResource = automation.GetVariable(sourceAction.VariableName);
                        var result = BrowserMain.Navigate(urlResource, 30);
                        result.Wait();

                        return result.Result;
                    });

                    TaskManager.WaitTaskComplete(bnvTask);
                    break;
                case ObjectInfo.Types.ScrollPageAction:
                    var spAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (ScrollPageAction)action;
                            if (sourceAction.X == "pages")
                            {
                                for(int sc=0; sc < Convert.ToInt32(sourceAction.Y); sc++)
                                {
                                    BrowserMain.Browser.Window.ScrollByPages(sc);
                                }                                
                            }
                            else if (sourceAction.X == "lines")
                            {
                                for (int sc = 0; sc < Convert.ToInt32(sourceAction.Y); sc++)
                                {
                                    BrowserMain.Browser.Window.ScrollByLines(sc);
                                }
                            }
                            else
                            {
                                BrowserMain.Browser.Window.ScrollBy(Convert.ToDouble(sourceAction.X), Convert.ToDouble(sourceAction.Y));
                            }
                            LOGG.DEBUG($"Page scrolled to {sourceAction.X} x {sourceAction.Y}.");
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(spAction);
                    break;
                case ObjectInfo.Types.SendTextToBrowserAction:
                    var stAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        
                            var sourceAction = (SendTextToBrowserAction)action;
                            Random rnd = new Random();
                            int timePrintCharDelayMS = rnd.Next(Convert.ToInt32(sourceAction.MinPrintDelayMS), Convert.ToInt32(sourceAction.MaxPrintDelayMS));
                            var text = automation.GetVariable(sourceAction.VariableName);
                            foreach (var ch in text)
                            {
                            this.BeginInvoke(new Action(() =>
                            {
                                SendKeyEvent("Type", ch.ToString(), false, false, false);
                            }));
                            System.Threading.Tasks.Task.Delay(timePrintCharDelayMS).Wait();
                            }
                        
                        return true;
                    });

                    TaskManager.WaitTaskComplete(stAction);
                    break;
                case ObjectInfo.Types.ResolveCaptchaAction:
                    var rcAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourcesAction = (ResolveCaptchaAction)action;
                            var recapi = automation.GetVariable(sourcesAction.KeyVariableName);
                            var result = geckoElementare.ResolveCaptcha(sourcesAction.XPath, recapi);
                            switch (result.State)
                            {
                                case ObjectInfo.States.Success:
                                    automation.SetVariable(sourcesAction.VariableName, result.Path);
                                    LOGG.DEBUG($"Captcha resolved succesfully");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    automation.SetVariable(sourcesAction.VariableName, "ERROR");
                                    LOGG.DEBUG($"Captcha not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    automation.SetVariable(sourcesAction.VariableName, "ERROR");
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(rcAction);
                    break;
                case ObjectInfo.Types.ScrollElementAction:
                    var seAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (ScrollElementAction)action;
                            var elements = BrowserMain.Browser.Document.EvaluateXPath(sourceAction.XPath);
                            var targetWebElement = elements.GetNodes().FirstOrDefault();
                            if (targetWebElement != null)
                            {
                                GeckoHtmlElement htmlElement = (GeckoHtmlElement)targetWebElement;
                                htmlElement.ScrollIntoView(true);
                                LOGG.DEBUG($"Scrolled to element.");
                            }
                            else
                            {
                                LOGG.DEBUG($"Element found");
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(seAction);
                break;
                case ObjectInfo.Types.CheckElementExistsAction:
                    LOGG.DEBUG($"case CheckElementExistsAction");
                    var cheAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var result = geckoElementare.CheckWebElementExists(((CheckElementExistsAction)action).XPath);
                            LOGG.DEBUG($"result state {result.State}");
                            switch (result.State)
                            {
                                case ObjectInfo.States.Success:
                                    automation.SetVariable(((CheckElementExistsAction)action).VariableName, "TRUE");
                                    LOGG.DEBUG($"Element found");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    automation.SetVariable(((CheckElementExistsAction)action).VariableName, "FALSE");
                                    LOGG.DEBUG($"Element not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(cheAction);
                    break;
                case ObjectInfo.Types.SetElementAttributeByValueAction:

                    var seavalAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourcesAction = (SetElementAttributeByValueAction)action;
                            var result = geckoElementare.SetWebElementText(sourcesAction.XPath, sourcesAction.AttributeValue);
                            switch(result.State)
                            {
                                case ObjectInfo.States.Success:
                                    LOGG.DEBUG($"Element Text setted succesfully, with message {result.Path}");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    LOGG.DEBUG($"Element not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                        }));
                        return true;
                       
                    });

                    TaskManager.WaitTaskComplete(seavalAction);
                    break;
                case ObjectInfo.Types.SetElementAttributeByVariableAction:

                    var seavarAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourcesAction = (SetElementAttributeByVariableAction)action;
                            string seavarVariableValue = automation.GetVariable(sourcesAction.VariableName);
                            var result = geckoElementare.SetWebElementValue(sourcesAction.XPath, seavarVariableValue);
                            switch (result.State)
                            {
                                case ObjectInfo.States.Success:
                                    LOGG.DEBUG($"Element setted succesfully with message {result.Path}");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    LOGG.DEBUG($"Element not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                        }));
                        return true;

                    });

                    TaskManager.WaitTaskComplete(seavarAction);
                    break;
                case ObjectInfo.Types.GetElementAttributeAction:
                    var geeAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (GetElementAttributeAction)action;
                            var result = geckoElementare.GetWebElementValueAttribute(sourceAction.XPath, sourceAction.AttributeName);
                            switch (result.State)
                            {
                                case ObjectInfo.States.Success:
                                    automation.SetVariable(((GetElementAttributeAction)action).VariableName, result.Path);
                                    LOGG.DEBUG($"Element got succesfully");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    LOGG.DEBUG($"Element not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                            
                        }));
                        return true;
                        
                    });

                    TaskManager.WaitTaskComplete(geeAction);
                    break;
                case ObjectInfo.Types.ClickElementAction:
                    var ceeAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (ClickElementAction)action;
                            var result = geckoElementare.ClickWebElement(sourceAction.XPath);
                            switch (result.State)
                            {
                                case ObjectInfo.States.Success:
                                    LOGG.DEBUG($"Element clicked succesfully");
                                    break;
                                case ObjectInfo.States.NotFound:
                                    LOGG.DEBUG($"Element not Found");
                                    break;
                                case ObjectInfo.States.Failure:
                                    LOGG.DEBUG($"Filure with message {result.Path}");
                                    break;
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(ceeAction);
                    break;
                case ObjectInfo.Types.MouseClickElementAction:
                    var mceaAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (MouseClickElementAction)action;
                            var elements = BrowserMain.Browser.Document.EvaluateXPath(sourceAction.XPath);
                            var targetWebElement = elements.GetNodes().FirstOrDefault();
                            if (targetWebElement != null)
                            {
                                GeckoHtmlElement htmlElement = (GeckoHtmlElement)targetWebElement;
                                var elementBound = htmlElement.GetBoundingClientRect();

                                Activate(true);

                                Mouse.SetCursorToControlPosition(BrowserMain.Browser, elementBound.X + elementBound.Width / 2 + Convert.ToInt32(sourceAction.Left), elementBound.Y + elementBound.Height / 2 - Convert.ToInt32(sourceAction.Top));
                                Mouse.ClickLeftButton();

                                Activate(false);
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(mceaAction);
                    break;
                case ObjectInfo.Types.MouseClickElementWidthAction:
                    var mcewaAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            var sourceAction = (MouseClickElementWidthAction)action;
                            var elements = BrowserMain.Browser.Document.EvaluateXPath(sourceAction.XPath);
                            var targetWebElement = elements.GetNodes().FirstOrDefault();
                            if (targetWebElement != null)
                            {
                                GeckoHtmlElement htmlElement = (GeckoHtmlElement)targetWebElement;
                                var elementBound = htmlElement.GetBoundingClientRect();

                                int width = Convert.ToInt32(formWidth) / 2 - Convert.ToInt32(sourceAction.Width)/2;

                                Activate(true);

                                Mouse.SetCursorToControlPosition(BrowserMain.Browser, elementBound.X + elementBound.Width / 2 + Convert.ToInt32(sourceAction.Left) + width, elementBound.Y + elementBound.Height / 2 - Convert.ToInt32(sourceAction.Top));
                                Mouse.ClickLeftButton();

                                Activate(false);
                            }
                        }));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(mcewaAction);
                    break;
                case ObjectInfo.Types.MouseMoveClick:
                    var mcTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        MouseController.ExecuteMouseMoveClick(this, ((MouseAction)action).Target, 20, 20, 10, 20, 10, 10);
                        return true;
                    });

                    TaskManager.WaitTaskComplete(mcTask);
                    break;
                case ObjectInfo.Types.MousePressLeft:
                    var plTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        Mouse.PressMouseLeftButton();
                        return true;
                    });

                    TaskManager.WaitTaskComplete(plTask);
                    break;
                case ObjectInfo.Types.MouseClickLeft:
                    var mclTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        Mouse.ClickLeftButton();
                        return true;
                    });

                    TaskManager.WaitTaskComplete(mclTask);

                    break;
                case ObjectInfo.Types.MouseMove:
                    var mTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        return MouseController.ExecuteMouseMove(this, ((MouseAction)action).Target, 20, 20, 10, 20, 10, 10);
                    });

                    TaskManager.WaitTaskComplete(mTask);

                    break;
                case ObjectInfo.Types.MouseUnpressLeft:
                    var ulTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        Mouse.UnpressMouseLeftButton();
                        return true;
                    });

                    TaskManager.WaitTaskComplete(ulTask);
                    break;
                case ObjectInfo.Types.WaitAction:
                    var wTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        Thread.Sleep(Convert.ToInt32(((WaitAction)action).TimeMilsec));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(wTask);
                    break;
                case ObjectInfo.Types.WaitRandomAction:
                    var wrTask = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        Random rnd = new Random();
                        var sourceAction = (WaitRandomAction)action;
                        Thread.Sleep(Convert.ToInt32(rnd.Next(Convert.ToInt32(sourceAction.TimeMilsecMin), Convert.ToInt32(sourceAction.TimeMilsecMax))));
                        return true;
                    });

                    TaskManager.WaitTaskComplete(wrTask);
                    break;
                case ObjectInfo.Types.PassCounterJumpAction:
                    var passCounterJumpAction = (PassCounterJumpAction)action;
                    string currentCounterVariable = automation.GetVariable(passCounterJumpAction.VariableName);
                    int currentCounter = currentCounterVariable == "undefined" ? 1 : Convert.ToInt32(currentCounterVariable);
                    currentCounter++;
                    automation.SetVariable(passCounterJumpAction.VariableName, currentCounter.ToString());
                    int jumpCounter = Convert.ToInt32(passCounterJumpAction.CounterValue);
                    if (currentCounter > jumpCounter)
                    {
                        currentActionIndex = Convert.ToInt32(passCounterJumpAction.TrueJump) - 1;
                        automation.SetVariable(passCounterJumpAction.VariableName, "1");
                        LOGG.DEBUG($"{currentCounterVariable} > {jumpCounter} goto {currentActionIndex} index");
                    }
                    break;
                case ObjectInfo.Types.ConditionalJumpAction:
                    var conditionJumpAction = (ConditionJumpAction)action;
                    switch (conditionJumpAction.Condition)
                    {
                        case "Contains":
                            string firstValue = automation.GetVariable(conditionJumpAction.VariableName);
                            if (firstValue.Contains(conditionJumpAction.VariableValue))
                            {
                                LOGG.DEBUG($"{conditionJumpAction.VariableValue} contains in variable {conditionJumpAction.VariableName} as {firstValue}");
                                currentActionIndex = Convert.ToInt32(conditionJumpAction.TrueJump) - 1;
                                LOGG.DEBUG($"Goto {currentActionIndex} index");
                            }
                            else
                            {
                                currentActionIndex = Convert.ToInt32(conditionJumpAction.FalseJump) - 1;
                            }
                            break;
                    }
                    break;
                case ObjectInfo.Types.GetVariableFromFileAction:
                    var gvffAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        var gvffSourceAction = (GetVariableFromFileAction)action;
                        int targetLine = Convert.ToInt32(gvffSourceAction.LineCount);
                        var lines = File.ReadAllLines(resourceDirectoryPath + gvffSourceAction.ResourceFileName);
                        if (targetLine > 0)
                        {
                            automation.SetVariable(gvffSourceAction.VariableName, lines[targetLine].Trim());
                        }
                        else
                        {
                            Random rnd = new Random();
                            targetLine = rnd.Next(0, lines.Length);
                            automation.SetVariable(gvffSourceAction.VariableName, lines[targetLine].Trim());
                        }
                        LOGG.DEBUG($"Get variable {gvffSourceAction.VariableName} from {gvffSourceAction.ResourceFileName} at line: {targetLine}");
                            return true;
                    });

                    TaskManager.WaitTaskComplete(gvffAction);
                    break;
                case ObjectInfo.Types.GetResourceFromRemoteAction:
                    var gvfrAction = taskManager.AddAction(System.Guid.NewGuid().ToString(), () =>
                    {
                        var gvfrSourceAction = (GetResourceFromRemoteAction)action;
                       
                        var ip = IPAddress.Parse(gvfrSourceAction.RemoteHost);
                        var port = Convert.ToInt32(gvfrSourceAction.RemotePort);
                        LOGG.DEBUG($"Get variables {gvfrSourceAction.VariablesScheme} from {ip.ToString()}: {port.ToString()} as {gvfrSourceAction.ResourceName} Resource.");

                        var response = SendMessageToServer(ip, port, ObjectInfo.Types.GetVariablesFromRemoteAction, $"{remoteGuid}|{gvfrSourceAction.ResourceName}");

                        LOGG.DEBUG($"Response = {response}");

                        
                        if ((string)response != "FALSE")
                        {
                            
                            if (gvfrSourceAction.VariablesScheme.Contains('|'))
                            {
                                var varNames = gvfrSourceAction.VariablesScheme.Split('|');
                                int varNameIndex = 0;
                                foreach (var val in ((string)response).Split('|'))
                                {
                                    automation.SetVariable(varNames[varNameIndex], val);
                                    varNameIndex++;
                                }
                            }
                            else
                            {
                                automation.SetVariable(gvfrSourceAction.VariablesScheme, (string)response);
                            }
                                
                            LOGG.DEBUG($"Get Resource from {gvfrSourceAction.ResourceName} as {(string)response} assigned to {gvfrSourceAction.VariablesScheme}");
                        }

                        return true;
                    });

                    TaskManager.WaitTaskComplete(gvfrAction);
                    break;
                case ObjectInfo.Types.SetVariableAction:
                    var setVariableAction = (SetVariableAction)action;
                    automation.SetVariable(setVariableAction.VariableName, setVariableAction.VariableValue);
                    LOGG.DEBUG($"Set variable {setVariableAction.VariableName} = {setVariableAction.VariableValue}");
                    break;
                case ObjectInfo.Types.SaveVariableAction:
                    var saveVariableAction = (SaveVariableAction)action;
                    var variableValue = automation.GetVariable(saveVariableAction.VariableName);
                    SendTextToManager(ObjectInfo.Types.SaveVariableAction, userName + "|" + variableValue);
                    LOGG.DEBUG($"Send variable {saveVariableAction.VariableName} = {variableValue} to manager");
                break;
                case ObjectInfo.Types.SaveTextAction:
                    var saveTextAction = (SaveTextAction)action;
                    SendTextToManager(ObjectInfo.Types.SaveTextAction, userName + "|" + saveTextAction.Text);
                    LOGG.DEBUG($"Send text {saveTextAction.Text} to manager");
                    break;
                case ObjectInfo.Types.ExitMessageAction:
                    var exitMessageAction = (ExitMessageAction)action;
                    automation.SetVariable("ExitMessage", exitMessageAction.TypeExit + "|" + exitMessageAction.TextExit + "|" + browserProfileName);
                    currentActionIndex = automationMaxActions;
                    LOGG.DEBUG($"Exit with message {exitMessageAction.TextExit} of type {exitMessageAction.TypeExit}");
                    break;
                case ObjectInfo.Types.Action:
                    var ae = (AutomationEntity)action;
                    ae.Execute(this, BrowserMain);
                    break;
            }

                return currentActionIndex;
        }


        void InitializeTaskManager()
        {
            taskManager = new TaskManager(20);
            taskManager.StartManager();
        }

        void Activate(bool status)
        {
            if(status)
            {
                this.TopMost = true;
                this.BringToFront();
                this.Activate();
                this.Focus();
                
            }
            else
            {
                this.TopMost = false;
            }

        }

        void SetStatusBusy(bool isBusy)
        {
            this.SafeInvoke(() =>
            {
                if (isBusy)
                {
                    startParameters.ContinueExecution = true;
                    formBorders.FormTitle.BackColor = Color.LightGreen;
                }
                else
                {
                    startParameters.ContinueExecution = false;
                    formBorders.FormTitle.BackColor = labelTitleBackColor;
                }
            });
        }


        void InitializeBrowser()
        {
            newFirefoxProfileFolderPath = AppDomain.CurrentDomain.BaseDirectory + $"BrowserProfiles\\{browserProfileName}\\Firefox32";
            newFirefoxEngineFolderPath = AppDomain.CurrentDomain.BaseDirectory + $"BrowserEngines\\{browserProfileName}_Firefox32Engine";
            resourceDirectoryPath = newFirefoxEngineFolderPath + "\\Resources\\";
            if (!Directory.Exists(newFirefoxProfileFolderPath))
                Directory.CreateDirectory(newFirefoxProfileFolderPath);

            if (!Directory.Exists(newFirefoxEngineFolderPath))
                Directory.CreateDirectory(newFirefoxEngineFolderPath);

            foreach (var fDataFile in Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory + "Firefox32"))
            {
                File.Copy(fDataFile, newFirefoxEngineFolderPath + "\\" + Path.GetFileName(fDataFile));
            }

            if (!Directory.Exists(resourceDirectoryPath))
                Directory.CreateDirectory(resourceDirectoryPath);

            foreach (var resourceFile in Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory + "Firefox32\\Resources"))
            {
                File.Copy(resourceFile, resourceDirectoryPath + Path.GetFileName(resourceFile));
            }

            // Extensions
            //var chromeDir = (nsIFile)Xpcom.NewNativeLocalFile(newFirefoxEngineFolderPath);
            //var chromeFile = chromeDir.Clone();
            //chromeFile.Append(new nsAString("chrome.manifest"));
            //Xpcom.ComponentRegistrar.AutoRegister(chromeFile);

            BrowserMain.Initialize(newFirefoxEngineFolderPath, newFirefoxProfileFolderPath, true);

            if (browserUserAgent != string.Empty)
                BrowserMain.SetUserAgent(browserUserAgent);

            if (browserProxyIP != string.Empty)
                BrowserMain.SetProxy(browserProxyIP, browserProxyPort, browserProxyUser, browserProxyPass);

            if (browserDefaultAddress != string.Empty)
                BrowserMain.Navigate(browserDefaultAddress, 30);
        }

        void SetPatternPresetByGuid()
        {
            foreach (var preset in automationCFG.parameters.Presets)
            {
               
                    if (preset.GUID == presetGuid)
                        targetPreset = preset;
            }

        }


        void InitializeForm()
        {
            labelTitleBackColor = Color.Gold;
            Form.CheckForIllegalCrossThreadCalls = false;

            formBorders = new LabelBorders(this);

            List<LabelTitleButton> userButtons = new List<LabelTitleButton>();

            /// Mouse button
            //userButtons.Add(new LabelTitleButton("M", 25, 15, Color.White, Color.DarkBlue, Color.Black, Color.DeepSkyBlue, () => 
            //{
            //    mouseController.ShowSetupForm(this, 0.6, true, false, ObjectInfo.Groups.InitializeActions, () => { LoadAutomationFile(); return targetPreset.Initialization; }, () => { automationCFG.Save(); });
            //}));

            /// OCR button
            //userButtons.Add(new LabelTitleButton("O", 25, 15, Color.White, Color.DarkBlue, Color.Black, Color.DeepSkyBlue, () =>
            //{
            //    ocrController.ShowSetupForm(this, 0.6, false, false, ObjectInfo.Groups.InitializeActions, () => { LoadAutomationFile(); return targetPreset.Initialization; }, () => { automationCFG.Save(); });
            //}));

            formBorders.Initialize
            (
                    null,
                    new LabelTitle(25, new Font("Segoe UI", 8), $"ATA:VER | {formWidth}x{formHeight} | {userName} ", ContentAlignment.MiddleCenter, 2, Color.Black, formBorders.GetFromToneDiapason(160, 245)),
                    null, // collapseButton
                    userButtons,
                    null, //new LabelTitleButton("-", 24, 20, Color.White, Color.Black, Color.Black, Color.Navy, null ),
                    null, //new LabelTitleButton('\u25A1'.ToString(), 24, 20, Color.White, Color.Black, Color.Black, Color.Navy, null ),
                    new LabelTitleButton("x", 25, 20, Color.Black, Color.Red, Color.Red, Color.Firebrick, null),
                    new LabelBorder(2, Color.Navy),
                    new LabelBorder(2, Color.Navy),
                    new LabelBorder(2, Color.Navy),
                    new LabelBorder(2, Color.Navy),
                    1,
                    false,
                    true
            );

            this.Height = Convert.ToInt32(formHeight);
            this.Width = Convert.ToInt32(formWidth);

        }

        public void SendKeyEvent(string type, string key, bool alt, bool ctrl, bool shift)
        {
            // Escape for JS.
            key = key.Replace("\\", "\\\\");
            var instance = Xpcom.CreateInstance<nsITextInputProcessor>("@mozilla.org/text-input-processor;1");
            using (var context = new AutoJSContext(BrowserMain.Browser.Window))
            {
                var result = context.EvaluateScript(
                    $@"new KeyboardEvent('', {{ key: '{key}', code: '', keyCode: 0, ctrlKey : {(ctrl ? "true" : "false")}, shiftKey : {(shift ? "true" : "false")}, altKey : {(alt ? "true" : "false")} }});");
                instance.BeginInputTransaction((mozIDOMWindow)((Gecko.GeckoWebBrowser)BrowserMain.Browser).Document.DefaultView.DomWindow, new TestCallback());
                instance.Keydown((nsIDOMEvent)result.ToObject(), 0, 1);
                instance.Keyup((nsIDOMEvent)result.ToObject(), 0, 1);
            }

            Marshal.ReleaseComObject(instance);
        }

        public class TestCallback : nsITextInputProcessorCallback
        {
            public bool OnNotify(nsITextInputProcessor aTextInputProcessor, nsITextInputProcessorNotification aNotification)
            {
                Console.WriteLine("TestCallback:OnNotify");
                return true;
            }
        }

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hWnd);
        private void FormMain_Load(object sender, EventArgs e)
        {

        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                taskManager.Stop();
            }
            catch
            { }

            try
            {
                botServer.Stop();
            }
            catch
            { }
        }
    }


    public static class FormsExt
    {
        public static void SafeInvoke(this Control instance, Action toDo)
        {
            if (instance.InvokeRequired)
            {
                instance.Invoke(toDo);
            }
            else
            {
                toDo();
            }
        }
    }


    public class Watchdog
    {
        public string Guid { get; set; }
        public System.Timers.Timer WatchdogTimer { get; set; }
        public TimeSpan Uptime { get; set; }
        public IPAddress ServerIP { get; set; }
        public int ServerPort { get; set; }
        public int TimerInterval { get; set; }

        public event Func<ClientServerMessage> TimerTickSendMessage;
        public Watchdog(string guid, IPAddress serverIP, int serverPort, int timerInterval)
        {
            Guid = guid;
            ServerIP = serverIP;
            ServerPort = serverPort;
            TimerInterval = timerInterval;

        }

        public void Start()
        {
            Uptime = new TimeSpan();
            WatchdogTimer = new System.Timers.Timer();
            WatchdogTimer.Interval = TimerInterval;
            WatchdogTimer.AutoReset = true;
            WatchdogTimer.Elapsed += (o, e) =>
            {
                try
                {
                    Uptime = Uptime.Add(TimeSpan.FromSeconds(1));
                    WatchdogTimer.Stop();
                    var message = new ClientServerMessage(ObjectInfo.Types.Unknown, null);   
                    if (TimerTickSendMessage != null)
                    {
                        message = TimerTickSendMessage?.Invoke();
                    }
                    else
                    {
                        message = new ClientServerMessage(ObjectInfo.Types.Watchdog, $"{Guid}|{GetUpTime()}");
                    }

                    SendMessageToServer(ServerIP, ServerPort, message);

                    WatchdogTimer.Start();
                }
                catch (Exception exp)
                {
                    LOGG.ERROR(exp.ToString());
                    WatchdogTimer.Start();
                }

            };

            WatchdogTimer.Start();
        }

        public string GetUpTime()
        {
            return Uptime.ToString(@"dd\.hh\:mm\:ss");
        }

        public void Stop()
        {
            try
            {
                WatchdogTimer.Stop();
            }
            catch (Exception e)
            { }
        }

        public object SendMessageToServer(IPAddress ip, int port, ClientServerMessage messageToServer)
        {
            try
            {
                var client = TransmitterFactory.CreateTCPClient(ip, port);
                return client.SendObject(messageToServer);
            }
            catch (Exception exp)
            {
                LOGG.ERROR(exp.ToString());
                return new ClientServerMessage(ObjectInfo.Types.Error, exp.ToString());
            }
        }

    }

}
