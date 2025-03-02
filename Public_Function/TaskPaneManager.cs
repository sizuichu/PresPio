using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;

namespace PresPio.Public_Function
    {
    public class TaskPaneManager
        {
        private static TaskPaneManager _instance;
        private readonly Dictionary<int, Microsoft.Office.Tools.CustomTaskPane> _taskPanes;
        private readonly Dictionary<int, TaskPanelController> _controllers;

        public static TaskPaneManager Instance
            {
            get
                {
                if (_instance == null)
                    {
                    _instance = new TaskPaneManager();
                    }
                return _instance;
                }
            }

        private TaskPaneManager()
            {
            _taskPanes = new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();
            _controllers = new Dictionary<int, TaskPanelController>();
            }

        public Microsoft.Office.Tools.CustomTaskPane GetOrCreateTaskPane(string url, string title = "PresPio", int width = 450)
            {
            var hwnd = Globals.ThisAddIn.Application.ActiveWindow.HWND;

            if (_taskPanes.TryGetValue(hwnd, out var existingPane))
                {
                if (_controllers.TryGetValue(hwnd, out var controller))
                    {
                    controller.NavigateToUrl(url);
                    }
                return existingPane;
                }

            var windowControl = new TaskPanelController();
            var taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(windowControl, title);

            _taskPanes[hwnd] = taskPane;
            _controllers[hwnd] = windowControl;

            // 配置任务窗格
            taskPane.Width = width;
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            // 导航到URL
            windowControl.NavigateToUrl(url);

            return taskPane;
            }

        public void ShowTaskPane(string url, string title = "PresPio", int width = 450)
            {
            var taskPane = GetOrCreateTaskPane(url, title, width);
            taskPane.Visible = true;
            }

        public void HideTaskPane()
            {
            var hwnd = Globals.ThisAddIn.Application.ActiveWindow.HWND;
            if (_taskPanes.TryGetValue(hwnd, out var taskPane))
                {
                taskPane.Visible = false;
                }
            }

        public void CloseTaskPane()
            {
            var hwnd = Globals.ThisAddIn.Application.ActiveWindow.HWND;
            if (_taskPanes.TryGetValue(hwnd, out var taskPane))
                {
                taskPane.Visible = false;
                if (_controllers.TryGetValue(hwnd, out var controller))
                    {
                    controller.Dispose();
                    _controllers.Remove(hwnd);
                    }
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
                _taskPanes.Remove(hwnd);
                }
            }

        public void CloseAllTaskPanes()
            {
            foreach (var hwnd in _taskPanes.Keys.ToList())
                {
                if (_taskPanes.TryGetValue(hwnd, out var taskPane))
                    {
                    taskPane.Visible = false;
                    if (_controllers.TryGetValue(hwnd, out var controller))
                        {
                        controller.Dispose();
                        _controllers.Remove(hwnd);
                        }
                    Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
                    _taskPanes.Remove(hwnd);
                    }
                }
            }

        public bool IsTaskPaneVisible(int hwnd)
            {
            return _taskPanes.TryGetValue(hwnd, out var taskPane) && taskPane.Visible;
            }

        public TaskPanelController GetCurrentController()
            {
            var hwnd = Globals.ThisAddIn.Application.ActiveWindow.HWND;
            _controllers.TryGetValue(hwnd, out var controller);
            return controller;
            }
        }
    }