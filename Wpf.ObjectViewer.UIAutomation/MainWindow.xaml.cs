using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using Swa = System.Windows.Automation;
using System.Windows.Automation;
using System.Windows.Threading;
using System.Diagnostics;
using Gma.System.MouseKeyHook;
using System.Windows.Forms;

namespace Wpf.ObjectViewer.UIAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private delegate void SetMessageCallback(string message);
        private AutomationElement _rootElement;
        private AutomationElement _lastTopLevelWindow;
        private AutomationFocusChangedEventHandler _focusHandler;
        private AutomationEventHandler _clickEventHandler;
        private AutomationElement _elementParent;
        private Thread _mainThread;
        private Thread _automateThread;
        private IKeyboardMouseEvents _GlobalHook;
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            //initialize listview, clear items
            lvProperties.Items.Clear();
            
            _mainThread = Thread.CurrentThread;

            _automateThread = new Thread(new ThreadStart(Automate));
            _automateThread.Start();
            //while (Thread.CurrentThread.Equals(_automateThread))
            //{
            //    _automateThread.Abort();
            //    _mainThread.
            //}
            //this.WindowState = WindowState.Maximized;
        }



        public void SubscribeToFocusChange()
        {

            Automation.RemoveAllEventHandlers();
            _focusHandler = new AutomationFocusChangedEventHandler(OnFocusChanged);
            //_clickEventHandler = new AutomationEventHandler(OnClickChanged);
            MinimizeMainWindow();
            Automation.AddAutomationFocusChangedEventHandler(_focusHandler);

            //SubscribeToInvoke(_lastTopLevelWindow);

            //while (true)
            //{
            //    System.Threading.Thread.Yield();
            //}

            //this.WindowState = WindowState.Maximized;

        }

        public void SubscribeToInvoke(AutomationElement element)
        {
            if (element != null)
            {
                Automation.AddAutomationEventHandler(InvokePattern.InvokedEvent, _lastTopLevelWindow, TreeScope.Element, _clickEventHandler);
                _lastTopLevelWindow = element;
            }
        }
        private void Automate()
        {
            //LogMessage("Getting RootElement...");
            //_rootElement = AutomationElement.RootElement;
            //_lastTopLevelWindow = _rootElement;
            //if (_rootElement != null)
            //{
               SubscribeToFocusChange();
            //}
            //Subscribe();
        }

        private AutomationElement GetTextElement(AutomationElement parentElement, string value)
        {
            Swa.Condition condition = new PropertyCondition(AutomationElement.AutomationIdProperty, value);
            AutomationElement txtElement = parentElement.FindFirst(TreeScope.Descendants, condition);
            return txtElement;
        }

        private void DisplayLogMessage(string message)
        {
            lvProperties.Items.Add(message);
        }

        private void LogMessage(string message)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new SetMessageCallback(DisplayLogMessage), message);
        }
        private void NormalizeMainWindow()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(NormalizeMainWindowCallBack), null);
        }
        private void NormalizeMainWindowCallBack(IAsyncResult ar)
        {
            this.WindowState = WindowState.Maximized;
            this.WindowState = WindowState.Normal;
            this.Activate();
            this.Focus();
        }
        private void MaximizeMainWindow()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(MaximizeMainWindowCallBack), null);
        }
        private void MaximizeMainWindowCallBack(IAsyncResult ar)
        {
            this.WindowState = WindowState.Normal;
            this.WindowState = WindowState.Maximized;
            this.Activate();
            this.Focus();
        }
        private void MinimizeMainWindow()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(MinimizeMainWindowCallBack), null);
        }

        private void MinimizeMainWindowCallBack(IAsyncResult ar)
        {
            this.WindowState = WindowState.Normal;
            this.WindowState = WindowState.Minimized;
        }

        private void ClearListViewItems()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(ClearListViewItemsCallBack), null);
        }

        private void ClearListViewItemsCallBack(IAsyncResult ar)
        {
            lvProperties.Items.Clear();
        }

        private void WriteLogError()
        {
            LogMessage("ERROR." + Environment.NewLine);
        }


        /// <summary>
        ///     Retrieves the top-level window that contains the specified
        ///     UI Automation element.
        /// </summary>
        /// <param name="element">The contained element.</param>
        /// <returns>The  top-level window element.</returns>
        private AutomationElement GetTopLevelWindow(AutomationElement element)
        {
            var walker = TreeWalker.ControlViewWalker;
            //AutomationElement elementParent;
            var node = element;
            try // In case the element disappears suddenly, as menu items are 
                // likely to do.
            {
                if (node == AutomationElement.RootElement)
                {
                    return node;
                }
                // Save parent element first
                _elementParent = walker.GetParent(node);
                // Walk up the tree to the child of the root.
                int increment = 0;
                while (true)
                {
                    _elementParent = walker.GetParent(node);
                    if (_elementParent == null)
                    {
                        return null;
                    }
                    if (_elementParent == AutomationElement.RootElement)
                    {
                        break;
                    }

                    GetNodeProperties(node, increment);
                    node = _elementParent;
                    increment++;
                }
            }
            catch (ElementNotAvailableException)
            {
                node = null;
            }
            catch (ArgumentNullException)
            {
                node = null;
            }
            return node;
        }

        private void GetNodeProperties(AutomationElement node, int increment)
        {
            if (node == null)
            {
                return;
            }
            LogMessage("=======================================================================================");
            // If top-level window has changed, announce it.
            if (node != _lastTopLevelWindow)
            {
                _lastTopLevelWindow = node;
                LogMessage("Focus moved to top-level window:");
                LogMessage("  " + node.Current.Name);

            }
            else
            {
                // Announce focused element.
                LogMessage("Focused element: ");
                LogMessage("  Type: " +
                                  node.Current.LocalizedControlType);
                LogMessage("  Name: " + node.Current.Name);

            }
            LogMessage("=======================================================================================");
            int[] runtimeIDs = node.GetRuntimeId();
            LogMessage("Number:  " + increment);
            LogMessage("Class Name:  " + node.Current.ClassName);
            LogMessage("Control ID:  " + runtimeIDs[1].ToString());
            LogMessage("Control Name:  " + node.Current.Name);
            LogMessage("Parent Name:  " + _elementParent.Current.Name);
            LogMessage("Process ID:  " + node.Current.ProcessId);
            LogMessage("Process Name:  " + Process.GetProcessById(node.Current.ProcessId).ProcessName);
        }
        /// <summary>
        ///     Handles focus-changed events. If the element that received focus is
        ///     in a different top-level window, announces that. If not, just
        ///     announces which element received focus.
        /// </summary>
        /// <param name="src">Object that raised the event.</param>
        /// <param name="e">Event arguments.</param>
        private void OnFocusChanged(object src, AutomationFocusChangedEventArgs e)
        {
            Click(src);
        }

        private void OnClickChanged(object src, AutomationEventArgs e)
        {
            AutomationElement sourceElement;
            try
            {
                sourceElement = src as AutomationElement;
            }
            catch (ElementNotAvailableException)
            {
                return;
            }
            if (e.EventId == InvokePattern.InvokedEvent)
            {
                Click(src);
            }
            else
            { }
        }

        private void Click(object src)
        {
            try
            {
                LogMessage("Getting RootElement...");
                _rootElement = AutomationElement.RootElement;
                _lastTopLevelWindow = _rootElement;
                ClearListViewItems();
                var elementFocused = src as AutomationElement;
                GetTopLevelWindow(elementFocused);
                
            }
            catch (ElementNotAvailableException)
            {
                return;
            }
            finally
            {

                //Automation.RemoveAllEventHandlers(); //uncomment if you want to remove all EventHandlers
                NormalizeMainWindow(); //uncomment the code if you want to activate MainWindow and normalize it
                ResizeListViewColumns();
                //Thread.Sleep(1000);
                //Thread.CurrentThread.Abort(); //uncomment if you want to stop the current thread.

                Automation.RemoveAutomationFocusChangedEventHandler(_focusHandler);
                
            }
        }
        private void ResizeListViewColumnsCallBack(IAsyncResult ar)
        {
            GridView gv = lvProperties.View as GridView;
            if (gv != null)
            {
                foreach (GridViewColumn gvc in gv.Columns)
                {
                    gvc.Width = gvc.ActualWidth;
                    gvc.Width = double.NaN;
                }
            }
            
        }
        private void ResizeListViewColumns()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(ResizeListViewColumnsCallBack), null);
        }

    }
}
