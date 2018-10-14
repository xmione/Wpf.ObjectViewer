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

namespace Wpf.ObjectViewer.MouseHook
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

        #region Constructor and Events
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

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Unsubscribe();
        }
        #endregion

        #region MouseKeyHook
        public void SubscribeCallBack(IAsyncResult ar)
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            _GlobalHook = Hook.GlobalEvents();

            _GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            //_GlobalHook.KeyPress += GlobalHookKeyPress; // not needed
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            LogMessage(string.Format("KeyPress: \t{0}", e.KeyChar));
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            //LogMessage(string.Format("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp));
            var element = GetElementFromCursor();
            GetTopLevelWindow(element);
            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        public void Unsubscribe()
        {
            _GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            _GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            _GlobalHook.Dispose();
        }
        #endregion MouseKeyHook

        #region CallBacks
        private void DisplayLogMessageCallBack(string message)
        {
            lvProperties.Items.Add(message);
        }

        private void LogMessage(string message)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new SetMessageCallback(DisplayLogMessageCallBack), message);
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

        private void Subscribe()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new AsyncCallback(SubscribeCallBack), null);
        }
        #endregion

        #region Other Methods
        private void Automate()
        {
            MinimizeMainWindow();
            Subscribe();
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
            finally
            {
                Unsubscribe();
                NormalizeMainWindow();
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

        private AutomationElement GetElementFromCursor()
        {
            // Convert mouse position from System.Drawing.Point to System.Windows.Point.
            System.Windows.Point point = new System.Windows.Point(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y);
            AutomationElement element = null;
            try
            {
                element = AutomationElement.FromPoint(point);
            }
            catch
            {
            }

            return element;
        }
        #endregion Other Methods

    }
}
