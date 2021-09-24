using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Condition = System.Windows.Automation.Condition;

namespace Pro7LooksHelper_Win10UIA
{
    //TODO: there seems to be no sensible way to regularly check the active look in Pro7 and keep the list refreshed - it's a one way street for now, you can set the look with this app,
    // but look changes in Pro7 will not be reflected in this app

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //  Lazy class level var.
        AutomationElementCollection aeLiveMenuItems;
        AutomationElement aeProPresenter;

        public MainWindow()
        {
            InitializeComponent();

            // Find ProPresenter (by it's name "ProPresenter")
            aeProPresenter = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "ProPresenter"));

            if (aeProPresenter != null) { 
                // Setup list of looks (TODO: consider refreshing this periodically)
                GetLooksListFromPro7UI();
            }

            DispatcherTimer timerRefresh = new DispatcherTimer();
            timerRefresh.Interval = TimeSpan.FromMilliseconds(500);
            timerRefresh.Tick += timerRefresh_Tick;
            timerRefresh.Start();
        }

        void timerRefresh_Tick(object sender, EventArgs e)
        {
            if (aeProPresenter == null)
            {
                // Find ProPresenter (by it's name "ProPresenter")
                aeProPresenter = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "ProPresenter"));

                if (aeProPresenter != null)
                {
                    // Setup list of looks (TODO: consider refreshing this periodically)
                    GetLooksListFromPro7UI();
                }
            }

            if (aeProPresenter != null) {
                // Try to get current look from the Looks window (Only works if looks window is open!)
                RefreshCurrentLook();
            }
        }

        private void RefreshCurrentLook()
        {
            // Setup a condition and find the first child Window named "Looks"
            Condition conditionLooksWindow = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.NameProperty, "Looks")
            );

            // To try find looks window
            AutomationElement aeLooksWindow = aeProPresenter.FindFirst(TreeScope.Children, conditionLooksWindow);
            //AutomationElement aeLiveLookList = aeLooksWindow.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));

            if (aeLooksWindow != null)
            {
                // Get the Looks List that has a single entry with TWO textblocks on a single line (One says "Live" the other has the live lookname)
                AutomationElement aeLiveLookList = aeLooksWindow.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ProPresenter.DO.Looks.AudienceLook"));
                
                // The second child (index 1) of above looks list will be an EditableTextBlock with the current look name
                AutomationElementCollection aecChildren = aeLiveLookList.FindAll(TreeScope.Children, Condition.TrueCondition);
                try
                {
                    //TODO: probably should first check if comboitem exists before trying to set it (eg if user has created a NEW look after running this app)
                    cboLooks.SelectedValue = GetAEText(aecChildren[1].FindFirst(TreeScope.Children, Condition.TrueCondition));
                }
                catch { }
            }
            
        }

        private string GetAEText(AutomationElement element)
        {
            // Thanks to Mike Zboray for this nice function to get text from text-control automation elements https://stackoverflow.com/questions/23850176/c-sharp-system-windows-automation-get-element-text
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }

        private void GetLooksListFromPro7UI()
        {
            cboLooks.Items.Clear();

            // Setup a condition and find the first menu named "Screens"
            Condition conditionScreensMenu = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                new PropertyCondition(AutomationElement.NameProperty, "Screens")
            );
            AutomationElement aeScreensMenu = aeProPresenter.FindFirst(TreeScope.Descendants, conditionScreensMenu);

            // If Screens menu is not currently expanded then expand it - as WPF menuItems need to be expanded to explore them.
            ExpandCollapsePattern expandPattern = aeScreensMenu.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
            if (expandPattern.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
                expandPattern.Expand();

            // Let's get all the sub menuitems of the Screens menu...
            AutomationElementCollection aecScreensMenuItems = aeScreensMenu.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

            // TODO: handle failure to get menu items better than this! (for now, just bail)
            if (aecScreensMenuItems.Count < 3)
                return;

            // If UI changes in design, this code will break - it assumes Looks menu is at index 2 of the screens menu items.
            // The third sub menuitem of the Screens menu is the Live:Look menu (well, at least in Pro7.4.1 - this may change in future)
            AutomationElement liveMenu = aecScreensMenuItems[2];

            // Grab the name of the current look
            String currentLookName = liveMenu.FindAll(TreeScope.Descendants, Condition.TrueCondition)[1].Current.Name;

            // If Live:Look menu is not currently expanded then expand it - as WPF menuItems need to be expanded to explore them.
            ExpandCollapsePattern expandPattern2 = liveMenu.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
            if (expandPattern2.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
                expandPattern2.Expand();

            // Let's get all the sub menuitems of the Live:Look menu (this is where we finally get a list of the looks)
            aeLiveMenuItems = liveMenu.FindAll(TreeScope.Children, Condition.TrueCondition);

            // - to get names, grab child textblock of each menuItem)
            AutomationElementCollection aeLiveMenuItems2 = aeLiveMenuItems[0].FindAll(TreeScope.Descendants, Condition.TrueCondition);

            // Add each look to a list....(ready to invoke later when clicked)
            foreach (AutomationElement liveMenuItem in aeLiveMenuItems)
            {
                this.cboLooks.Items.Add(liveMenuItem.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)).Current.Name);
            }

            if (currentLookName.Length > 0)
                this.cboLooks.SelectedItem = currentLookName;

            // Collapse Menu - we are done exploring..
            expandPattern2.Collapse();
            expandPattern.Collapse();

        }

        private void cboLooks_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboLooks.Items.Count > 0 && aeLiveMenuItems[cboLooks.SelectedIndex].Current.IsEnabled)
            {
                // Invoke the menu for selected look (indexes match)
                (aeLiveMenuItems[cboLooks.SelectedIndex].GetCurrentPattern(InvokePatternIdentifiers.Pattern) as InvokePattern).Invoke();

                // Put focus back on Pro7
                if (aeProPresenter != null) {
                    aeProPresenter.SetFocus();
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // Enable drag-move of window with mouse down on window itself
            if (e.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Top = Properties.Settings.Default.Top;
            this.Left = Properties.Settings.Default.Left;

            //TODO: Consider checking if location is visible on a screen and move to make visible if not (eg During last run, user positioned on screen that is now unplugged)
        }

        private void Window_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // Update position when user releases mouse (after dragging to new position)
            Properties.Settings.Default.Top = this.Top;
            Properties.Settings.Default.Left = this.Left;
            Properties.Settings.Default.Save();
        }


    }
}
