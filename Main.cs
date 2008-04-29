using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Xml.Serialization;
using System.IO;
using CustomUIControls;
using Rss;
using System.Text.RegularExpressions;
using System.Timers;
using System.Web;

namespace RSSReader
{
    public partial class Main : Form
    {
        delegate void addToItemsCallback(string icon, string description, string link, string title, DateTime pubdate);
        delegate void addToSortCallback();

        ImageList imageList = new ImageList();
        rssFeeds feedList;
        TaskbarNotifier taskbarNotifier;

        int textBoxDelayShowing = 500;
        int textBoxDelayStaying = 20000;
        int textBoxDelayHiding = 500;

        // Should check the feeds every 10seconds
        int rssTimerDelay = 10000;
        System.Timers.Timer rssTimer = new System.Timers.Timer();

        string messageTitle = "Default Title";
        string messageBody = "Default Body";
        string statusBarText = "";

        public Main()
        {
            InitializeComponent();
            this.Text = Application.ProductName.ToString() + " v" + Application.ProductVersion.ToString();
            notifyIcon.Text = Application.ProductName.ToString() + " v" + Application.ProductVersion.ToString();

            // Load the icons for the treeview
            statusBar("Loading Icons");
            LoadIcons();

            // Load in the RSS feeds
            statusBar("Loading RSS Feeds");
            LoadFeeds();

            // Setup the Taskbar Notifier
            statusBar("Loading Notifier");
            LoadNotifier();

            // Go fetch the RSS feeds
            statusBar("Fetching RSS Feeds");
            ReadRSS();

            // Start the timer
            statusBar("Starting RSS Timer");
            StartRSSTimer();

            // Save the RSS feeds
            SaveFeeds();
        }

        #region RSS Settings Class
        [XmlRoot("rssFeeds")]
        public class rssFeeds
        {
            private ArrayList listFeeds;

            public rssFeeds()
            {
                listFeeds = new ArrayList();
            }

            [XmlElement("feed")]
            public Item[] Feeds
            {
                get
                {
                    Item[] feeds = new Item[listFeeds.Count];
                    listFeeds.CopyTo(feeds);
                    return feeds;
                }
                set
                {
                    if (value == null) return;
                    Item[] feeds = (Item[])value;
                    listFeeds.Clear();
                    foreach (Item feed in feeds)
                        listFeeds.Add(feed);
                }
            }

            public int AddFeed(Item feed)
            {
                return listFeeds.Add(feed);
            }
        }

        public class Item
        {
            [XmlAttribute("name")]
            public string name;
            [XmlAttribute("url")]
            public string url;
            [XmlAttribute("icon")]
            public string icon;

            public Item()
            {
            }

            public Item(string Name, string Url, string Icon)
            {
                name = Name;
                url = Url;
                icon = Icon;
            }
        }
        #endregion

        #region Saving & Loading
        public void LoadFeeds()
        {
            XmlSerializer s = new XmlSerializer(typeof(rssFeeds));
            TextReader r = new StreamReader("feeds.xml");

            feedList = (rssFeeds)s.Deserialize(r);
            r.Close();

            BuildTree();
        }

        public void LoadIcons()
        {
            // Ultimatly we can't add icons to the program on the fly
            // So only do this on load (ie. when the program starts)
            foreach (DictionaryEntry resourceEntry in Properties.Resources.ResourceManager.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, true, true))
            {
                if (resourceEntry.Value.GetType().ToString() == "System.Drawing.Icon")
                {
                    imageList.Images.Add((System.Drawing.Icon)Properties.Resources.ResourceManager.GetObject(resourceEntry.Key.ToString()));
                    imageList.Images.SetKeyName(imageList.Images.Count - 1, resourceEntry.Key.ToString());
                }
            }
        }

        public void SaveFeeds()
        {
            rssFeeds myList = new rssFeeds();

            foreach (Item feed in feedList.Feeds)
            {
                myList.AddFeed(new Item(feed.name, feed.url, feed.icon));
            }

            // Serialization
            XmlSerializer s = new XmlSerializer(typeof(rssFeeds));
            TextWriter w = new StreamWriter(@"feeds.xml");
            s.Serialize(w, myList);
            w.Close();
        }
        #endregion

        #region GUI Controls
        public void BuildTree()
        {
            TreeNode rssFeeds = rssFeedList.Nodes.Add("Feeds");
            rssFeeds.ImageKey = "rss";
            rssFeeds.SelectedImageKey = "rss";
            rssFeeds.Expand();
            rssFeedList.ImageList = imageList;

            foreach (Item feed in feedList.Feeds)
            {
                TreeNode feedEntry = rssFeeds.Nodes.Add(feed.name);
                feedEntry.ToolTipText = feed.url;
                feedEntry.ImageKey = feed.icon;
                feedEntry.SelectedImageKey = feed.icon;
            }
        }

        public System.Drawing.Icon FindIcon(string name)
        {
            return (System.Drawing.Icon)Properties.Resources.ResourceManager.GetObject(name);
        }

        private void rssItemList_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < rssItemList.SelectedRows.Count; i++)
            {
                DataGridViewRow row = rssItemList.SelectedRows[i];

                for (int y = 0; y < row.Cells.Count; y++)
                {
                    DataGridViewCell cell = row.Cells[y];

                    if (cell.OwningColumn.Name.ToString() == "itemLink")
                    {
                        webBrowser.Url = new Uri(cell.Value.ToString());
                    }
                }
            }

        }

        private void addToItems(string icon, string description, string link, string title, DateTime pubdate)
        {
            if (this.rssItemList.InvokeRequired)
            {
                addToItemsCallback d = new addToItemsCallback(addToItems);
                this.Invoke(d, new object[] { icon, description, link, title, pubdate });
            }
            else
            {
                rssItemList.Rows.Add(FindIcon(icon), description, link, title, pubdate);
                SortRSS();
            }
        }

        private void Main_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
                Hide();
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Maximized;
        }

        private void statusBar(string statusText)
        {
            statusBarText = statusText;
            MethodInvoker mi = new MethodInvoker(statusBarTextThread);
            mi.BeginInvoke(null, null); // This will not block.
        }

        private void statusBarTextThread()
        {
            toolStripStatusLabel1.Text = statusBarText;
        }

        #endregion

        #region Taskbar Controls
        public void LoadNotifier()
        {
            taskbarNotifier = new TaskbarNotifier();
            taskbarNotifier.CloseClickable = true;
            taskbarNotifier.ContentClickable = true;
            taskbarNotifier.EnableSelectionRectangle = false;
            taskbarNotifier.KeepVisibleOnMousOver = true;
            taskbarNotifier.ReShowOnMouseOver = true;
            taskbarNotifier.SetBackgroundBitmap(Properties.Resources.skin, Color.FromArgb(255, 0, 255));
            taskbarNotifier.SetCloseBitmap(Properties.Resources.close, Color.FromArgb(255, 0, 255), new Point(281, 6));
            taskbarNotifier.TitleRectangle = new Rectangle(5, 1, 275, 25);
            taskbarNotifier.ContentRectangle = new Rectangle(10, 28, 280, 100);
            taskbarNotifier.TitleClick += new EventHandler(TitleClick);
            taskbarNotifier.ContentClick += new EventHandler(ContentClick);
        }

        void TitleClick(object obj, EventArgs ea)
        {
        }

        void ContentClick(object obj, EventArgs ea)
        {
        }

        void ShowPopup()
        {
            if ((messageTitle != "") && (messageBody != ""))
            {
                //messageBody.Replace("#
                taskbarNotifier.BringToFront();
                taskbarNotifier.Hide();
                taskbarNotifier.Show(Strip(messageTitle), Strip(messageBody), textBoxDelayShowing, textBoxDelayStaying, textBoxDelayHiding);
            }

        }

        void ShowPopup(object o, System.EventArgs e)
        {
            ShowPopup();
        }

        #endregion

        #region RSS Reading & Formatting
        public void StartRSSTimer()
        {
            rssTimer.Elapsed += new ElapsedEventHandler(CheckRSS);
            rssTimer.Interval = rssTimerDelay;
            rssTimer.Start();
        }

        public void CheckRSS(object source, ElapsedEventArgs e)
        {
            CheckRSS();
        }

        public void CheckRSS()
        {
            foreach (Item rssFeed in feedList.Feeds)
            {
                RssFeed siteFeed = RssFeed.Read(rssFeed.url);
                RssItemCollection items = siteFeed.Channels[0].Items;
                foreach (RssItem item in items)
                {
                    if (SearchRSSItems(item.Link.ToString()) == false)
                    {
                        addToItems(rssFeed.icon, item.Description, item.Link.ToString(), item.Title, item.PubDate);
                        messageBody = item.Description;
                        messageTitle = item.Title;
                        object[] pList = { this, System.EventArgs.Empty };
                        taskbarNotifier.BeginInvoke(new System.EventHandler(ShowPopup), pList);
                    }
                }
            }
        }
        
        public void ReadRSS()
        {
            MethodInvoker mi = new MethodInvoker(ReadRSSThread);
            mi.BeginInvoke(null, null); // This will not block.
        }

        public void ReadRSSThread()
        {
            rssItemList.Rows.Clear();
            rssItemList.RowsDefaultCellStyle.BackColor = Color.White;
            rssItemList.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

            foreach (Item rssFeed in feedList.Feeds)
            {
                RssFeed siteFeed = RssFeed.Read(rssFeed.url);
                RssItemCollection items = siteFeed.Channels[0].Items;
                foreach (RssItem item in items)
                {
                    addToItems(rssFeed.icon, Strip(item.Description), item.Link.ToString(), Strip(item.Title), item.PubDate);
                }
            }
            SortRSS();
        }

        public void SortRSS()
        {
            if (this.rssItemList.InvokeRequired)
            {
                addToSortCallback d = new addToSortCallback(SortRSS);
                this.Invoke(d, new object[] { });
            }
            else
            {
                rssItemList.Sort(rssItemList.Columns.GetLastColumn(DataGridViewElementStates.Visible, DataGridViewElementStates.None), ListSortDirection.Descending);
            }
        }

        public bool SearchRSSItems(string itemURL)
        {
            bool found = false;
            
            for (int i = 0; i < rssItemList.Rows.Count; i++)
            {
                DataGridViewRow row = rssItemList.Rows[i];

                for (int y = 0; y < row.Cells.Count; y++)
                {
                    DataGridViewCell cell = row.Cells[y];

                    if (cell.OwningColumn.Name.ToString() == "itemLink")
                    {
                        if (cell.Value.ToString() == itemURL)
                        {
                            found = true;
                            break;
                        }
                    }
                }
            }

            return found;
        }

        public string Strip(string text)
        {
            return Regex.Replace(HttpUtility.HtmlDecode(text), @"<(.|\n)*?>", string.Empty);
        }
        #endregion

    }
}