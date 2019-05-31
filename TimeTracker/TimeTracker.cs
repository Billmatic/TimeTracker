using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using TimeTracker.Adaptors;
using TimeTracker.Common;
using TimeTracker.Entity;
using iTextSharp.text.pdf;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk.WebServiceClient;

namespace TimeTracker
{
    public partial class TimeTracker : Form
    {
        #region Private Variables

        /// <summary>
        /// 
        /// </summary>
        Session session;

        /// <summary>
        /// 
        /// </summary>
        CRMAdaptor crmAdaptor;

        /// <summary>
        /// 
        /// </summary>
        string[] crmGuids = new string[30];

        /// <summary>
        /// 
        /// </summary>
        string[] crmExternalCommentGuids = new string[30];

        /// <summary>
        /// The timer
        /// </summary>
        private Timer timer1 = new System.Windows.Forms.Timer();

        public string TimeTrackerMode = Constants.MODE_DEFAULT;

        /// <summary>
        /// To open a save dialog
        /// </summary>
        private SaveFileDialog saveDialog = new SaveFileDialog();

        /// <summary>
        /// To open a new file dialog
        /// </summary>
        private OpenFileDialog openDialog = new OpenFileDialog();

        #endregion

        #region Constructor
        public TimeTracker()
        {
            this.session = new Session();
            this.crmAdaptor = new CRMAdaptor();
            this.crmAdaptor.GetCRMConnection();
            GetTimeTrackerMode(new Uri("https://csp.api.crm.dynamics.com/XRMServices/2011/Organization.svc"));
            timer1.Interval = 1000;
            timer1.Tick += Timer_Tick;
            InitializeComponent();
            GetData();
        }
        #endregion

        #region Event Handlers

        protected void isBillableCheckBox_Click(object sender, EventArgs e)
        {
            CalcuateTotals();
        }

        private void GetTimeTrackerMode(Uri crmUri)
        {
            if (crmUri == null)
            {
                this.TimeTrackerMode = Constants.MODE_DEFAULT;
            }
            else if (crmUri.ToString() == Constants.JarvisOrgURI)
            {
                this.TimeTrackerMode = Constants.MODE_JARVIS;
            }
            else
            {
                this.TimeTrackerMode = Constants.MODE_CRM_DEFAULT;
            }
        }

        protected void isSynchedCheckBox_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            if(checkBox.Checked == true)
            {
                checkBox.BackColor = Color.Green;
            }
            else
            {
                checkBox.BackColor = Color.Red;
            }
        }

        protected void newItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateNewItem();
        }

        protected void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Do you want to start a new time tracker?", "Open a new doc", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                AutoSave();
                MainPanel.Controls.Clear(); //to remove all controls
                session.x = 0;
                session.y = 0;
                session.index = 0;
                session.activeIndex = 0;
                session.totalLines = 0;
                session.isTimerActive = false;
                session.fileDirectory = "";
                GetData();
                totalTimeTextBox.Text = "";
                billableTimeTextBox.Text = "";
                crmGuids = new string[30];
                crmExternalCommentGuids = new string[30];
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }

        }

        protected void CreateCaseTask(object sender, EventArgs e)
        {
            try
            {
                CRMAdaptor crmAdaptor = new CRMAdaptor();
                bool isValid = crmAdaptor.GetCRMConnection();

                if (!isValid)
                {
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;

                List<TimeItem> workItems = MapTimeItems();
                workItems = crmAdaptor.CreateCaseTask(workItems);

                foreach (TimeItem item in workItems)
                {
                    if (item.crmTaskId != null)
                    {
                        crmGuids[(int)item.index] = item.crmTaskId != null ? item.crmTaskId.ToString() : null;
                        crmExternalCommentGuids[(int)item.index] = item.crmExternalCommentId != null ? item.crmExternalCommentId.ToString() : null;
                        CheckBox submitted = this.Controls.Find(Constants.isSubmitted + item.index, true).FirstOrDefault() as CheckBox;
                        submitted.Checked = (bool)item.isCRMSubmitted;
                        submitted.Appearance = Appearance.Button;
                        submitted.BackColor = Color.Green;
                    }
                }
                AutoSave();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+"\n"+ex.InnerException.Message);
            }
        }

        protected void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Get file name.
            crmGuids = new string[30];
            crmExternalCommentGuids = new string[30];
            session.fileDirectory = openDialog.FileName;
            XmlTextReader reader = new XmlTextReader(session.fileDirectory);
            MainPanel.Controls.Clear(); //to remove all controls
            session.x = 0;
            session.y = 0;
            session.index = 0;
            session.activeIndex = 0;
            session.totalLines = 0;
            session.isTimerActive = false;
            string element = "";

            CreateLabelControl("projectLabel_" + session.index, "Project Name", session.x + 3, session.y, 30, 120);
            CreateLabelControl("titleLabel_" + session.index, "Task Name", session.x + 130, session.y, 30, 195);
            CreateLabelControl("timeLabel_" + session.index, "Time", session.x + 420, session.y, 30, 60);
            CreateLabelControl("noteLabel_" + session.index, "Notes", session.x + 503, session.y, 30, 430);
            CreateLabelControl("BillableLabel_" + session.index, "Bill", session.x + 938, session.y, 30, 30);
            CreateLabelControl("externalCommentLabel_" + session.index, "Ext.\nComm.", session.x + 970, session.y, 30, 50);
            CreateLabelControl("submittedLabel_" + session.index, "CRM \nSynched", session.x + 1030, session.y, 30, 100);

            session.y += 18;


            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    element = reader.Name;

                    if (element == "line")
                    {
                        CreateNewItem();
                    }
                }
                else if (reader.NodeType == XmlNodeType.Text)
                {
                    switch (element)
                    {
                        case "project": //Display the text in each element.
                            TextBox projectText = this.Controls.Find(Constants.project + (session.index - 1), true).FirstOrDefault() as TextBox;
                            projectText.Text = reader.Value;
                            break;
                        case "title": //Display the text in each element.
                            TextBox titleText = this.Controls.Find(Constants.title + (session.index - 1), true).FirstOrDefault() as TextBox;
                            titleText.Text = reader.Value;
                            break;
                        case "time": //Display the end of the element.
                            MaskedTextBox timeText = this.Controls.Find(Constants.time + (session.index - 1), true).FirstOrDefault() as MaskedTextBox;
                            timeText.Text = reader.Value;
                            break;
                        case "description": //Display the end of the element.
                            TextBox descriptionText = this.Controls.Find(Constants.description + (session.index - 1), true).FirstOrDefault() as TextBox;
                            descriptionText.Text = reader.Value;
                            break;
                        case "billable": //Display the end of the element.
                            CheckBox billableCheckBox = this.Controls.Find(Constants.isBillable + (session.index - 1), true).FirstOrDefault() as CheckBox;
                            billableCheckBox.Checked = Convert.ToBoolean(reader.Value);
                            break;
                        case "externalcomment": //Display the end of the element.
                            CheckBox externalCommentCheckBox = this.Controls.Find(Constants.isExternalComment + (session.index - 1), true).FirstOrDefault() as CheckBox;
                            externalCommentCheckBox.Checked = Convert.ToBoolean(reader.Value);
                            break;
                        case "CRMSubmitted": //Display the end of the element.
                            CheckBox CRMSubmittedCheckBox = this.Controls.Find(Constants.isSubmitted + (session.index - 1), true).FirstOrDefault() as CheckBox;
                            CRMSubmittedCheckBox.Checked = Convert.ToBoolean(reader.Value);
                            if(CRMSubmittedCheckBox.Checked == true) { CRMSubmittedCheckBox.BackColor = Color.Green; }
                            break;
                        case "CRMGuid": //Display the end of the element.
                            crmGuids[session.index - 1] = reader.Value;
                            break;
                        case "crmExternalCommentId": //Display the end of the element.
                            crmExternalCommentGuids[session.index - 1] = reader.Value;
                            break;
                    }

                    CalcuateTotals();
                    HideControlsByMode(this.TimeTrackerMode);
                }
            }

            reader.Close();
        }

        private void HideControlsByMode(string mode)
        {
            if (mode == Constants.MODE_DEFAULT)
            {
                Label externalCommentLabel = this.Controls.Find("externalCommentLabel_0", true).FirstOrDefault() as Label;
                externalCommentLabel.Visible = false;

                Label CRMSubmittedLabel = this.Controls.Find("submittedLabel_0", true).FirstOrDefault() as Label;
                CRMSubmittedLabel.Visible = false;

                button1.Visible = false;

                for (int i = 0; i<= session.index-1;i++)
                {
                    CheckBox externalCommentCheckBox = this.Controls.Find(Constants.isExternalComment + (session.index - 1), true).FirstOrDefault() as CheckBox;
                    externalCommentCheckBox.Visible = false;
                    CheckBox CRMSubmittedCheckBox = this.Controls.Find(Constants.isSubmitted + (session.index - 1), true).FirstOrDefault() as CheckBox;
                    CRMSubmittedCheckBox.Visible = false;
                }

                this.Width = 1020;

            }
        }

        protected void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openDialog.ShowDialog();
        }

        protected void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveDialog.ShowDialog();
        }

        protected void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            // Get file name.
            session.fileDirectory = saveDialog.FileName;

            //Generate file content
            string text = GenerateXML();
            session.fileDirectory = session.fileDirectory.Replace(".xml", "");
            File.WriteAllText(session.fileDirectory + ".xml", text);
        }

        protected void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(session.fileDirectory))
            {
                saveAsToolStripMenuItem_Click(null, EventArgs.Empty);
                return;
            }

            string text = GenerateXML();
            session.fileDirectory = session.fileDirectory.Replace(".xml", "");
            File.WriteAllText(session.fileDirectory + ".xml", text);
        }

        protected void TimerStart(object sender, EventArgs e)
        {
            Button button = sender as Button;
            button.Text = "Stop";
            timer1.Start();
            timer1.Enabled = true;
            button.Click -= new EventHandler(TimerStart);
            button.Click += new EventHandler(TimerStop);
            string activeIndexString = button.Name.Replace("btn_", "");
            session.activeIndex = Convert.ToInt32(activeIndexString);
            DisableAllButtons();
            button.BackColor = System.Drawing.Color.Red;
            button.Enabled = true;
            session.isTimerActive = true;

            //Set row to not sycnhed.
            string ControlName = Constants.isSubmitted + session.activeIndex;

            CheckBox isSubmittedCheckBox = this.Controls.Find(ControlName, true).FirstOrDefault() as CheckBox;
            isSubmittedCheckBox.Checked = false;
            isSubmittedCheckBox.Appearance = Appearance.Button;
            isSubmittedCheckBox.BackColor = Color.Red;
        }

        protected void TimerStop(object sender, EventArgs e)
        {
            Button button = sender as Button;
            button.Text = "Start";
            button.BackColor = System.Drawing.Color.White;
            timer1.Stop();
            timer1.Enabled = false;
            button.Click -= new EventHandler(TimerStop);
            button.Click += new EventHandler(TimerStart);
            ActivateAllButtons();
            session.isTimerActive = false;
            AutoSave();
        }

        protected void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                string ControlName = Constants.time + session.activeIndex;

                MaskedTextBox DigiClockTextBox = this.Controls.Find(ControlName, true).FirstOrDefault() as MaskedTextBox;
                if (DigiClockTextBox == null)
                {
                    return;
                }

                double updateSeconds = TimeSpan.Parse(DigiClockTextBox.Text).TotalSeconds;
                updateSeconds++;
                TimeSpan t = TimeSpan.FromSeconds(updateSeconds);
                string answer = string.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                DigiClockTextBox.Text = answer;
                CalcuateTotals();

            }
            catch 
            {
                timer1.Stop();
                DeactivateAllButtons();
            }
        }

        #endregion

        #region Private Methods

        private void ActivateAllButtons()
        {
            int buttonIndex = 0;
            while (buttonIndex < session.totalLines)
            {
                Button button = this.Controls.Find("btn_" + buttonIndex, true).FirstOrDefault() as Button;
                button.Enabled = true;
                buttonIndex++;
            }
        }

        private void AutoSave()
        {
            if (String.IsNullOrEmpty(session.fileDirectory))
            {
                return;
            }

            //Generate file content
            string text = GenerateXML();
            session.fileDirectory = session.fileDirectory.Replace(".xml", "");
            File.WriteAllText(session.fileDirectory + ".xml", text);

        }

        private void CalcuateTotals()
        {
            int i = 0;
            double billableSeconds = 0;
            double totalSeconds = 0;
            TimeSpan t = new TimeSpan();
            while (i < session.totalLines)
            {
                CheckBox isBillable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                MaskedTextBox time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as MaskedTextBox;
                if (isBillable.Checked == true)
                {

                    billableSeconds += TimeSpan.Parse(time.Text).TotalSeconds;
                }

                totalSeconds += TimeSpan.Parse(time.Text).TotalSeconds;
                i++;
            }
            t = TimeSpan.FromSeconds(billableSeconds);
            billableTimeTextBox.Text = string.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            t = TimeSpan.FromSeconds(totalSeconds);
            totalTimeTextBox.Text = string.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
        }

        private void CreateNewItem()
        {
            CreateTextBoxControl(Constants.project + session.index,"", 3 + session.x, 14 + session.y,30,120);
            CreateTextBoxControl(Constants.title + session.index, "", 130 + session.x, 14 + session.y, 30, 195);
            Button btn = CreateButtonControl("btn_" + session.index, "Start", 330 + session.x, 13 + session.y,23, 80);
            btn.Click += new EventHandler(TimerStart);
            CreateTimerTextBox(Constants.time + session.index, 420 + session.x, 14 + session.y, 30,60);
            CreateTextBoxControl(Constants.description + session.index,"", 503 + session.x, 14 + session.y, 30,430);
            CheckBox isBillableCheckBox = CreateCheckBoxControl(Constants.isBillable + session.index, 943 + session.x, 14 + session.y,25,30);
            isBillableCheckBox.Click += isBillableCheckBox_Click;
            CreateCheckBoxControl(Constants.isExternalComment + session.index, 975 + session.x, 14 + session.y, 25,30);
            CheckBox isSubmittedCheckBox = CreateCheckBoxControl(Constants.isSubmitted + session.index, 1040 + session.x, 14 + session.y, 25,30);
            isSubmittedCheckBox.Appearance = Appearance.Button;
            isSubmittedCheckBox.BackColor = Color.Red;
            isSubmittedCheckBox.Click += new System.EventHandler(this.isSynchedCheckBox_Click);

            session.y += 22;
            session.index++;
            session.totalLines++;

            if (session.isTimerActive)
            {
                DisableAllButtons();
                Button button = this.Controls.Find("btn_" + session.activeIndex, true).FirstOrDefault() as Button;
                button.Enabled = true;
                button.BackColor = System.Drawing.Color.Red;
            }

            HideControlsByMode(this.TimeTrackerMode);
        }

        private void DisableAllButtons()
        {
            int buttonIndex = 0;
            while (buttonIndex < session.totalLines)
            {
                Button button = this.Controls.Find("btn_" + buttonIndex, true).FirstOrDefault() as Button;
                button.Enabled = false;
                button.BackColor = System.Drawing.Color.White;
                buttonIndex++;
            }
        }

        private void DeactivateAllButtons()
        {
            int buttonIndex = 0;
            while (buttonIndex < session.totalLines)
            {
                Button button = this.Controls.Find("btn_" + buttonIndex, true).FirstOrDefault() as Button;
                button.Enabled = true;
                button.BackColor = System.Drawing.Color.White;
                buttonIndex++;
            }
        }

        private string GenerateXML()
        {
            string xml = "";
            xml += "<?xml version=\"1.0\" encoding=\"UTF-8\"?><!--IIS configuration sections.For schema documentation, see %windir%\\system32\\inetsrv\\config\\schema\\IIS_schema.xml.  Please make a backup of this file before making any changes to it.-->";
            xml += "<projects>";
            int i = 0;
            while (i < session.totalLines)
            {
                TextBox project = this.Controls.Find(Constants.project + i, true).FirstOrDefault() as TextBox;
                TextBox title = this.Controls.Find(Constants.title + i, true).FirstOrDefault() as TextBox;
                TextBox description = this.Controls.Find(Constants.description + i, true).FirstOrDefault() as TextBox;
                MaskedTextBox time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as MaskedTextBox;
                CheckBox billable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                CheckBox externalComments = this.Controls.Find(Constants.isExternalComment + i, true).FirstOrDefault() as CheckBox;
                CheckBox submitted = this.Controls.Find(Constants.isSubmitted + i, true).FirstOrDefault() as CheckBox;

                xml += "<line>";
                xml += "<project>" + project.Text + "</project>";
                xml += "<title>" + title.Text + "</title>";
                xml += "<description>" + description.Text + "</description>";
                xml += "<time>" + time.Text + "</time>";
                xml += "<billable>" + billable.Checked + "</billable>";
                xml += "<CRMSubmitted>" + submitted.Checked + "</CRMSubmitted>";

                if (crmGuids[i] != null)
                {
                    xml += "<CRMGuid>" + crmGuids[i].ToString() + "</CRMGuid>";
                }

                xml += "<externalcomment>" + externalComments.Checked + "</externalcomment>";

                if (crmExternalCommentGuids[i] != null)
                {
                    xml += "<crmExternalCommentId>" + crmExternalCommentGuids[i].ToString() + "</crmExternalCommentId>";
                }
                xml += "</line>";

                i++;
            }

            xml += "</projects>";
            return xml;
        }

        private Label CreateLabelControl(string controlPrefix, string LabelName, int posX, int posY, int height, int width)
        {
            Label label = new Label();
            label.Name = controlPrefix;
            label.Text = LabelName;
            label.Height = height;
            label.Width = width;
            label.Location = new Point(posX, posY);
            label.Font = new Font(label.Font, FontStyle.Bold);
            MainPanel.Controls.Add(label);
            return label;
        }

        private TextBox CreateTextBoxControl(string controlPrefix, string LabelName, int posX, int posY, int height, int width)
        {
            TextBox textBox = new TextBox();
            textBox.Name = controlPrefix;
            textBox.Text = "";
            textBox.Width = width;
            textBox.Height = height;
            textBox.Location = new Point(posX, posY);
            MainPanel.Controls.Add(textBox);
            return textBox;
        }

        private CheckBox CreateCheckBoxControl(string controlPrefix, int posX, int posY, int height, int width)
        {
            CheckBox checkBox = new CheckBox();
            checkBox.Name = controlPrefix;
            checkBox.Location = new Point(posX, posY);
            checkBox.Width = width;
            checkBox.Height = height;
            MainPanel.Controls.Add(checkBox);
            return checkBox;
        }

        private Button CreateButtonControl(string controlName,string buttonLabel, int posX, int posY, int height, int width)
        {
            Button btn = new Button();
            btn.Name = controlName;
            btn.Text = buttonLabel;
            btn.Location = new Point(posX, posY);
            btn.BackColor = System.Drawing.Color.White;
            btn.Height = height;
            btn.Width = width;
            MainPanel.Controls.Add(btn);
            return btn;
        }

        private MaskedTextBox CreateTimerTextBox(string controlName, int posX, int posY, int height, int width)
        {
            MaskedTextBox textBox = new MaskedTextBox();
            textBox.Name = controlName;
            textBox.Mask = "00:00:00";
            textBox.Text = "00:00:00";
            textBox.Width = width;
            textBox.Height = height;
            textBox.Location = new Point(posX, posY);
            MainPanel.Controls.Add(textBox);
            return textBox;
        }

        private void GetData()
        {
            CreateLabelControl("projectLabel_" + session.index, "Project Name", session.x+3, session.y, 30,120);
            CreateTextBoxControl(Constants.project + session.index, "", session.x + 3, session.y+30, 30, 120);
            CreateLabelControl("titleLabel_" + session.index, "Task Name",  session.x+130, session.y, 30,195);
            CreateTextBoxControl(Constants.title + session.index, "", session.x + 130, session.y + 30, 30, 195);
            Button btn = CreateButtonControl("btn_" + session.index, "Start", 330 + session.x, 29 + session.y, 23, 80);
            btn.Click += new EventHandler(TimerStart);
            CreateLabelControl("timeLabel_" + session.index, "Time", session.x + 420, session.y, 30, 60);
            CreateTimerTextBox(Constants.time + session.index, 420 + session.x, 30 + session.y, 30, 60);
            CreateLabelControl("noteLabel_" + session.index, "Notes", session.x + 503, session.y, 30, 430);
            CreateTextBoxControl(Constants.description + session.index, "", session.x + 503, session.y + 30, 30, 430);
            CreateLabelControl("BillableLabel_" + session.index, "Bill", session.x + 938, session.y, 30, 30);
            CheckBox isBillableCheckBox  = CreateCheckBoxControl(Constants.isBillable + session.index, session.x + 943, session.y + 30, 25, 30);
            isBillableCheckBox.Click += isBillableCheckBox_Click;
            CreateLabelControl("externalCommentLabel_" + session.index, "Ext.\nComm.", session.x + 970, session.y, 30, 50);
            CreateCheckBoxControl(Constants.isExternalComment + session.index, session.x + 975, session.y + 30, 25, 30);
            CreateLabelControl("submittedLabel_" + session.index, "CRM \nSynched", session.x + 1030, session.y, 30, 100);
            CheckBox isSubmittedCheckBox  = CreateCheckBoxControl(Constants.isSubmitted + session.index, session.x + 1040, session.y + 30, 25, 30);
            isSubmittedCheckBox.Appearance = Appearance.Button;
            isSubmittedCheckBox.BackColor = Color.Red;
            isSubmittedCheckBox.Click += new System.EventHandler(this.isSynchedCheckBox_Click);

            session.y += 38;
            session.index++;
            session.totalLines++;
            saveDialog.FileOk += saveFileDialog_FileOk;
            openDialog.FileOk += openFileDialog_FileOk;

            if (session.isTimerActive)
            {
                DisableAllButtons();
                Button button = this.Controls.Find("btn_" + session.activeIndex, true).FirstOrDefault() as Button;
                button.Enabled = true;
                button.BackColor = System.Drawing.Color.Red;
            }

            HideControlsByMode(this.TimeTrackerMode);
        }

        private TimeItem BuildTimeItem()
        {
            TextBox title = this.Controls.Find(Constants.title + session.activeIndex, true).FirstOrDefault() as TextBox;
            TextBox description = this.Controls.Find(Constants.description + session.activeIndex, true).FirstOrDefault() as TextBox;
            MaskedTextBox time = this.Controls.Find(Constants.time + session.activeIndex, true).FirstOrDefault() as MaskedTextBox;
            CheckBox billable = this.Controls.Find(Constants.isBillable + session.activeIndex, true).FirstOrDefault() as CheckBox;
            CheckBox submitted = this.Controls.Find(Constants.isSubmitted + session.activeIndex, true).FirstOrDefault() as CheckBox;
            CheckBox externalComment = this.Controls.Find(Constants.isExternalComment + session.activeIndex, true).FirstOrDefault() as CheckBox;
            TimeItem timeItem = new TimeItem();

            timeItem.description = description.Text;
            timeItem.isBillable = billable.Checked;
            timeItem.time = time.Text;
            timeItem.title = title.Text;
            timeItem.isCRMSubmitted = submitted.Checked;
            timeItem.isExternalComment = externalComment.Checked;

            return timeItem;
        }

        /// <summary>
        /// Returns a timeitem list from all controls
        /// </summary>
        /// <returns></returns>
        private List<TimeItem> MapTimeItems()
        {
            TextBox title = new TextBox();
            TextBox description = new TextBox();
            MaskedTextBox time = new MaskedTextBox();
            CheckBox billable = new CheckBox();
            CheckBox submitted = new CheckBox();
            CheckBox externalComment = new CheckBox();
            List<TimeItem> timeItems = new List<TimeItem>();

            int i = 0;
            while (i < session.totalLines)
            {
                TimeItem timeItem = new TimeItem();
                title = this.Controls.Find(Constants.title + i, true).FirstOrDefault() as TextBox;
                description = this.Controls.Find(Constants.description + i, true).FirstOrDefault() as TextBox;
                time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as MaskedTextBox;
                billable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                externalComment = this.Controls.Find(Constants.isExternalComment + i, true).FirstOrDefault() as CheckBox;
                submitted = this.Controls.Find(Constants.isSubmitted + i, true).FirstOrDefault() as CheckBox;
                
                timeItem.description = description.Text;
                timeItem.isBillable = billable.Checked;
                timeItem.time = time.Text;
                timeItem.title = title.Text;
                timeItem.isCRMSubmitted = submitted.Checked;
                timeItem.isExternalComment = externalComment.Checked;

                if (crmGuids[i] != null)
                {
                    timeItem.crmTaskId = new Guid(crmGuids[i]);
                }

                if (crmExternalCommentGuids[i] != null)
                {
                    timeItem.crmExternalCommentId = new Guid(crmExternalCommentGuids[i]);
                }

                timeItem.index = i;

                timeItems.Add(timeItem);
                i++;
            }

            return timeItems;
        }

        private void CreateThorPDF()
        {
            iTextSharp.text.Document document = new iTextSharp.text.Document();
            string docName = "ThorDoc " + DateTime.Now.Month + "-" + DateTime.Now.Day + " " + DateTime.Now.Hour + DateTime.Now.Minute+DateTime.Now.Second+ ".pdf";
            PdfWriter.GetInstance(document, new FileStream(docName, FileMode.Create));
            document.Open();
            int i = 0;

            List<TimeItem> timeItemList = new List<TimeItem>();

            while (i < session.totalLines)
            {
                TextBox project = this.Controls.Find(Constants.project + i, true).FirstOrDefault() as TextBox;
                TextBox title = this.Controls.Find(Constants.title + i, true).FirstOrDefault() as TextBox;
                TextBox description = this.Controls.Find(Constants.description + i, true).FirstOrDefault() as TextBox;
                MaskedTextBox time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as MaskedTextBox;
                CheckBox billable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                CheckBox submitted = this.Controls.Find(Constants.isSubmitted + i, true).FirstOrDefault() as CheckBox;

                TimeItem timeItem = new TimeItem();
                timeItem.project = project.Text;
                timeItem.description = description.Text;
                timeItem.isBillable = billable.Checked;
                timeItem.time = time.Text;
                timeItem.title = title.Text;
                timeItem.isCRMSubmitted = submitted.Checked;

                bool ismatch = false;
                foreach (TimeItem ti in timeItemList)
                {
                    if (ti.project == timeItem.project && ti.isBillable == timeItem.isBillable)
                    {
                        ismatch = true;
                        ti.description = timeItem.title+": " + ti.description + timeItem.description;
                        decimal totalSeconds = ((decimal)(TimeSpan.Parse(ti.time).TotalSeconds) + ((decimal)TimeSpan.Parse(timeItem.time).TotalSeconds));
                        TimeSpan t = TimeSpan.FromSeconds((double)totalSeconds);
                        ti.time = string.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                    }
                }

                if (ismatch == false)
                {
                    timeItem.description = timeItem.title + ":" + timeItem.description;
                    timeItemList.Add(timeItem);
                }

                i++;

            }

            foreach (TimeItem ti in timeItemList)
            {
                iTextSharp.text.Paragraph h = new iTextSharp.text.Paragraph(ti.project, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 18, iTextSharp.text.Font.BOLD));
                string content = ti.description;
                if (content.Length > 240)
                {
                    content = content.Substring(0, 212) + "[FULL DETAILS IN TICKET]";
                }

                iTextSharp.text.Paragraph p = new iTextSharp.text.Paragraph(content, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.NORMAL));
                iTextSharp.text.Paragraph b = new iTextSharp.text.Paragraph();

                if (ti.isBillable == true)
                {
                    b = new iTextSharp.text.Paragraph("Billable Hours: " + TimeSpan.Parse(ti.time).TotalHours, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.NORMAL));
                }
                else
                {
                    b = new iTextSharp.text.Paragraph("Non-Billable Hours: " + TimeSpan.Parse(ti.time).TotalHours, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.NORMAL));
                }
                
                document.Add(h);
                document.Add(p);
                document.Add(b);
 
            }
              
            document.Close();
            System.Diagnostics.Process.Start(docName);

        }

        #endregion

        private void GenerateThorButton_Click(object sender, EventArgs e)
        {
            CreateThorPDF();
        }

        private void testConnection()
        {
            string organizationUrl = "https://csp-build.crm.dynamics.com";
            string resourceURL = "https://csp-build.api.crm.dynamics.com" + "/api/data/";
            string clientId = "c4e4407b-66d1-4452-9b05-db0a0ce9baef"; // Client Id
            string appKey = "Sy[?Mk106C2OvHXZ:Krytwj=_XN_KKlh"; //Client Secret

            //Create the Client credentials to pass for authentication
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
           

            //get the authentication parameters
            AuthenticationParameters authParam = AuthenticationParameters.CreateFromResourceUrlAsync(new Uri(resourceURL)).Result;

            //Generate the authentication context - this is the azure login url specific to the tenant
            string authority = authParam.Authority;

            //request token
            AuthenticationResult authenticationResult = new AuthenticationContext(authority).AcquireTokenAsync(organizationUrl, clientcred).Result;

            //get the token              
            string token = authenticationResult.AccessToken;

            Uri serviceUrl = new Uri(organizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=9.1");
            OrganizationWebProxyClient sdkService;
            Microsoft.Xrm.Sdk.IOrganizationService _orgService;
            
            sdkService = new OrganizationWebProxyClient(serviceUrl, false);
            sdkService.CallerId = new Guid("{A9D51F0B-B11B-E611-80E0-5065F38AA901}");
            sdkService.HeaderToken = token;

            _orgService = (Microsoft.Xrm.Sdk.IOrganizationService)sdkService != null ? (Microsoft.Xrm.Sdk.IOrganizationService)sdkService : null;

            
            Microsoft.Xrm.Sdk.Entity user = _orgService.Retrieve("systemuser", new Guid("{A9D51F0B-B11B-E611-80E0-5065F38AA901}"), new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
        }



    }
}
