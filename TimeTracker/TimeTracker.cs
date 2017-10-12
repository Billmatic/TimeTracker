using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using TimeTracker.Adaptors;
using TimeTracker.Common;
using TimeTracker.Entity;

namespace TimeTracker
{
    public partial class TimeTracker : Form
    {
        #region Private Variables

        Session session;
        CRMAdaptor crmAdaptor;
        string[] crmGuids = new string[30];

        /// <summary>
        /// The timer
        /// </summary>
        private Timer timer1 = new System.Windows.Forms.Timer();

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
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }

        }

        protected void CreateCaseTask(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            List<TimeItem> workItems = MapTimeItems();
            workItems = crmAdaptor.CreateCaseTask(workItems);

            foreach (TimeItem item in workItems)
            {
                if (item.crmTaskId != null)
                {
                    crmGuids[(int)item.index] = item.crmTaskId.ToString();
                    CheckBox submitted = this.Controls.Find(Constants.isSubmitted + item.index, true).FirstOrDefault() as CheckBox;
                    submitted.Checked = (bool)item.isCRMSubmitted;
                }              
            }     
            AutoSave();

            Cursor.Current = Cursors.Default;
        }

        protected void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Get file name.
            crmGuids = new string[30];
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

            Label titlelabel = new Label();
            titlelabel.Name = "titleLabel_" + session.index;
            titlelabel.Text = "Task Name";
            titlelabel.Location = new Point(3 + session.x, session.y);
            titlelabel.Font = new Font(titlelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(titlelabel);

            Label timelabel = new Label();
            timelabel.Name = "timeLabel_" + session.index;
            timelabel.Text = "Time";
            timelabel.Width = 60;
            timelabel.Location = new Point(220 + session.x, session.y);
            timelabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(timelabel);

            Label notelabel = new Label();
            notelabel.Name = "noteLabel_" + session.index;
            notelabel.Text = "Notes";
            notelabel.Location = new Point(303 + session.x, session.y);
            notelabel.Font = new Font(notelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(notelabel);

            Label isBillableLabel = new Label();
            isBillableLabel.Name = "BillableLabel_" + session.index;
            isBillableLabel.Text = "Bill";
            isBillableLabel.Width = 30;
            isBillableLabel.Location = new Point(738 + session.x, session.y);
            isBillableLabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(isBillableLabel);

            Label isCRMSubmittedLabel = new Label();
            isCRMSubmittedLabel.Name = "submittedLabel_" + session.index;
            isCRMSubmittedLabel.Text = "CRM";
            isCRMSubmittedLabel.Width = 50;
            isCRMSubmittedLabel.Location = new Point(770 + session.x, session.y);
            isCRMSubmittedLabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            timeTrackerToolTip.SetToolTip(isCRMSubmittedLabel, "Does not update to CRM when checked");
            MainPanel.Controls.Add(isCRMSubmittedLabel);

            session.y += 12;


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
                        case "title": //Display the text in each element.
                            TextBox titleText = this.Controls.Find(Constants.title + (session.index - 1), true).FirstOrDefault() as TextBox;
                            titleText.Text = reader.Value;
                            break;
                        case "time": //Display the end of the element.
                            TextBox timeText = this.Controls.Find(Constants.time + (session.index - 1), true).FirstOrDefault() as TextBox;
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
                        case "CRMSubmitted": //Display the end of the element.
                            CheckBox CRMSubmittedCheckBox = this.Controls.Find(Constants.isSubmitted + (session.index - 1), true).FirstOrDefault() as CheckBox;
                            CRMSubmittedCheckBox.Checked = Convert.ToBoolean(reader.Value);
                            break;
                        case "CRMGuid": //Display the end of the element.
                            crmGuids[session.index - 1] = reader.Value;
                            break;
                    }

                    CalcuateTotals();
                }
            }

            reader.Close();
            //Get file and set to controls
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

                TextBox DigiClockTextBox = this.Controls.Find(ControlName, true).FirstOrDefault() as TextBox;
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
                TextBox time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as TextBox;
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
            //Create Project Title Textbox
            TextBox titleTextBox = new TextBox();
            titleTextBox.Name = Constants.title + session.index;
            titleTextBox.Text = "";
            titleTextBox.Width = 120;
            titleTextBox.Location = new Point(3 + session.x, 14 + session.y);
            MainPanel.Controls.Add(titleTextBox);

            //This block dynamically creates a Button and adds it to the form
            Button btn = new Button();
            btn.Name = "btn_" + session.index;
            btn.Text = "Start";
            btn.Location = new Point(130 + session.x, 13 + session.y);
            btn.BackColor = System.Drawing.Color.White;
            btn.Click += new EventHandler(TimerStart);
            MainPanel.Controls.Add(btn);

            TextBox textBox = new TextBox();
            textBox.Name = Constants.time + session.index;
            textBox.Text = "00:00:00";
            textBox.Width = 60;
            textBox.Location = new Point(220 + session.x, 14 + session.y);
            MainPanel.Controls.Add(textBox);

            //Create Project Notes Textbox
            TextBox noteTextBox = new TextBox();
            noteTextBox.Name = Constants.description + session.index;
            noteTextBox.Text = "";
            noteTextBox.Width = 430;
            noteTextBox.Location = new Point(303 + session.x, 14 + session.y);
            MainPanel.Controls.Add(noteTextBox);

            //Create is Billable Checkbox
            CheckBox isBillableCheckBox = new CheckBox();
            isBillableCheckBox.Name = Constants.isBillable + session.index;
            isBillableCheckBox.Location = new Point(743 + session.x, 14 + session.y);
            isBillableCheckBox.Width = 30;
            isBillableCheckBox.Click += isBillableCheckBox_Click;
            MainPanel.Controls.Add(isBillableCheckBox);

            //Create is CRM Submitted Checkbox
            CheckBox isSubmittedCheckBox = new CheckBox();
            isSubmittedCheckBox.Name = Constants.isSubmitted + session.index;
            isSubmittedCheckBox.Location = new Point(785 + session.x, 14 + session.y);
            isSubmittedCheckBox.Width = 30;
            MainPanel.Controls.Add(isSubmittedCheckBox);

            session.y += 21;
            session.index++;
            session.totalLines++;

            if (session.isTimerActive)
            {
                DisableAllButtons();
                Button button = this.Controls.Find("btn_" + session.activeIndex, true).FirstOrDefault() as Button;
                button.Enabled = true;
                button.BackColor = System.Drawing.Color.Red;
            }
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
                TextBox title = this.Controls.Find(Constants.title + i, true).FirstOrDefault() as TextBox;
                TextBox description = this.Controls.Find(Constants.description + i, true).FirstOrDefault() as TextBox;
                TextBox time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as TextBox;
                CheckBox billable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                CheckBox submitted = this.Controls.Find(Constants.isSubmitted + i, true).FirstOrDefault() as CheckBox;

                xml += "<line>";
                xml += "<title>" + title.Text + "</title>";
                xml += "<description>" + description.Text + "</description>";
                xml += "<time>" + time.Text + "</time>";
                xml += "<billable>" + billable.Checked + "</billable>";
                xml += "<CRMSubmitted>" + submitted.Checked + "</CRMSubmitted>";
                if (crmGuids[i] != null)
                {
                    xml += "<CRMGuid>" + crmGuids[i].ToString() + "</CRMGuid>";
                }
                xml += "</line>";

                i++;
            }

            xml += "</projects>";
            return xml;
        }

        private void GetData()
        {
            Label titlelabel = new Label();
            titlelabel.Name = "titleLabel_" + session.index;
            titlelabel.Text = "Task Name";
            titlelabel.Location = new Point(3 + session.x, session.y);
            titlelabel.Font = new Font(titlelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(titlelabel);

            //Create Project Title Textbox
            TextBox titleTextBox = new TextBox();
            titleTextBox.Name = Constants.title + session.index;
            titleTextBox.Text = "";
            titleTextBox.Width = 120;
            titleTextBox.Location = new Point(3 + session.x, 24 + session.y);
            MainPanel.Controls.Add(titleTextBox);


            //This block dynamically creates a Button and adds it to the form
            Button btn = new Button();
            btn.Name = "btn_" + session.index;
            btn.Text = "Start";
            btn.Location = new Point(130 + session.x, 23 + session.y);
            btn.BackColor = System.Drawing.Color.White;
            btn.Click += new EventHandler(TimerStart);
            MainPanel.Controls.Add(btn);

            Label timelabel = new Label();
            timelabel.Name = "timeLabel_" + session.index;
            timelabel.Text = "Time";
            timelabel.Width = 60;
            timelabel.Location = new Point(220 + session.x, session.y);
            timelabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(timelabel);

            //Create Total Time Textbox
            TextBox textBox = new TextBox();
            textBox.Name = Constants.time + session.index;
            textBox.Text = "00:00:00";
            textBox.Width = 60;
            textBox.Location = new Point(220 + session.x, 24 + session.y);
            MainPanel.Controls.Add(textBox);

            Label notelabel = new Label();
            notelabel.Name = "noteLabel_" + session.index;
            notelabel.Text = "Notes";
            notelabel.Location = new Point(303 + session.x, session.y);
            notelabel.Font = new Font(notelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(notelabel);

            //Create Project Notes Textbox
            TextBox noteTextBox = new TextBox();
            noteTextBox.Name = Constants.description + session.index;
            noteTextBox.Text = "";
            noteTextBox.Width = 430;
            noteTextBox.Location = new Point(303 + session.x, 24 + session.y);
            MainPanel.Controls.Add(noteTextBox);

            Label isBillableLabel = new Label();
            isBillableLabel.Name = "BillableLabel_" + session.index;
            isBillableLabel.Text = "Bill";
            isBillableLabel.Width = 30;
            isBillableLabel.Location = new Point(738 + session.x, session.y);
            isBillableLabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            MainPanel.Controls.Add(isBillableLabel);

            //Create is Billable Checkbox
            CheckBox isBillableCheckBox = new CheckBox();
            isBillableCheckBox.Name = Constants.isBillable + session.index;
            isBillableCheckBox.Location = new Point(743 + session.x, 24 + session.y);
            isBillableCheckBox.Width = 30;
            isBillableCheckBox.Click += isBillableCheckBox_Click;
            MainPanel.Controls.Add(isBillableCheckBox);

            Label isCRMSubmittedLabel = new Label();
            isCRMSubmittedLabel.Name = "submittedLabel_" + session.index;
            isCRMSubmittedLabel.Text = "CRM";
            isCRMSubmittedLabel.Width = 50;
            isCRMSubmittedLabel.Location = new Point(770 + session.x, session.y);
            isCRMSubmittedLabel.Font = new Font(timelabel.Font, FontStyle.Bold);
            timeTrackerToolTip.SetToolTip(isCRMSubmittedLabel,"Does not update to CRM when checked");
            MainPanel.Controls.Add(isCRMSubmittedLabel);

            //Create is CRM Submitted Checkbox
            CheckBox isSubmittedCheckBox = new CheckBox();
            isSubmittedCheckBox.Name = Constants.isSubmitted + session.index;
            isSubmittedCheckBox.Location = new Point(785 + session.x, 24 + session.y);
            isSubmittedCheckBox.Width = 30;
            MainPanel.Controls.Add(isSubmittedCheckBox);

            session.y += 31;
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
        }

        /// <summary>
        /// Not Used
        /// </summary>
        /// <returns></returns>
        //private TimeItem MapTimeItem()
        //{
        //    TextBox title = this.Controls.Find(Constants.title + session.activeIndex, true).FirstOrDefault() as TextBox;
        //    TextBox description = this.Controls.Find(Constants.description + session.activeIndex, true).FirstOrDefault() as TextBox;
        //    TextBox time = this.Controls.Find(Constants.time + session.activeIndex, true).FirstOrDefault() as TextBox;
        //    CheckBox billable = this.Controls.Find(Constants.isBillable + session.activeIndex, true).FirstOrDefault() as CheckBox;
        //    CheckBox submitted = this.Controls.Find(Constants.isSubmitted + session.activeIndex, true).FirstOrDefault() as CheckBox;

        //    TimeItem timeItem = new TimeItem();

        //    timeItem.description = description.Text;
        //    timeItem.isBillable = billable.Checked;
        //    timeItem.time = time.Text;
        //    timeItem.title = title.Text;
        //    timeItem.isCRMSubmitted = submitted.Checked;

        //    return timeItem;
        //}

        private TimeItem BuildTimeItem()
        {
            TextBox title = this.Controls.Find(Constants.title + session.activeIndex, true).FirstOrDefault() as TextBox;
            TextBox description = this.Controls.Find(Constants.description + session.activeIndex, true).FirstOrDefault() as TextBox;
            TextBox time = this.Controls.Find(Constants.time + session.activeIndex, true).FirstOrDefault() as TextBox;
            CheckBox billable = this.Controls.Find(Constants.isBillable + session.activeIndex, true).FirstOrDefault() as CheckBox;
            CheckBox submitted = this.Controls.Find(Constants.isSubmitted + session.activeIndex, true).FirstOrDefault() as CheckBox;

            TimeItem timeItem = new TimeItem();

            timeItem.description = description.Text;
            timeItem.isBillable = billable.Checked;
            timeItem.time = time.Text;
            timeItem.title = title.Text;
            timeItem.isCRMSubmitted = submitted.Checked;

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
            TextBox time = new TextBox();
            CheckBox billable = new CheckBox();
            CheckBox submitted = new CheckBox();


            List<TimeItem> timeItems = new List<TimeItem>();

            int i = 0;
            while (i < session.totalLines)
            {
                TimeItem timeItem = new TimeItem();
                title = this.Controls.Find(Constants.title + i, true).FirstOrDefault() as TextBox;
                description = this.Controls.Find(Constants.description + i, true).FirstOrDefault() as TextBox;
                time = this.Controls.Find(Constants.time + i, true).FirstOrDefault() as TextBox;
                billable = this.Controls.Find(Constants.isBillable + i, true).FirstOrDefault() as CheckBox;
                submitted = this.Controls.Find(Constants.isSubmitted + i, true).FirstOrDefault() as CheckBox;

                timeItem.description = description.Text;
                timeItem.isBillable = billable.Checked;
                timeItem.time = time.Text;
                timeItem.title = title.Text;
                timeItem.isCRMSubmitted = submitted.Checked;

                if (crmGuids[i] != null)
                {
                    timeItem.crmTaskId = new Guid(crmGuids[i]);
                }
                
                timeItem.index = i;

                timeItems.Add(timeItem);
                i++;
            }

            return timeItems;
        }
 
        #endregion

    }
}
