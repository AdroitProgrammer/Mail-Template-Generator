using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Recognition;

namespace Mail_Template_Generator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private AutoCompleteStringCollection collection = new AutoCompleteStringCollection();
        private SpeechRecognitionEngine speechRecog = new SpeechRecognitionEngine();
        private List<string> CSV = new List<string>();
        private Dictionary<string, string> namesAndAddresses = new Dictionary<string, string>();
        private MailTemplate mailTemplate = new MailTemplate();
        private string filename;
        List<Bitmap> GeneratedImages = new List<Bitmap>();
        private const int width = 1700, height = 2200;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            filename = @"C:\Users\jb4817\Desktop\Mailing Templates for  Spring 2019\Adjunct Addresses Spring 2019 - Sheet1.csv";
            //filename = @"C:\Users\jb4817\Desktop\Mailing Templates for  Spring 2019\Fake Names.csv";
            listView1.SmallImageList = new ImageList();
            listView1.LargeImageList = new ImageList();
            Load_CSV(filename);
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            var collection = new AutoCompleteStringCollection();
            collection.AddRange(getNames());
            textBox1.AutoCompleteCustomSource = collection;
            loadSpeechRecog();
        }

        public void loadSpeechRecog()
        {
            speechRecog.SpeechRecognized += SpeechRecog_SpeechRecognized;
            speechRecog.SetInputToDefaultAudioDevice();
            Choices choices = new Choices(getNames());
            GrammarBuilder builder = new GrammarBuilder(choices);
            speechRecog.LoadGrammar(new Grammar(builder));
            speechRecog.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void SpeechRecog_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string fullName = "";
            foreach(RecognizedWordUnit name in e.Result.Words)
            {
                fullName += name.Text + " ";
            }
            string final = fullName.Substring(0, fullName.Length - 1);
            addToPending(final);
        }

        private string[] getNames()
        {
            var amount = CSV.Count();
            string[] names = new string[amount - 1];

            for (int i = 1; i < amount; i++)
            {
                names[i - 1] = CSV[i].Split(',')[0];
            }

            return names;
        }

        private void cSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();

            openFile.ShowDialog();

            CSV.Clear();

            if (openFile.FileName == "")
                return;


            Load_CSV(openFile.FileName);


        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedIndex == -1)
                return;

            var selectedItem = listBox1.Items[listBox1.SelectedIndex].ToString();

            foreach (KeyValuePair<string, string> kvp in namesAndAddresses)
            {
                if (kvp.Key == selectedItem)
                {
                    var image = MailTemplate.createTemplate(kvp.Key, kvp.Value);
                    tabControl1.SelectedTab.BackgroundImage = image;
                    break;
                }
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            if (listBox1.SelectedIndex == -1)
                return;

            var selectedItem = listBox1.Items[listBox1.SelectedIndex].ToString();
            addToPending(selectedItem);
        }

        private string[] addedItems
        {
            get
            {
                string[] items = new string[listView1.Items.Count];

                for (var i = 0; i < items.Length; i++)
                    items[i] = listView1.Items[i].ToString().Split('{')[1].Split('}')[0].ToLower();

                return items;
            }
        }

        private void addToPending( string selectedItem)
        {
            
            if (addedItems.Contains(selectedItem.ToLower()))
            {
                MessageBox.Show("Item already added to the list!","Duplicate Item",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            foreach (KeyValuePair<string, string> kvp in namesAndAddresses)
            {
                if (kvp.Key.ToLower() == selectedItem.ToLower())
                {
                    selectedItem = kvp.Key;
                    GeneratedImages.Add(MailTemplate.createTemplate(kvp.Key, kvp.Value));
                    mailTemplate.Add(GeneratedImages[GeneratedImages.Count - 1]);

                    var templates = mailTemplate.AllTemplates;

                    if (tabControl1.TabPages.Count < mailTemplate.Pages.Count)
                    {
                        tabControl1.TabPages.Add("Template " + (tabControl1.TabPages.Count + 1));
                        tabControl1.TabPages[tabControl1.TabPages.Count - 1].BackgroundImageLayout = ImageLayout.Stretch;
                        tabControl1.TabPages[tabControl1.TabPages.Count - 1].BackColor = Color.White;
                        tabControl1.SelectedIndex = (tabControl1.TabPages.Count - 1);
                    }

                    for (var k = 0; k < templates.Length; k++)
                    {
                        tabControl1.TabPages[k].BackgroundImage = templates[k];
                    }
                    break;
                }
            }

            listView1.SmallImageList.Images.Add(GeneratedImages[GeneratedImages.Count - 1]);
            listView1.LargeImageList.Images.Add(GeneratedImages[GeneratedImages.Count - 1]);
            listView1.Items.Add(selectedItem, GeneratedImages.Count - 1);

            for (var i = 0; i < listBox1.Items.Count; i++)
            {
                if (listBox1.Items[i].ToString().ToLower() == selectedItem.ToLower())
                {
                    listBox1.Items.RemoveAt(i);
                    break;
                }
            }

        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            tabControl1.SelectedTab.BackgroundImage = GeneratedImages[listView1.SelectedIndices[0]];
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            listBox1.Items.Add(listView1.Items[listView1.SelectedIndices[0]].Text);
            listView1.Items.RemoveAt(listView1.SelectedIndices[0]);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

                if (textBox1.Text.Length > 1)
                {

                    if (collection.Count > 0)
                        collection.Clear();

                    var text = textBox1.Text.ToLower();

                    List<string> suggestText = new List<string>();

                    for (var i = 0; i < listBox1.Items.Count; i++)
                    {
                        if (listBox1.Items[i].ToString().ToLower().Contains(text))
                        {
                            suggestText.Add(listBox1.Items[i].ToString());
                        }
                    }
                }
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                var selectedItem = textBox1.Text;
                addToPending(selectedItem);
                textBox1.Clear();
            }
        }

        private void groupToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void allTemplatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.ShowDialog();

            if (folderBrowser.SelectedPath == null)
                return;

            var savePath = folderBrowser.SelectedPath;

            var counter = 1;

            foreach (Bitmap b in mailTemplate.AllTemplates)
            {
                b.Save(savePath +  "\\Mail Template #" + counter + ".jpg");
                counter++;
            }

            MessageBox.Show("Template Exported!", "Completed Exportation",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void currentTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.ShowDialog();

            if (folderBrowser.SelectedPath == null)
                return;

            var savePath = folderBrowser.SelectedPath;

            var currentNumb = tabControl1.SelectedIndex;

            var bitmap = mailTemplate.Template(currentNumb);

            bitmap.Save(savePath + "\\Mail Template #" + (currentNumb + 1 ) + ".jpg");

            MessageBox.Show("Template Exported!", "Completed Exportation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void newUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewUser userForm = new NewUser(this);
            userForm.Show();
        }


        public void add_To_CSV(string name , string address)
        {
            using (StreamWriter sw = new StreamWriter(filename,true))
            {
                sw.WriteLine(name + "," + "\"" + address + "\"");
            }

            namesAndAddresses.Add(name, address);
            listBox1.Items.Add(name);
            addToPending(name);
        }

        private void Load_CSV(string filename)
        {
            using (StreamReader sr = new StreamReader(filename))
            {
                string line;
                int count = 0;

                while ((line = sr.ReadLine()) != null)
                {
                    CSV.Add(line);
                    if (count != 0)
                    {
                        var array = line.Split(',');
                        var name = array[0];
                        var address = array[1].Split('"')[1] + "," + array[2].Split('"')[0];
                        namesAndAddresses.Add(name, address);
                    }
                    count++;
                }
            }

            var names = getNames();

            listBox1.Items.Clear();

            for (int j = 0; j < names.Length; j++)
            {
                listBox1.Items.Add(names[j]);
            }
        }


    }

}
