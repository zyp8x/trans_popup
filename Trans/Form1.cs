using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Web;
using System.Net;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;



namespace Trans
{
    public partial class Form1 : Form
    {
        
        //for rounded border form
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );
        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
        }



        SpeechSynthesizer voice = null;
        private void Form1_Load(object sender, EventArgs e)
        {
            
            this.TopMost = true;
            
            this.WindowState = FormWindowState.Maximized;

           // MessageBox.Show("Plz choose Zh voice");

            voice = new SpeechSynthesizer();
            foreach (var v in  voice.GetInstalledVoices())
            {
                comboBox1.Items.Add(v.VoiceInfo.Name);
            }
        }
        //button close
        private void label1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //for moveable without border
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int)m.Result == 0x1)
                        m.Result = (IntPtr)0x2;
                    return;
            }

            base.WndProc(ref m);
        }
        //API for Google translatator 
        public string TranslateText(string input)
        {
            string url = String.Format
            ("https://translate.googleapis.com/translate_a/single?client=gtx&sl={0}&tl={1}&dt=t&q={2}",
             "zh", "en", Uri.EscapeUriString(input));
            HttpClient httpClient = new HttpClient();
            string result = httpClient.GetStringAsync(url).Result;
            var jsonData = new JavaScriptSerializer().Deserialize<List<dynamic>>(result);
            var translationItems = jsonData[0];
            string translation = "";
            foreach (object item in translationItems)
            {
                IEnumerable translationLineObject = item as IEnumerable;
                IEnumerator translationLineString = translationLineObject.GetEnumerator();
                translationLineString.MoveNext();
                translation += string.Format(" {0}", Convert.ToString(translationLineString.Current));
            }
            if (translation.Length > 1) { translation = translation.Substring(1); };
            return translation;
        }

        //button translate
        private void button1_Click_1(object sender, EventArgs e)
        {
               textBox2.Text = TranslateText(textBox1.Text);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        //button tts zh
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                voice.SelectVoice(comboBox1.Text);
                try
                {
                    switch (cmb_voiceType.SelectedIndex)
                    {
                        case 0:
                            voice.SelectVoiceByHints(VoiceGender.Male);
                            break;
                        case 1:
                            voice.SelectVoiceByHints(VoiceGender.Female);
                            break;
                    }
                    
                    voice.SpeakAsync(textBox1.Text);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        //button tts en
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                voice.SelectVoice(comboBox1.Text);
                switch (cmb_voiceType.SelectedIndex)
                {
                    case 0:
                        voice.SelectVoiceByHints(VoiceGender.Male);
                        break;
                    case 1:
                        voice.SelectVoiceByHints(VoiceGender.Female);
                        break;
                    
                }
                voice.SpeakAsync(textBox2.Text);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            this.Hide();
            f2.ShowDialog(); // Shows Form2
        }
    }
}
