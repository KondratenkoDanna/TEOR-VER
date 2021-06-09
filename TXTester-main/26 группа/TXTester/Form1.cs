using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using CheckBox = System.Windows.Forms.CheckBox;

namespace TXTester
{
    public partial class f_TXTexter : Form
    {
        public Question[] questions = new Question[256];
        public bool[] marks = new bool[256];
        public int questionNumber = 0;
        public int currentQuestion = 0;

        public f_TXTexter()
        {
            InitializeComponent();

            readTxt();

            Random rnd = new Random();
            //questions = questions.OrderBy(_ => rnd.Next());
            Array.Sort(questions, x => (rnd.Next() % 3) - 1);
        }

        private void b_Start_Click(object sender, EventArgs e)
        {
            
            tb_Title.Enabled = false;
            tb_Title.Visible = false;
            tb_Description.Enabled = false;
            tb_Description.Visible = false;
            b_Start.Enabled = false;
            b_Start.Visible = false;
            b_Next.Enabled = true;
            b_Next.Visible = true;

            if (currentQuestion == questionNumber - 1) b_Next.Text = "Finish";
            flp_QuestionPanel.Visible = true;

            Control[] questionElements = new Control[questions[0].answerCount + 2];
            questions[0].buildQuestion().Controls.CopyTo(questionElements, 0);
            flp_QuestionPanel.Controls.AddRange(questionElements);
        }

        public bool readTxt()
        {
            string[] toParse = File.ReadAllLines(@"resources\test.txt", System.Text.Encoding.Default);
            if (toParse[0].StartsWith("Title: "))
            {
                tb_Title.Text = toParse[0].Substring(7);
            }
            else return false;

            if (toParse[1].StartsWith("Discription: "))
            {
                tb_Description.Text = toParse[1].Substring(13);
            }
            else return false;

            int str = 2;
            while(str < toParse.Length)
            {
                char type = toParse[str][1];

                string pictureName;
                string questionText;
                if (toParse[str][3] == 'P')
                {
                    int index = 5;
                    while (toParse[str][index] != ')') index++;
                    pictureName = toParse[str].Substring(5, index - 5);
                    questionText = toParse[str].Substring(index + 2);
                }
                else
                {
                    pictureName = null;
                    questionText = toParse[str].Substring(4);
                }
                str++;
                int answerCount = 0;

                while (str < toParse.Length && (toParse[str][0] == '+' || toParse[str][0] == '-'))
                {
                    answerCount++;
                    str++;
                }

                bool[] answerMark = new bool[answerCount];
                string[] answerText = new string[answerCount];

                for(int i = 0; i < answerCount; ++i)
                {
                    if (toParse[str - answerCount + i][0] == '+') answerMark[i] = true;
                    else answerMark[i] = false;
                    answerText[i] = toParse[str - answerCount + i].Substring(2);
                }

                Question question = new Question(questionText, answerCount, type, answerMark, answerText, pictureName);
                questions[questionNumber] = question;
                questionNumber++;

            }

            for (int i = 0; i < questionNumber; ++i) marks[i] = false;
            return true;
        }

        private void b_Next_Click(object sender, EventArgs e)
        {
            bool isAlright = true;
            if (questions[currentQuestion].type == 'C')
            {
                List<CheckBox> checkBoxes = flp_QuestionPanel.Controls.OfType<CheckBox>().ToList();
                for (int i = 0; i < checkBoxes.Count; ++i)
                {
                    if (checkBoxes[i].Checked != questions[currentQuestion].answerMark[i]) isAlright = false;
                }
                if (isAlright) marks[currentQuestion] = true;
                else marks[currentQuestion] = false;
            } else
            {
                List<RadioButton> radioButtons = flp_QuestionPanel.Controls.OfType<RadioButton>().ToList();
                for (int i = 0; i < radioButtons.Count; ++i)
                {
                    if (radioButtons[i].Checked != questions[currentQuestion].answerMark[i]) isAlright = false;
                }
                if (isAlright) marks[currentQuestion] = true;
                else marks[currentQuestion] = false;
            }


            if (b_Next.Text == "Finish")
            {
                flp_QuestionPanel.Enabled = false;
                flp_QuestionPanel.Visible = false;
                b_Next.Enabled = false;
                b_Next.Visible = false;
                b_Previous.Enabled = false;
                b_Previous.Visible = false;
                tb_Title.Enabled = true;
                tb_Title.Visible = true;
                int score = 0;
                for (int i = 0; i < questionNumber; ++i) if (marks[i]) ++score;
                tb_Title.Text = "Score: " + score.ToString() + "/" + questionNumber.ToString();
                return;
            }

            currentQuestion++;
            b_Previous.Enabled = true;
            b_Previous.Visible = true;
            if (currentQuestion == questionNumber - 1) b_Next.Text = "Finish";


            flp_QuestionPanel.Controls.Clear();
            Control[] questionElements = new Control[questions[currentQuestion].answerCount + 2];
            questions[currentQuestion].buildQuestion().Controls.CopyTo(questionElements, 0);
            flp_QuestionPanel.Controls.AddRange(questionElements);

            
            // Do some "Next" things
        }

        private void b_Previous_Click(object sender, EventArgs e)
        {
            marks[currentQuestion] = false;
            currentQuestion--;
            if (b_Next.Text == "Finish") b_Next.Text = "Next";
            if (currentQuestion == 0)
            {
                b_Previous.Enabled = false;
                b_Previous.Visible = false;
            }

            flp_QuestionPanel.Controls.Clear();
            Control[] questionElements = new Control[questions[currentQuestion].answerCount + 2];
            questions[currentQuestion].buildQuestion().Controls.CopyTo(questionElements, 0);
            flp_QuestionPanel.Controls.AddRange(questionElements);

            // Do some "Previous" things
        }

        private Word.Document GetDoc(Word.Application app, string path) //Вот этот ГетДок
        {
            Word.Document oDoc = app.Documents.Add(path);
            SetTemplate(oDoc);
            return oDoc;
        }

        private void SetTemplate(Word.Document oDoc)
        {
            Random rnd = new Random();
            for (int i=1; i <= 5; ++i)
            {

                oDoc.Bookmarks["Test" + i.ToString()].Range.Text = tb_Title.Text;
                oDoc.Bookmarks["Description" + i.ToString()].Range.Text = tb_Description.Text;
                oDoc.Bookmarks["Variant_" + i.ToString()].Range.Text = "Вариант " + i.ToString();

                HashSet<int> q1 = new HashSet<int>();
                do
                {
                    q1.Add(rnd.Next(0, 15));
                } while (q1.Count < 4);

                int j = 1;
                foreach (int q in q1)
                {
                    Word.Range rng = oDoc.Bookmarks["Question" + i.ToString() + "_" + j.ToString()].Range;
                    rng.Text = questions[q].question;
                    if (questions[q].pictureName != null)
                    {
                        rng.InlineShapes.AddPicture(FileName: Environment.CurrentDirectory + "\\resources\\" + questions[q].pictureName);
                    }
                    //oDoc.Bookmarks["Question" + i.ToString() + "_" + j.ToString()].Range.Text = questions[q].question;
                    bool first = true;
                    string forAnswer = "";
                    for (int k=1; k<=4; ++k)
                    {
                        if (questions[q].answerMark[k - 1])
                        {
                            if (first) { forAnswer += "" + (char)(k - 1 + 'A'); first = false; }
                            else forAnswer += ", " + (char)(k - 1 + 'A');
                        }
                        oDoc.Bookmarks["Answer" + i.ToString() + "_" + j.ToString() + "_" + k.ToString()].Range.Text = questions[q].answerText[k - 1];
                    }
                    oDoc.Bookmarks["forAnswer" + i.ToString() + j.ToString()].Range.Text = forAnswer;
                    j++;
                }

                HashSet<int> q2 = new HashSet<int>();
                do
                {
                    q2.Add(rnd.Next(15, 25));
                } while (q2.Count < 4);
                
                j = 1;
                foreach (int q in q2)
                {
                    //Word.Bookmark bookmark = oDoc.Bookmarks["Question" + i.ToString() + "_" + (j + 4).ToString()];
                    Word.Range rng = oDoc.Bookmarks["Question" + i.ToString() + "_" + (j + 4).ToString()].Range;
                    rng.Text = questions[q].question;
                    
                    if (questions[q].pictureName != null) {
                        rng.InlineShapes.AddPicture(FileName: Environment.CurrentDirectory + "\\resources\\" + questions[q].pictureName);
                        //bookmark.Range.InlineShapes.AddPicture(FileName: Environment.CurrentDirectory + "\\resources\\" + questions[q].pictureName, true, true);
                        //Clipboard.SetImage(Image.FromFile(@"resources\\" + questions[q2[j - 1]].pictureName)); //Image.FromFile(@"resources\\" + pictureName);
                        //object oBookmark = "Question" + i.ToString() + "_" + (j + 4).ToString();
                        //oDoc.Bookmarks.get_Item(ref oBookmark).Range.InlineShapes.AddPicture(@"resources\\" + questions[q2[j - 1]].pictureName);
                        //Word.Application application = new Word.Application();
                        //application.ActiveDocument.Paragraphs[1].Range.Paste();
                        //oDoc.Bookmarks["Question" + i.ToString() + "_" + (j + 4).ToString()].Range = Clipboard.GetImage();
                    }
                    //oDoc.Bookmarks["Question" + i.ToString() + "_" + (j + 4).ToString()].Range.Expand(rng);
                    bool first = true;
                    string forAnswer = "";
                    for (int k = 1; k <= 4; ++k)
                    {
                        if (questions[q].answerMark[k - 1])
                        {
                            if (first) { forAnswer += "" + (char)(k - 1 + 'A'); first = false; }
                            else forAnswer += ", " + (char)(k - 1 + 'A');
                        }
                        oDoc.Bookmarks["Answer" + i.ToString() + "_" + (j + 4).ToString() + "_" + k.ToString()].Range.Text = questions[q].answerText[k - 1];
                    }
                    oDoc.Bookmarks["forAnswer" + i.ToString() + (j + 4).ToString()].Range.Text = forAnswer;
                    ++j;
                }
            }
            
            // если нужно заменять другие закладки, тогда копируем верхнюю строку изменяя на нужные параметры 

        }

        private void b_toDoc_Click(object sender, EventArgs e)
        {
            var app = new Word.Application(); 
            app.Visible = true;

            Word.Document oDoc = GetDoc(app, Environment.CurrentDirectory + "\\Shablon.dotx");
            oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\For_print.docx");
            oDoc.Close();
            app.Quit();
        }
    }

    public class Question
    {
        public string question;
        public int answerCount;
        public char type;
        public string pictureName;
        public bool[] answerMark;
        public string[] answerText;

        public Question(string question, int answerCount, char type, bool[] answerMark, string[] answerText, string pictureName = null)
        {
            this.question = question;
            this.answerCount = answerCount;
            this.type = type;
            this.pictureName = pictureName;
            this.answerMark = answerMark;
            this.answerText = answerText;
        }

        public FlowLayoutPanel buildQuestion()
        {
            FlowLayoutPanel flp_Result = new FlowLayoutPanel();

            if (pictureName != null)
            {
                PictureBox pb_Picture = new PictureBox();
                pb_Picture.Image = Image.FromFile(@"resources\\" + pictureName);
                pb_Picture.Width = 765;
                pb_Picture.Height = 200;
                pb_Picture.SizeMode = PictureBoxSizeMode.Zoom;
                flp_Result.Controls.Add(pb_Picture);
            }

            TextBox tb_Question = new TextBox();
            tb_Question.Width = 765;
            tb_Question.Text = question;
            tb_Question.ReadOnly = true;
            flp_Result.Controls.Add(tb_Question);

            for (int i = 0; i < answerCount; ++i)
            {
                if(type == 'C')
                {
                    CheckBox cb_Answer = new CheckBox();
                    cb_Answer.AutoSize = true;
                    cb_Answer.Text = answerText[i];
                    flp_Result.Controls.Add(cb_Answer);
                } else
                {
                    RadioButton rb_Answer = new RadioButton();
                    rb_Answer.AutoSize = true;
                    rb_Answer.Text = answerText[i];
                    flp_Result.Controls.Add(rb_Answer);
                }
            }
            
            return flp_Result;
        }
    }
}