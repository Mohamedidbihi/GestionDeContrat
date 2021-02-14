using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;


namespace word
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        //Find and Replace Method
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Creeate the Doc Method
        private void CreateWordDocument(object filename, object SaveAs)
        {
            
            
                Word.Application wordApp = new Word.Application();
                object missing = Missing.Value;
                Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();
                // find and replace

                this.FindAndReplace(wordApp, "<nom>", textnom.Text);
                this.FindAndReplace(wordApp, "<prenom>", textprenom.Text);
                this.FindAndReplace(wordApp, "<datenaissance>", dateTimePicker1.Value.Date);
                this.FindAndReplace(wordApp, "<lieunaissance>", textlieunaiss.Text);
                this.FindAndReplace(wordApp, "<sf>", comboBox1.Text);
                this.FindAndReplace(wordApp, "<Nationalite>", comboBox2.Text);
                this.FindAndReplace(wordApp, "<cin>", textCin.Text);
                this.FindAndReplace(wordApp, "<cnss>", textCnss.Text);
                this.FindAndReplace(wordApp, "<gsm>", textGsm.Text);
                this.FindAndReplace(wordApp, "<lieumission>", textlieumission.Text);
                this.FindAndReplace(wordApp, "<entrepriseutili>", textentrepriseutulisatrice.Text);
                this.FindAndReplace(wordApp, "<chantier>", textchantier.Text);
                this.FindAndReplace(wordApp, "<qualif>", textQualifi.Text);
                this.FindAndReplace(wordApp, "<datedemission>", dateTimePicker2.Value.Date);
                this.FindAndReplace(wordApp, "<Datesys>", DateTime.Now.ToShortDateString());
                //
                this.FindAndReplace(wordApp, "<immat>", numericUpDown2.Text);
                this.FindAndReplace(wordApp, "<vestepantalon>", numericUpDown11.Text);
                this.FindAndReplace(wordApp, "<chaussures>", numericUpDown10.Text);
                this.FindAndReplace(wordApp, "<Gilet>", numericUpDown9.Text);
                this.FindAndReplace(wordApp, "<Casque>", numericUpDown1.Text);
                this.FindAndReplace(wordApp, "<Gants> ", numericUpDown6.Text);
                this.FindAndReplace(wordApp, "<lunettes>", numericUpDown7.Text);
                this.FindAndReplace(wordApp, "<ceinture>", numericUpDown8.Text);
                this.FindAndReplace(wordApp, "<ncontrat>", numericUpDown3.Text);


                
                if (textSalairehoraire.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Salaire horaire brut>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Salaire horaire brut>", "Salaire horaire brut :" + textSalairehoraire.Text);
                }
                if (textsalairebrut.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Salaire brut>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Salaire brut>", "Salaire brut :" + textsalairebrut.Text);
                }
                if (textSalaieNet.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Salaire net>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Salaire net>", "Salaire net :" + textSalaieNet.Text);
                }
                if (textprimepanier.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Prime de panier>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Prime de panier>", "Prime de panier :" + textprimepanier.Text);
                }
                if (textindeminitelait.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Indemnité de lait>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Indemnité de lait>", "Indemnité de lait :" + textindeminitelait.Text);
                }
                if (textindeminitetransport.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Indemnité de transport>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Indemnité de transport>", "Indemnité de transport :" + textindeminitetransport.Text);
                }
                if (textprimedesalissure.Text == "")
                {
                    this.FindAndReplace(wordApp, "<Prime de salissure>", "");
                }
                else
                {
                    this.FindAndReplace(wordApp, "<Prime de salissure>", "Prime de salissure :" + textprimedesalissure.Text);
                }
            }
            else
            {
                MessageBox.Show("File not Found!");
            }
                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
           ref missing, ref missing, ref missing,
           ref missing, ref missing, ref missing,
           ref missing, ref missing, ref missing,
           ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
                
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (textnom.Text == "" || textprenom.Text == "" || textlieunaiss.Text == "" || textCin.Text == "" || textGsm.Text == "")
                {
                    MessageBox.Show("Veuillez Remplir Tous Les champs Demandés  ", "Attention  !!!!");


                }
                else
                {

                    CreateWordDocument(@"C:\Users\rh\Desktop\contrat\temp.docx", @"C:\Users\rh\Desktop\contrat\Contrat.docx");
                    MessageBox.Show("File Created  !!");

                }
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message); 
            }
       
            }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void textnom_TextChanged(object sender, EventArgs e)
        {

        }

        private void textprenom_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void textlieunaiss_TextChanged(object sender, EventArgs e)
        {

        }

        private void textCin_TextChanged(object sender, EventArgs e)
        {

        }

        private void textCnss_TextChanged(object sender, EventArgs e)
        {

        }

        private void textGsm_TextChanged(object sender, EventArgs e)
        {

        }

        private void textlieumission_TextChanged(object sender, EventArgs e)
        {

        }

        private void textentrepriseutulisatrice_TextChanged(object sender, EventArgs e)
        {

        }

        private void textchantier_TextChanged(object sender, EventArgs e)
        {

        }

        private void textQualifi_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void textsalairebrut_TextChanged(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void textprimepanier_TextChanged(object sender, EventArgs e)
        {

        }

        private void textindeminitelait_TextChanged(object sender, EventArgs e)
        {

        }

        private void textindeminitetransport_TextChanged(object sender, EventArgs e)
        {

        }

        private void textprimedesalissure_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void textSalaieNet_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void textSalairehoraire_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label23_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;
            numericUpDown3.Value = 0;
            numericUpDown6.Value = 0;
            numericUpDown7.Value = 0;
            numericUpDown8.Value = 0;
            numericUpDown9.Value = 0;
            numericUpDown10.Value = 0;
            numericUpDown11.Value = 0;
            textnom.Text = "";
            textprenom.Text = "";
            textGsm.Text = "";
            textindeminitelait.Text = "";
            textchantier.Text = "";
            textCin.Text = "";
            textCnss.Text = "";
            textlieumission.Text = "";
            textlieunaiss.Text = "";
            textQualifi.Text = "";
            textentrepriseutulisatrice.Text = "";
            textsalairebrut.Text = "";
            textSalairehoraire.Text = "";
            textSalaieNet.Text = "";
            textindeminitetransport.Text = "";
            textprimedesalissure.Text = "";
            textprimepanier.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
     
        }

        private void label29_Click_1(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void numericUpDown6_ValueChanged_1(object sender, EventArgs e)
        {

        }

        private void label25_Click_1(object sender, EventArgs e)
        {

        }
    }
}
