using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;

namespace WindowsFormsApp8
{
    public partial class SomeForm : Form
    {
        public SomeForm()
        {
            InitializeComponent();
        }

        string connectionString = @"Server=localhost; Database=policework; Uid = root; Pwd=qwerty787;";
        TextChanger formm = new TextChanger();

        private void SomeForm_Load(object sender, EventArgs e)
        {
            GridFill();
            //allProceedings.SetSelected(0, true);
            today.Text = DateTime.Now.ToString("dd.MM.yy");
            startDate.Text = "010112";
            endDate.Text = DateTime.Now.ToString("dd.MM.yy");

            comboBox1.SelectedIndex = 0;
            allProceedings.SelectedIndex = 0;
            GridFillCompany();
            if (checkedListBox1.Items.Count > 0) { checkedListBox1.SetSelected(0, true); };

            richTextBox8.SelectAll();
            richTextBox8.SelectionIndent += 8;
            richTextBox8.SelectionLength = 0;

        }

        //отобразить производства из базы данных
        void GridFill()
        {
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                string man = "";

                //if (comboBox1.SelectedIndex=3) { }
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        man = "Фєдосов Д.В.";
                        break;
                    case 1:
                        man = "Сидорчук Р.А.";
                        break;
                    case 2:
                        man = "Фєдосов М.В.";
                        break;
                    default:
                        Console.WriteLine("Вы нажали неизвестную букву");
                        break;
                }


                allProceedings.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT number FROM investigation WHERE name='" + man + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    allProceedings.Items.Add(dr["number"].ToString());
                }
            }
            GridFillBanks();
            GridFillOurPeople();
            GridFillPeopleInCompany();
            GridFillRegions();
        }

        //отобразить всю информацию из базы данных
        private void allProceedings_SelectedIndexChanged(object sender, EventArgs e)
        {

            GridFillCompany();
            if (checkedListBox1.Items.Count > 0) { checkedListBox1.SetSelected(0, true); };

            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM investigation WHERE number='" + allProceedings.SelectedItem.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    fabulaProceedings.Text = dr["caseQ"].ToString();
                    numbProceedings.Text = dr["number"].ToString();
                    label1qq.Text = dr["number"].ToString();
                    dateProceedings.Text = dr["date"].ToString();
                    factProceedings.Text = dr["fact"].ToString();
                    qualificationProceedings.Text = dr["qualification"].ToString();
                }
            }
            numbProceedings.ReadOnly = dateProceedings.ReadOnly = qualificationProceedings.ReadOnly = factProceedings.ReadOnly = fabulaProceedings.ReadOnly = true;
            numbProceedings.Enabled = dateProceedings.Enabled = qualificationProceedings.Enabled = factProceedings.Enabled = fabulaProceedings.Enabled = false;
        }


        ///////////////////////////////////////////////////выделить по банкам////////////////////////////////////////////////////////
        //////по банкам (редакция 07.10.2018)
        void GridFillBanks()
        {
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                bankNames.Items.Clear();
                listBox3.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT name FROM banks";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    bankNames.Items.Add(dr["name"].ToString());
                    listBox3.Items.Add(dr["name"].ToString());
                }
            }
        }

        private void bankNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            funcc();
        }

        void funcc()
        {
            listBox3.SelectedIndex = bankNames.SelectedIndex;
            bankname.Text = listBox3.SelectedItem.ToString();

            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM banks WHERE name='" + bankNames.SelectedItem.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    nameOfBank.Text = dr["name"].ToString();
                    mfo.Text = dr["mfo"].ToString();
                    city.Text = dr["city"].ToString();
                    firstAdress1.Text = dr["adressOne"].ToString();
                    richTextBox8.Text = dr["other"].ToString();
                    //richTextBox1.Text = dr["other"].ToString();
                    postcode.Text = dr["postcode"].ToString();
                    label12.Text = dr["work"].ToString();
                }

                //if (postcode.Text == "04112") {tabPage4.BackColor = Color.FromArgb(255, 0, 0); }
                if (label12.Text == "minus") { tabPage4.BackColor = Color.FromArgb(255, 176, 153); }
                else { tabPage4.BackColor = Color.FromArgb(84, 255, 178); }

                button42.Text = bankNames.Items.Count.ToString();
            }
        }




        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            bankNames.SelectedIndex = listBox3.SelectedIndex;
            bankname.Text = listBox3.SelectedItem.ToString();
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////выделить по служебным лицам////////////////////////////////////////////////////////
        void GridFillPeopleInCompany()
        {
            /// по людям служебным лицам
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                listBox5.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT secondname FROM peopleInCompany";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    listBox5.Items.Add(dr["secondname"].ToString());
                }
            }
        }

        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox5.Items.Count > 0)
            {

                label22.Text = listBox5.SelectedItem.ToString();

                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand cmd = mysqlCon.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM peopleincompany WHERE secondname='" + listBox5.SelectedItem.ToString() + "'";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        secondNameUA.Text = dr["secondname"].ToString();
                        firstNameUA.Text = dr["firstname"].ToString();
                        thirdNameUA.Text = dr["thirdname"].ToString();
                        itn.Text = dr["ipn"].ToString();
                    }
                }
                fromIPNtoDate();
                string full = fullName.Text = secondNameUA.Text + " " + firstNameUA.Text + " " + thirdNameUA.Text;
            }
        }
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////по нашим сотрудникам///////////////////////////////////////////////////
        void GridFillOurPeople()
        {
            /// по людям нашим сотрудникам
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                peopleNames.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT name FROM people";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    peopleNames.Items.Add(dr["name"].ToString());
                }
            }
        }


        private void peopleNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM people WHERE name='" + peopleNames.SelectedItem.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    label10.Text = dr["fullname"].ToString();
                    namefulll.Text = dr["shortname"].ToString();
                    position.Text = dr["position"].ToString();
                    label16.Text = dr["rank"].ToString();
                    smalwithnumber.Text = dr["smalwithnumber"].ToString();
                    shortname.Text = dr["newshortname"].ToString();
                }
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////про предприятиям///////////////////////////////////////////////////


        void GridFillCompany()
        {
            ///по предприятиям
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                checkedListBox1.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT name FROM company WHERE number = '" + allProceedings.SelectedItem.ToString() + "'";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    checkedListBox1.Items.Add(dr["name"].ToString());
                }
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkedListBox1.Items.Count > 0)
            {
                label20.Text = checkedListBox1.SelectedItem.ToString();

                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand cmd = mysqlCon.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM company WHERE name='" + checkedListBox1.SelectedItem.ToString() + "'";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        profiter1.Text = dr["name"].ToString();
                        profiter2.Text = dr["code"].ToString();
                        textBox8.Text = dr["adress"].ToString();
                        //fspdlong.Text = dr["fspd"].ToString();
                    }
                }
                label7.Text = profiter1.Text + " (код " + profiter2.Text + ")";
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////про регионам///////////////////////////////////////////////////

        void GridFillRegions()
        {
            ///по регионам
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                listBox12.Items.Clear();
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT name FROM regions";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    listBox12.Items.Add(dr["name"].ToString());
                }
            }
        }

        private void listBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox12.Items.Count > 0)
            {
                label8.Text = listBox12.SelectedItem.ToString();

                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand cmd = mysqlCon.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM regions WHERE name='" + listBox12.SelectedItem.ToString() + "'";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        label8.Text = dr["name"].ToString();
                        label13.Text = dr["city"].ToString();
                        label9.Text = dr["street"].ToString();
                        label19.Text = dr["house"].ToString();
                        label18.Text = dr["index"].ToString();
                        label21.Text = dr["site"].ToString();
                        textBox3.Text = dr["telnumb"].ToString();
                    }
                }
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        private void Doput_Click(object sender, EventArgs e)
        {

        }

        //удалить из базы данных
        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
            {
                mysqlCon.Open();
                MySqlCommand cmd = mysqlCon.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE FROM investigation WHERE number='" + allProceedings.SelectedItem.ToString() + "'";
                cmd.ExecuteNonQuery();

                GridFill();
                allProceedings.SetSelected(0, true);
            }
        }

        //перенести в другое окно
        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            numbProceedings.Text = dateProceedings.Text = qualificationProceedings.Text = factProceedings.Text = fabulaProceedings.Text = "";
            numbProceedings.ReadOnly = dateProceedings.ReadOnly = qualificationProceedings.ReadOnly = factProceedings.ReadOnly = fabulaProceedings.ReadOnly = false;
            numbProceedings.Enabled = dateProceedings.Enabled = qualificationProceedings.Enabled = factProceedings.Enabled = fabulaProceedings.Enabled = true;
            toolStripMenuItem1.Enabled = true;
        }

        //добавить в базу данных
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (numbProceedings.Enabled == true && numbProceedings.Text != "")
            {
                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand cmd = mysqlCon.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "INSERT INTO investigation (number, caseQ, date, qualification, fact) VALUES ('" + numbProceedings.Text + "', '" + fabulaProceedings.Text + "', '" + dateProceedings.Text + "', '" + qualificationProceedings.Text + "', '" + factProceedings.Text + "')";
                    cmd.ExecuteNonQuery();

                    GridFill();
                    allProceedings.SetSelected(0, true);
                }
                toolStripMenuItem1.Enabled = false;
            }
        }

        //обновить базу данных
        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (numbProceedings.Enabled == false && numbProceedings.Text != "")
            {
                using (MySqlConnection mysqlCon = new MySqlConnection(connectionString))
                {
                    mysqlCon.Open();
                    MySqlCommand cmd = mysqlCon.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "UPDATE investigation SET caseQ = 'Фабула', date = '99.00.00', qualification = 'квалификация', fact = 'уклонение' WHERE number = 9";
                    cmd.ExecuteNonQuery();

                    GridFill();
                    allProceedings.SetSelected(0, true);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            formm.trafficPolice(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, position, shortname, smalwithnumber, requests);
        }

        private void courtWithout_Click(object sender, EventArgs e)
        {
            formm.withoutVisiting(numbProceedings, shortname, position, qualificationProceedings, factProceedings, courtList, dateProceedings);
        }

        private void courtBack_Click(object sender, EventArgs e)
        {
            formm.courtBack(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, courtList, position, namefulll);
        }

        private void courtTitle_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(shortBanks.Text);
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label1qq.Text);
        }

        private void button33_Click(object sender, EventArgs e)
        {

        }

        private void button37_Click(object sender, EventArgs e)
        {

        }

        private void button38_Click(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {
            MethodsWithWord.goTo(@"http://www.fg.gov.ua/not-paying");
        }

        private void button40_Click(object sender, EventArgs e)
        {
            string mfo = maskedTextBox2.Text;
            MethodsWithWord.goTo("https://bank.gov.ua/control/uk/bankdict/search?name=&type=&region=&mfo=" + mfo + "&edrpou=&size=&group=&fromDate=&toDate=");
        }

        private void copy_Click_1(object sender, EventArgs e)
        {
            string addbank = nameOfBank.Text + " (МФО " + mfo.Text + ", " + city.Text + ", " + firstAdress1.Text + "), ";
            shortBanks.Text += addbank;
            string addbankshort = nameOfBank.Text + " (МФО " + mfo.Text + "), ";
            fullBanks.Text += addbankshort;
        }

        private void button41_Click(object sender, EventArgs e)
        {
            formm.requestAllBank(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, today, fabulaProceedings, fullBanks, shortBanks, enterprisesAll);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridFill();
            allProceedings.SelectedIndex = 0;
            peopleNames.SelectedIndex = comboBox1.SelectedIndex;
            GridFillCompany();
        }


        private void button1_Click_2(object sender, EventArgs e)
        {
            fullBanks.Text = "";
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            shortBanks.Text = "";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            enterprisesAll.Font = new Font("Microsoft San Serif", 8);
        }

        private void nameOfBank_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(nameOfBank.Text);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            fromIPNtoDate();
        }

        void fromIPNtoDate()
        {
            string a = itn.Text;

            string c;
            c = a;
            c = c.Substring(8, 1);
            //int q = Convert.ToInt32(c);

            /*
            int r = 6;
            if (q % 2 != 0)
                label2.Text = "М";
            else
                label2.Text = "Ж";*/

            string b = a.Remove(5, 5);
            DateTime d = Convert.ToDateTime("01.01.1900");
            int n = Convert.ToInt32(b);
            var hi = d.AddDays(n - 1).ToString("dd.MM.yyyy");
            birthDay.Text = hi;
            //Clipboard.SetText(", " + hi + " р.н., і.п.н. " + maskedTextBox1.Text + ";");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string full = fullName.Text = secondNameUA.Text + " " + firstNameUA.Text + " " + thirdNameUA.Text;
            Clipboard.SetText(full);
        }

        private void button45_Click(object sender, EventArgs e)
        {
            string XX = secondNameUA.Text + " " + firstNameUA.Text + " " + thirdNameUA.Text + ", " + birthDay.Text + " р.н., і.п.н. " + itn.Text + ";";
            requests.Text += XX + Environment.NewLine;
        }

        private void label7_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label7.Text);
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label6.Text);
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void pasteInForm_Click(object sender, EventArgs e)
        {
            fspdlong.Text = "";
            foreach (string s in checkedListBox1.CheckedItems) { fspdlong.Text += s + ", "; };
            string deleteLastTwo = fspdlong.Text.Remove(fspdlong.TextLength - 2); //
            fspdlong.Text = deleteLastTwo;
        }

        private void activateMain_Click(object sender, EventArgs e)
        {
            switch (mainList.SelectedItem)
            {
                case "Довідка 222":
                    formm.non_disclosureStatement(numbProceedings, fabulaProceedings, dateProceedings, qualificationProceedings);
                    break;
                case "Група прокурорів":
                    formm.prosecutorsGroup(fabulaProceedings, dateProceedings, numbProceedings, qualificationProceedings, today);
                    break;
                case "Доручення на проведення ДР":
                    formm.instructionForDR(numbProceedings, fabulaProceedings, dateProceedings, today);
                    break;
                case "Група слідчих":
                    formm.groupOfInvestigator(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, today);
                    break;
                case "Вказівки":

                    break;
                case "План розслідування":
                    formm.investigationPlan(numbProceedings, factProceedings, dateProceedings, today, fabulaProceedings);
                    break;
                case "Довідка велика":
                    formm.bigInquiry(fabulaProceedings, dateProceedings, numbProceedings, factProceedings, today);
                    break;
                case "Довідка маленька":
                    formm.shortInguiry(fabulaProceedings, dateProceedings, numbProceedings, factProceedings, qualificationProceedings);
                    break;
                case "Схема правопорушення":

                    break;
                default:
                    MessageBox.Show("exctption");
                    break;

            }
        }

        private void activateRequest_Click(object sender, EventArgs e)
        {
            switch (requestList.SelectedItem)
            {
                case "1-ДФ":

                    break;
                case "ДАЇ":
                    formm.trafficPolice(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, position, shortname, smalwithnumber, requests);
                    break;
                case "Освіта":
                    formm.educationRequest(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, position, shortname, smalwithnumber, requests);
                    break;
                case "Міграційна":
                    formm.requestMigration(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, requests, position, shortname, smalwithnumber);
                    break;
                case "Смерть-шлюб":
                    formm.marrigeRequest(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, position, shortname, smalwithnumber, requests);
                    break;
                case "Прикордонники":
                    formm.borderRequest(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, position, shortname, smalwithnumber, requests);
                    break;
                default:
                    MessageBox.Show("exctption");
                    break;
            }
        }

        private void ActivateCompany_Click(object sender, EventArgs e)
        {
            switch (ActivateCompanyList.SelectedItem)
            {
                case "Протокол огляду ФСПД":
                    formm.overviewFSPD(numbProceedings, profiter1, profiter2, today, startDate, endDate);
                    break;
                case "ПК/ПЗ (схема)":
                    formm.excelChange(profiter1, profiter2, startDate, endDate);
                    break;
                case "Постанова почеркознавча":
                    formm.handwritingExamination(numbProceedings, today, dateProceedings, factProceedings, qualificationProceedings, profiter1, profiter2, fabulaProceedings, fspdlong);
                    break;
                case "Огляд податкової звітності":
                    formm.databaseInspection(profiter1, profiter2, fspdlong, today, numbProceedings, startDate, endDate);
                    break;
                case "Признание документами":
                    formm.addDocumentSix(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, today, fspdlong, bankname, startDate, endDate);
                    break;
                case "Осмотр ТДРД":
                    formm.overviewSix(numbProceedings, today, fspdlong, bankname, startDate, endDate);
                    break;
                case "Осмотр ТДРД (Печерская ДПИ)":
                    formm.overviewPechersk(numbProceedings, today, fspdlong, bankname, startDate, endDate);
                    break;
                case "Признание документами (Печерская ДПИ)":
                    formm.addDocumentPechersk(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, today, fspdlong, bankname, startDate, endDate);
                    break;
                case "Конверт (підриємство)":
                    formm.companyEnvelope(numbProceedings, profiter1, textBox8);
                    break;
                default:
                    MessageBox.Show("exeption");
                    break;
            }
        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (listBox4.SelectedItem)
            {
                case "Ухвала":
                    Clipboard.SetText("Ухвала Печерського районного суду м. Києва від " + MethodsWithWord.ConvertDate(startDate) + " року по справі " + textBox1.Text + ", про ТДРД до документів що перебувають у володінні " + bankNames.SelectedItem.ToString());
                    break;
                case "Протокол":
                    Clipboard.SetText("Протокол тимчасового доступу до речей і документів від " + MethodsWithWord.ConvertDate(endDate) + " року");
                    break;
                case "Опис":
                    Clipboard.SetText("Опис речей та документів, які вилучаються на підставі ухвали слідчого судді Печерського районного суду м. Києва від " + MethodsWithWord.ConvertDate(startDate) + " року (справа №" + textBox1.Text + ")");
                    break;
                case "Документи":
                    Clipboard.SetText("Документи " + bankNames.SelectedItem.ToString() + ", вилучені на підставі ухвали Печерського районного суду м. Києва від " + MethodsWithWord.ConvertDate(startDate) + " року по справі " + textBox1.Text + " у відповідності до Опис речей та документів які вилучаються на підставі ухвали слідчого судді Печерського районного суду м. Києва від " + MethodsWithWord.ConvertDate(startDate) + " року (справа №" + textBox1.Text + ") ");
                    break;
                case "Проведено":
                    Clipboard.SetText(" проведено тимчасовий доступ до копій документів, які перебувають у володінні " + bankNames.SelectedItem.ToString() + ", у яких міститься інформація щодо руху грошових коштів по рахункам");
                    break;
                case "Огляд як документів":
                    Clipboard.SetText("Протокол огляду документів " + fspdlong.Text + " в " + bankname.Text + " від " + MethodsWithWord.fullDate(today.Text));
                    break;
                case "Постанова про приєднання":
                    Clipboard.SetText("Постанова про визнання і приєднання до кримінального провадження документів " + fspdlong.Text + " в " + bankname.Text + " від " + MethodsWithWord.fullDate(today.Text));
                    break;
                default:
                    MessageBox.Show("exctption");
                    break;
            }
        }

        // переменные для распечаток титулок, описей и последних страниц
        string first;
        string second;
        string third;

        private void ListCourtPrint_SelectedIndexChanged(object sender, EventArgs e)
        {


            switch (ListCourtPrint.SelectedItem)
            {
                case "Матеріали провадження":
                    first = "титулка матеріали провадження";
                    second = "опис матеріали провадження";
                    third = "остання матеріали провадження";
                    break;
                case "Тимчасовий доступ":
                    first = "титулка тимчасовий доступ";
                    second = "опис тимчасовий доступя";
                    third = "остання тимчасовий доступ";
                    break;
                case "Арешт":
                    first = "титулка арешт";
                    second = "опис арешт";
                    third = "остання арешт";
                    break;
                case "Обшук":
                    first = "титулка обшук";
                    second = "опис обшук";
                    third = "остання обшук";
                    break;
                case "Застава":
                    first = "титулка застава";
                    second = "опис застава";
                    third = "остання застава";
                    break;
                default:
                    MessageBox.Show("exctption");
                    break;
            }
        }

        private void title_Click(object sender, EventArgs e)
        {
            MessageBox.Show(first);
        }

        private void opus_Click(object sender, EventArgs e)
        {
            MessageBox.Show(second);
        }

        private void lastPage_Click(object sender, EventArgs e)
        {
            MessageBox.Show(third);
        }

        private void convertToSlash_Click(object sender, EventArgs e)
        {
            ChangedText.Text = "";
            int numberOfSymbols = fullText.Text.Length;
            for (int i = 0; i < numberOfSymbols * 1.5; i++) { ChangedText.Text += "_"; Clipboard.SetText(ChangedText.Text); };
        }

        private void button2_Click(object sender, EventArgs e)
        {
            switch (peopleRequest.SelectedItem)
            {
                case "Повестка":
                    formm.subpoena(numbProceedings, secondNameUA, firstNameUA, thirdNameUA, peopleAdress);
                    break;
                case "Судимость":
                    formm.conviction(numbProceedings, dateProceedings, birthDay, birthAdress, peopleAdress, secondNameUA, firstNameUA, thirdNameUA);
                    break;
                case "Доверенность":
                    formm.procuratory(numbProceedings, dateProceedings, secondNameUA, firstNameUA, thirdNameUA, itn);
                    break;
                case "Активи":
                    formm.active(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, secondNameUA, firstNameUA, thirdNameUA, birthDay, startDate, profiter1, profiter2, itn);
                    break;
                default:
                    Console.WriteLine("Вы нажали неизвестную букву");
                    break;
            }
        }

        private void listBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (listBox7.SelectedItem)
            {
                case "Протокол тимчасового доступу":
                    MethodsWithWord.PringSmth(@"C:\Users\Mahatma\Desktop\Бланки\1. Протокол банк.doc");
                    break;
                case "Опис тимчасовий доступ":
                    MethodsWithWord.PringSmth(@"C:\Users\Mahatma\Desktop\Бланки\2. Маленький опис.docx");
                    break;
                case "Повістка":
                    MethodsWithWord.PringSmth(@"C:\Users\Mahatma\Desktop\Бланки\3. Повістка.docx");
                    break;
                default:
                    Console.WriteLine("Вы нажали неизвестную букву");
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            switch (listBox8.SelectedItem)
            {
                case "Доручення":
                    formm.rescue(numbProceedings, dateProceedings, factProceedings, qualificationProceedings);
                    break;
                case "Нагадування":
                    formm.requestBank(numbProceedings, dateProceedings, factProceedings, qualificationProceedings, today);
                    break;
                case "Повторне":

                    break;
                default:
                    Console.WriteLine("Вы нажали неизвестную букву");
                    break;
            }
        }

        private void companyCheckButton_Click(object sender, EventArgs e)
        {
            switch (listBox9.SelectedItem)
            {
                case "Цінні папери":
                    MethodsWithWord.goTo(@"https://smida.gov.ua/db/participant/" + profiter2.Text);
                    break;
                case "Особа масової реєстрації":
                    MethodsWithWord.goTo(@"https://nomis.com.ua/ru/" + profiter2.Text);
                    break;
                case "Адреса масової реєстрації":
                    MethodsWithWord.goTo(@"https://clarity-project.info/edr/" + profiter2.Text);
                    break;
                case "Публічні фінанси":
                    MethodsWithWord.goTo(@"https://spending.gov.ua/spa/transactions/search");
                    Clipboard.SetText(profiter2.Text);
                    break;
                case "Судові рішення":
                    MethodsWithWord.goTo(@"http://www.reyestr.court.gov.ua/");
                    Clipboard.SetText(profiter2.Text);
                    break;
                case "Тендери":
                    MethodsWithWord.goTo(@"https://prozorro.gov.ua/tender/search?edrpou=" + profiter2.Text);
                    break;
                case "Адреси реєстрації":
                    string linkDeclaration = HttpUtility.UrlEncode(textBox8.Text);
                    MethodsWithWord.goTo("https://clarity-project.info/edrs/?search=" + linkDeclaration);
                    break;
                case "Участь у тендерах":
                    MethodsWithWord.goTo("https://clarity-project.info/tenders/?tenderer=" + profiter2.Text);
                    break;
                case "Замовник тендерів":
                    MethodsWithWord.goTo("https://clarity-project.info/tenders/?entity=" + profiter2.Text);
                    break;
                default:
                    MessageBox.Show("нет такого");
                    break;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            switch (listBox10.SelectedItem)
            {
                case "Миротворец":
                    string linkMirotvoretc = HttpUtility.UrlEncode(fullName.Text);
                    MethodsWithWord.goTo(@"https://myrotvorets.center/criminal/?cf%5Bname%5D=" + linkMirotvoretc + "&cf%5Bcountry%5D=&cf%5Baddress%5D=&cf%5Bphone%5D=&cf%5Bdesc%5D=");
                    break;
                case "Декларация":
                    string linkDeclaration = HttpUtility.UrlEncode(fullName.Text);
                    MethodsWithWord.goTo("https://public.nazk.gov.ua/search?page=1&q=" + linkDeclaration + "&declarationType=&declarationYear=&documentType=&dtStart=&dtEnd=&isRisk=");
                    break;
                case "Google":
                    string linkGoogle = HttpUtility.UrlEncode(fullName.Text);
                    MethodsWithWord.goTo("https://www.google.com/search?ei=iqHpXNmaLYrsrgTPtqGIBA&q=" + linkGoogle + "&oq=" + linkGoogle + "&gs_l=psy-ab.3..0l10.203896.212836..213027...14.0..0.164.2378.16j8......0....1..gws-wiz.....6..0i71j35i39j35i39i19j0i131j0i10i1j0i10j0i203j0i10i203j0i10i1i42j0i30j0i10i42j0i10i1i67i42j0i67j0i131i67.n1Ruz-WEmtc");
                    break;
                case "Участь у тендерах":
                    MethodsWithWord.goTo("https://clarity-project.info/tenders/?tenderer=" + itn.Text);
                    break;
                case "Замовник тендерів":
                    MethodsWithWord.goTo("https://clarity-project.info/tenders/?entity=" + itn.Text);
                    break;
                default:
                    MessageBox.Show("Человека пробито");
                    break;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            profiter1.Text += "«";
            profiter1.Focus();
            profiter1.SelectionStart = profiter1.Text.Length;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            profiter1.Text += "»";
            profiter1.Focus();
            profiter1.SelectionStart = profiter1.Text.Length;
        }

        private void profiter1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //string profiterONE = profiter1.Text.ToUpper();
            //profiter1.Text = profiterONE;
            //profiter1.SelectionStart = profiter1.Text.Length;
        }

        private void listBox11_Click(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            switch (listBox11.SelectedItem)
            {
                case "Конверт":
                    formm.bankEnvelope(numbProceedings, nameOfBank, firstAdress1, postcode, city);
                    break;
                case "Клопотання по одному банку":
                    MessageBox.Show("Klopotania po odnomy banky");
                    break;
                case "Запит":
                    formm.bankRequest(numbProceedings, factProceedings, qualificationProceedings, dateProceedings, nameOfBank, firstAdress1, postcode, city);
                    break;
                case "Протокол огляду банку":
                    formm.accountMovement(profiter1, profiter2, numbProceedings, startDate, endDate, today, bankname);
                    break;
                case "Аналітика банку":
                    formm.excelBankChange(profiter1, profiter2, startDate, endDate, bankname);
                    break;
                default:
                    MessageBox.Show("default");
                    break;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            formm.createBlankEnverope(numbProceedings);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            string linkCourt = HttpUtility.UrlEncode(Convert.ToString(courtList.SelectedItem));
            MethodsWithWord.goTo(@"https://grd.gov.ua/judges?q=" + linkCourt);
        }

        public void methodQWR()
        {
            API_class nameAPI = new API_class();
            string someShit = nameAPI.getInformationAPI(profiter2.Text);
            textBox8.Text = someShit;
            //label7.Text = someShit;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            methodQWR();
        }

        public void methodQWER()
        {
            API_class nameAPI = new API_class();
            string someShit = nameAPI.getInformationAPIname(profiter2.Text);
            profiter1.Text = someShit;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            methodQWER();
            label7.Text = profiter1.Text + " (код " + profiter2.Text + ")";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            switch (listBox6.SelectedItem)
            {
                case "Адвокат":
                    MethodsWithWord.goTo(@"http://erau.unba.org.ua/");
                    break;
                case "Суддя":
                    MessageBox.Show("hello");
                    break;
                default:
                    MessageBox.Show("Hello");
                    break;
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            MethodsWithWord.goTo(@"https://cab.vkksu.gov.ua/search.html#/");
        }

        private void button21_Click(object sender, EventArgs e)
        {
            MethodsWithWord.goTo(@"https://court.gov.ua/opendata/");
        }

        public void methodAPII()
        {
            API_class nameAPI = new API_class();
            string someShit = nameAPI.getInformationAPIshortName(profiter2.Text);
            profiter1.Text = someShit;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            methodAPII();
            Clipboard.SetText(profiter1.Text);
            label7.Text = profiter1.Text + " (код " + profiter2.Text + ")";
        }

        public void methodAPIII()
        {
            API_class nameAPI = new API_class();
            string someShit = nameAPI.getInformationAPIaddressF(profiter2.Text);
            textBox8.Text = someShit;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            methodAPIII();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            requests.Text = "Клопотання про тимчасовий доступ до речей і документів, відомості яких становлять банківську таємницю від 00.00.0000 щодо " + fspdlong.Text + " в "+ bankname.Text + "; ухвала слідчого судді від " + MethodsWithWord.fullDate(startDate.Text) + " про тимчасовий доступ до речей і документів щодо " + fspdlong.Text + " в " + bankname.Text + "; протокол тимчасового доступу до речей і документів від " + MethodsWithWord.fullDate(endDate.Text) + " та опис речей і документів від " + MethodsWithWord.fullDate(endDate.Text) + "; документи " + fspdlong.Text + ", що вказані в описі речей і документів від " + MethodsWithWord.fullDate(endDate.Text) + ", які вилучені в " + bankname.Text + "; протокол огляду від " + MethodsWithWord.fullDate(today.Text) +"; постанова про визнання і приєднання до кримінального провадження документів від " + MethodsWithWord.fullDate(today.Text) + ". Загальна кількість";
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button25_Click(object sender, EventArgs e)
        {
            MethodsWithWord.goTo("https://www.google.com/maps/search/" + textBox8.Text);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            fspdlong.Text = "";
        }

        private void button28_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }

        private void qualificationProceedings_Click(object sender, EventArgs e)
        {
            
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(qualificationProceedings.Text);
        }

        private void button30_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(factProceedings.Text);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            API_class nameAPI = new API_class();
            object someShit = nameAPI.getInformationDirector(profiter2.Text);

            MessageBox.Show(Convert.ToString(someShit));
        }

        private void button32_Click(object sender, EventArgs e)
        {
            //listBox2.Items.Add(profiter1.Text);
            if(!listBox2.Items.Contains(profiter1.Text)) listBox2.Items.Add(profiter1.Text);
            if(!listBox14.Items.Contains(profiter2.Text)) listBox14.Items.Add(profiter2.Text);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            //listBox13.Items.Add(profiter1.Text);
            if (!listBox13.Items.Contains(profiter1.Text)) listBox13.Items.Add(profiter1.Text);
            if (!listBox15.Items.Contains(profiter2.Text)) listBox15.Items.Add(profiter2.Text);
        }

        private void button33_Click_1(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox14.Items.Clear();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            listBox13.Items.Clear();
            listBox15.Items.Clear();
        }

        private void button35_Click(object sender, EventArgs e)
        {

            switch (listBox16.SelectedItem)
            {
                case "Конверт":
                    formm.regionEnvelope(numbProceedings, label8, label9, label19, label18, label13);
                    break;
                default:
                    MessageBox.Show("Hello");
                    break;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MethodsWithWord.goTo(label21.Text);
        }
    }
}
