using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Diagnostics;


namespace WindowsFormsApp8
{
    public static class MethodsWithWord
    {
        public static void ReplaceSomeTextNew(string changeFrom, string changeTo, int howMuch, Word.Document wordDocument)
        {
            for (int i = 0; i < howMuch; i++)
            {
                wordDocument.Content.Find.Execute(FindText: changeFrom, ReplaceWith: changeTo);
            }
        }

        public static string ConvertDate(MaskedTextBox nameOfForm)
        {

            string day = nameOfForm.Text.Substring(0, 2);
            string month = nameOfForm.Text.Substring(3, 2);
            string year = nameOfForm.Text.Substring(6, 2);

            string nameOfMonth = null;

            switch (month)
            {
                case "01": nameOfMonth = "січня"; break;
                case "02": nameOfMonth = "лютого"; break;
                case "03": nameOfMonth = "березня"; break;
                case "04": nameOfMonth = "квітня"; break;
                case "05": nameOfMonth = "травня"; break;
                case "06": nameOfMonth = "червня"; break;
                case "07": nameOfMonth = "липня"; break;
                case "08": nameOfMonth = "серпня"; break;
                case "09": nameOfMonth = "вересня"; break;
                case "10": nameOfMonth = "жовтня"; break;
                case "11": nameOfMonth = "листопада"; break;
                case "12": nameOfMonth = "грудня"; break;
                default: nameOfMonth = "<<<---НеВеРнО вВеДеНа ДаТа--->>>"; break;
            }

            string fullDate = day + " " + nameOfMonth + " " + "20" + year;

            return fullDate;

        }

        public static string fullDate(string innertext)
        {
            string result = innertext.Insert(6, "20");
            return result;
        }

        public static void goTo (string site)
        {
            System.Diagnostics.Process.Start(site);
        }

        public static void PringSmth (string link)
        {
            ProcessStartInfo info = new ProcessStartInfo(link);
            info.Verb = "Print";
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(info);
            System.Threading.Thread.Sleep(50);
        }

        public static void TestMessage()
        {
            MessageBox.Show("heeo");
        }

    }
}
