using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp8
{
    class TextChanger
    {
        public void withoutVisiting(MaskedTextBox numbProceedings, Label shortname, Label position, TextBox qualificationProceedings, TextBox factProceedings, ListBox courtList, MaskedTextBox dateProceedings)
        {
            string TemplateFileName = Application.StartupPath + @"\source\10. Титулки, описи, заявления\3. Документы в суд\1. Заява без участі.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{courtList}", courtList.SelectedItem.ToString(), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{shortname}", shortname.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Заява без участі", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void overviewFSPD(MaskedTextBox numbProceedings, TextBox profiter1, TextBox profiter2, MaskedTextBox today, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string profiter = profiter1.Text + " (код " + profiter2.Text + ")";

            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\1.4 Протокол огляду по фспд.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", profiter, 26, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 3, wordDocument);
            
            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + profiter1.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void procuratory(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox secondNameUA, TextBox firstNameUA, TextBox thirdNameUA, MaskedTextBox itn)
        {
            string TemplateFileName = Application.StartupPath + @"\source\7. Допрос\01. Довіреність.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{ipn}", itn.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{secondname}", secondNameUA.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{firstname}", firstNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{thirdname}", thirdNameUA.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " Довіреність" + secondNameUA.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }
        
        public void accountMovement(TextBox profiter1, TextBox profiter2, MaskedTextBox numbProceedings, MaskedTextBox startDate, MaskedTextBox endDate, MaskedTextBox today, Label bankname)
        {
            string profiter = profiter1.Text + " (код " + profiter2.Text + ")";
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\1.1 Огляд руху коштів.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", profiter, 5, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", bankname.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text + " " + "Огляд руху", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }
      
        public void handwritingExamination(MaskedTextBox numbProceedings, MaskedTextBox today, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, TextBox profiter1, TextBox profiter2, TextBox fabulaProceedings, RichTextBox fspdlong)
        {
            var fabula = fabulaProceedings.Text;
            var fspd0 = fspdlong.Text;
            var fspd1 = fspdlong.Text;
            var fspd2 = fspdlong.Text;

            string profiter = profiter1.Text + " (код " + profiter2.Text + ")";
            
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\1.3 Почеркознавча.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", profiter, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[4] { fabula, fspd0, fspd1, fspd2 };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Почеркознавча", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void rescue(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings)
        {
            string TemplateFileName = Application.StartupPath + @"\source\8. Оперативники\Доручення.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " Доручення", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void requestBank(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox today)
        {
            string TemplateFileName = Application.StartupPath + @"\source\8. Оперативники\Нагадування на доручення.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " Нагадування на доручення", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }
        
        public void requestAllBank(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox today, TextBox fabulaProceedings, RichTextBox fullBanks, RichTextBox shortBanks, RichTextBox enterprisesAll)
        {
            var fabula = fabulaProceedings.Text;
            var allbanks1 = fullBanks.Text;
            var allbanks2 = shortBanks.Text;
            var allfirms = enterprisesAll.Text;

            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\Багато банків .dotx";

            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 4, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);

            
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data1 = new string[4] {allbanks1, allbanks2, allfirms, fabula };

            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data1[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Багато банків", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void databaseInspection(TextBox profiter1, TextBox profiter2, RichTextBox fspdlong, MaskedTextBox today, MaskedTextBox numbProceedings, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string profiter = profiter1.Text + " (" + profiter2.Text + ")";
            var fspd = fspdlong.Text;
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\1.2 Огляд баз.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", profiter, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 2, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fspd };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text + " " + "Огляд баз", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        
        public void bankRequest(MaskedTextBox numbProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox dateProceedings, Label nameOfBank, Label firstAdress1, Label postcode, Label city)
        {
            string TemplateFileName = Application.StartupPath + @"\source\Запрос на банк.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", nameOfBank.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", firstAdress1.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{postcode}", postcode.Text, 1, wordDocument);                       
            MethodsWithWord.ReplaceSomeTextNew("{city}", city.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит " + nameOfBank.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void bankEnvelope(MaskedTextBox numbProceedings, Label nameOfBank, Label firstAdress1, Label postcode, Label city)
        {
            string TemplateFileName = Application.StartupPath + @"\source\Конверт.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{name}", nameOfBank.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", firstAdress1.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{postcode}", postcode.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{city}", city.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Конверт " + nameOfBank.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void regionEnvelope(MaskedTextBox numbProceedings, Label label8, Label label9, Label label19, Label label18, Label label13)
        {
            string TemplateFileName = Application.StartupPath + @"\source\Конверт (регион).dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{name}", label8.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", label9.Text + ", " + label19.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{postcode}", label18.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{city}", label13.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Конверт регион", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }




        public void investigationPlan(MaskedTextBox numbProceedings, TextBox factProceedings, MaskedTextBox dateProceedings, MaskedTextBox today, TextBox fabulaProceedings)
        {
            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\2. Отчет по производству\План слідчих дій.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.fullDate(dateProceedings.Text), 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.fullDate(today.Text), 1, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text + " " + "План слідчих дій", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void bigInquiry(TextBox fabulaProceedings, MaskedTextBox dateProceedings, MaskedTextBox numbProceedings, TextBox factProceedings, MaskedTextBox today)
        {
            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\2. Отчет по производству\Велика довідка.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.fullDate(dateProceedings.Text), 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text + " " + "велика", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void shortInguiry(TextBox fabulaProceedings, MaskedTextBox dateProceedings, MaskedTextBox numbProceedings, TextBox factProceedings, TextBox qualificationProceedings)
        {
            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\2. Отчет по производству\Стисла довідка.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text + " " + "стисла", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void prosecutorsGroup(TextBox fabulaProceedings, MaskedTextBox dateProceedings, MaskedTextBox numbProceedings, TextBox qualificationProceedings, MaskedTextBox today)
        {
            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\1. Первоначальные документы\4. Група прокурорів.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Група прокурорів", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }
       
        public void instructionForDR(MaskedTextBox numbProceedings, TextBox fabulaProceedings, MaskedTextBox dateProceedings, MaskedTextBox today)
        {
            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\1. Первоначальные документы\1. Доручення про проведення др.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Доручення про проведення ДР", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void non_disclosureStatement(MaskedTextBox numbProceedings, TextBox fabulaProceedings, MaskedTextBox dateProceedings, TextBox qualificationProceedings)
        {
            string fullDateResult = MethodsWithWord.fullDate(dateProceedings.Text);

            var fabula = fabulaProceedings.Text;
            string TemplateFileName = Application.StartupPath + @"\source\1. Первоначальные документы\2. Довідка про нерозголошення.dotx";

            //часть с закладкой
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            //часть с заменой
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 4, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{fulldate}", fullDateResult, 2, wordDocument);

            //продолжение части с закладкой
            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "222", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void requestMigration(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, RichTextBox requests, Label position, Label shortname, Label smalwithnumber)
        {
            var fabula = requests.Text;

            string TemplateFileName = Application.StartupPath + @"\source\4. Запросы\03. Міграційна.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{newshortname}", shortname.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{smalwithnumber}", smalwithnumber.Text, 1, wordDocument);

            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит міграційна", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void groupOfInvestigator(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox today)
        {
            string TemplateFileName = Application.StartupPath + @"\source\1. Первоначальные документы\3. Група слідчих.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);

            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Група слідчих", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        
        public void courtBack(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, ListBox courtList, Label position, Label namefulll)
        {
            string TemplateFileName = Application.StartupPath + @"\source\10. Титулки, описи, заявления\3. Документы в суд\2. Заява повернути.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{courtList}", courtList.SelectedItem.ToString(), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{shortname}", namefulll.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Заява повернути", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void overviewSix(MaskedTextBox numbProceedings, MaskedTextBox today, RichTextBox fspdlong, Label bankname, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\Для суду\1. Огляд баз ШСВ.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", fspdlong.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", bankname.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 2, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Огляд банку ШСВ " + " " + bankname.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void overviewPechersk(MaskedTextBox numbProceedings, MaskedTextBox today, RichTextBox fspdlong, Label bankname, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\Для суду\3. Огляд баз Печерський.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", fspdlong.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", bankname.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 2, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Огляд банку Печерськ " + " " + bankname.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }


        public void addDocumentSix(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox today, RichTextBox fspdlong, Label bankname, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\Для суду\2. Визнання документами ШСВ.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 4, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{shorttoday}", MethodsWithWord.fullDate(today.Text), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", fspdlong.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", bankname.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 4, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Визнання документами ШСВ " + " " + bankname.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void addDocumentPechersk(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, MaskedTextBox today, RichTextBox fspdlong, Label bankname, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string TemplateFileName = Application.StartupPath + @"\source\5. ТДРД\Для суду\4. Визнання документами Печерський.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 4, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{today}", MethodsWithWord.ConvertDate(today), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{shorttoday}", MethodsWithWord.fullDate(today.Text), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", fspdlong.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{name}", bankname.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{endDate}", MethodsWithWord.fullDate(endDate.Text), 4, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Визнання документами Печерськ " + " " + bankname.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }


        public void trafficPolice(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, Label position, Label shortname, Label smalwithnumber, RichTextBox requests)
        {
            var fabula = requests.Text;

            string TemplateFileName = Application.StartupPath + @"\source\4. Запросы\06. ДАЇ.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{newshortname}", shortname.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{smalwithnumber}", smalwithnumber.Text, 1, wordDocument);

            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }


            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит ДАЇ", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void educationRequest(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings,Label position, Label shortname, Label smalwithnumber, RichTextBox requests)
        {
            var fabula = requests.Text;

            string TemplateFileName = Application.StartupPath + @"\source\4. Запросы\02. Освіта.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{newshortname}", shortname.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{smalwithnumber}", smalwithnumber.Text, 1, wordDocument);

            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит освіта", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

       
        public void marrigeRequest(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, Label position, Label shortname, Label smalwithnumber, RichTextBox requests)
        {
            var fabula = requests.Text;

            string TemplateFileName = Application.StartupPath + @"\source\4. Запросы\05. Смерть-шлюб.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{newshortname}", shortname.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{smalwithnumber}", smalwithnumber.Text, 1, wordDocument);

            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит смерть-шлюб", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        
        public void borderRequest(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, Label position, Label shortname, Label smalwithnumber, RichTextBox requests)
        {
            var fabula = requests.Text;

            string TemplateFileName = Application.StartupPath + @"\source\4. Запросы\04. Прикордонники.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.ConvertDate(dateProceedings), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{position}", position.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{newshortname}", shortname.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{smalwithnumber}", smalwithnumber.Text, 1, wordDocument);

            Word.Bookmarks wBookmarks = wordDocument.Bookmarks;
            Word.Range wRange;
            int d = 0;
            string[] data = new string[1] { fabula };
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[d];
                d++;
            }

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Запит прикордонники", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void subpoena(MaskedTextBox numbProceedings, TextBox secondNameUA, TextBox firstNameUA, TextBox thirdNameUA, TextBox peopleAdress)
        {
            string TemplateFileName = Application.StartupPath + @"\source\7. Допрос\Повестка.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{firstname}", secondNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{secondname}", firstNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{thirdname}", thirdNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", peopleAdress.Text, 1, wordDocument);
            //MethodsWithWord.ReplaceSomeTextNew("{newshortname}", label19.Text, 1, wordDocument);
            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " Повістка " + secondNameUA.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void excelChange(TextBox profiter1, TextBox profiter2, MaskedTextBox startDate, MaskedTextBox endDate)
        {
            string NameExcel = Application.StartupPath + @"\source\3. АІС Податковий блок и др\Схема.xltm";
            var excelApp = new Excel.Application();
            Excel.Workbook ObjWorkBook = excelApp.Workbooks.Open(NameExcel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            
            excelApp.Visible = true;
            Excel.Worksheet X = excelApp.ActiveSheet as Excel.Worksheet;

            X.Range["h3"].Replace("{profiter}", profiter1.Text + " (код " + profiter2.Text + ")");
            X.Range["k1"].Replace("{startDate}", MethodsWithWord.fullDate(startDate.Text));
            X.Range["k1"].Replace("{endDate}", MethodsWithWord.fullDate(endDate.Text));

            X.SaveAs(Filename: Application.StartupPath + @"\result\" + " схема ПК-ПЗ " + profiter2.Text, FileFormat: Excel.XlFileFormat.xlExcel8);
        }

        public void createBlankEnverope (MaskedTextBox numbProceedings)
        {

            string TemplateFileName = Application.StartupPath + @"\source\Конверт (бланк).dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);
            
            wordDocument.SaveAs2(FileName: @"C:\Users\Mahatma\Desktop\" + numbProceedings.Text.Substring(13) + " Конверт пустий", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void companyEnvelope(MaskedTextBox numbProceedings, TextBox profiter1, TextBox textBox8)
        {
            string TemplateFileName = Application.StartupPath + @"\source\Конверт (підприємство).dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{company}", profiter1.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", textBox8.Text, 1, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " " + "Конверт ", FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void excelBankChange(TextBox profiter1, TextBox profiter2, MaskedTextBox startDate, MaskedTextBox endDate, Label bankname)
        {
            string NameExcel = Application.StartupPath + @"\source\5. ТДРД\Аналіз банку (шаблон).xltx";
            var excelApp = new Excel.Application();
            Excel.Workbook ObjWorkBook = excelApp.Workbooks.Open(NameExcel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            excelApp.Visible = true;
            Excel.Worksheet X = excelApp.ActiveSheet as Excel.Worksheet;

            X.Range["a1"].Replace("{profiter}", profiter1.Text + " (код " + profiter2.Text + ")");
            X.Range["a1"].Replace("{startDate}", MethodsWithWord.fullDate(startDate.Text));
            X.Range["a1"].Replace("{endDate}", MethodsWithWord.fullDate(endDate.Text));
            X.Range["a1"].Replace("{name}", bankname.Text);
        }

        public void conviction (MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox birthDay, TextBox birthAdress, TextBox peopleAdress, TextBox secondNameUA, TextBox firstNameUA, TextBox thirdNameUA)
        {
            string TemplateFileName = Application.StartupPath + @"\source\8. Оперативники\Судимість.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.fullDate(dateProceedings.Text), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{birthDay}", birthDay.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{birthplace}", birthAdress.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{address}", peopleAdress.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{secondname}", secondNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{firstname}", firstNameUA.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{thirdname}", thirdNameUA.Text, 2, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " Судимість" + secondNameUA.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }

        public void active(MaskedTextBox numbProceedings, MaskedTextBox dateProceedings, TextBox factProceedings, TextBox qualificationProceedings, TextBox secondNameUA, TextBox firstNameUA, TextBox thirdNameUA, TextBox birthDay, MaskedTextBox startDate, TextBox profiter1, TextBox profiter2, MaskedTextBox itn)
        {
            string profiter = profiter1.Text + " (код " + profiter2.Text + ")";

            string TemplateFileName = Application.StartupPath + @"\source\7. Допрос\Розшук активів.dotx";
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            var wordDocument = wordApp.Documents.Open(TemplateFileName);

            MethodsWithWord.ReplaceSomeTextNew("{numbProceedings}", numbProceedings.Text, 2, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{dateProceedings}", MethodsWithWord.fullDate(dateProceedings.Text), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{factProceedings}", factProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{qualificationProceedings}", qualificationProceedings.Text, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{firstname}", firstNameUA.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{secondname}", secondNameUA.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{thirdname}", thirdNameUA.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{birthDay}", birthDay.Text, 3, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{startDate}", MethodsWithWord.fullDate(startDate.Text), 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{profiter}", profiter, 1, wordDocument);
            MethodsWithWord.ReplaceSomeTextNew("{ipn}", itn.Text, 3, wordDocument);

            wordApp.Visible = true;
            wordDocument.SaveAs2(FileName: Application.StartupPath + @"\result\" + numbProceedings.Text.Substring(13) + " НАУВРУА" + secondNameUA.Text, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
        }
    }
}
