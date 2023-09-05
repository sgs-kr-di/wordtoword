using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Ookii;
using BorderStyle = Spire.Doc.Documents.BorderStyle;
using HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment;
using Ookii.Dialogs.WinForms;
using Spire.Pdf;
using FileFormat = Spire.Doc.FileFormat;
using System.Text.RegularExpressions;

namespace wordtoword
{
    public partial class Form1 : Form
    {
        string[] filepaths;

        public Form1()
        {
            InitializeComponent();
        }

        class AutoClosingMessageBox
        {
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
            static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
            [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
            System.Threading.Timer _timeoutTimer;

            //쓰레드 타이머
            string _caption;

            //close 명령
            const int WM_CLOSE = 0x0010;

            AutoClosingMessageBox(string text, string caption, int timeout)
            {
                _caption = caption;
                _timeoutTimer = new System.Threading.Timer(OnTimerElapsed, null, timeout, System.Threading.Timeout.Infinite);
                MessageBox.Show(text, caption);
            }

            //생성자 함수
            public static void Show(string text, string caption, int timeout)
            {
                new AutoClosingMessageBox(text, caption, timeout);
            }

            //시간이 다되면 close 메세지를 보냄
            void OnTimerElapsed(object state)
            {
                IntPtr mbWnd = FindWindow(null, _caption);
                if (mbWnd != IntPtr.Zero) SendMessage(mbWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                _timeoutTimer.Dispose();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Integr";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);

                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴!

                    bool hasfooterimage = false;
                    foreach (Paragraph par in document.Sections[0].HeadersFooters.Footer.Paragraphs)
                    {
                        foreach (DocumentObject docObject in par.ChildObjects)
                        {
                            //If Type of Document Object is Picture, add it into image list
                            if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                hasfooterimage = true;

                            }
                        }
                    }


                    if (hasfooterimage == false)
                    {
                        insertImOnfooter(document);

                        removeBlankP(document);


                        //Details and styles
                        Table table = (Table)document.Sections[0].Tables[1];
                        table.Rows[3].Height = 20;
                        table.Rows[3].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;

                        foreach (Section sec in document.Sections)
                        {
                            Table tb = (Table)sec.Tables[1];
                            TextRange range = (tb.Rows[1].Cells[2].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range.CharacterFormat.FontSize = 9;
                            TextRange range2 = (tb.Rows[1].Cells[4].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range2.CharacterFormat.FontSize = 9;
                        }


                        merge_cells(document);
                        removeBorder(document, 1);

                        //saveDocument
                        document.SaveToFile(filepath, FileFormat.Docx);
                    }

                }
            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }

        }

        private void merge_cells(Document document)
        {

            ParagraphStyle list = new ParagraphStyle(document);
            list.Name = "list";
            list.CharacterFormat.FontName = "Arial";
            list.CharacterFormat.FontSize = 9.5F;

            document.Styles.Add(list);



            //PAGE7

            Table T7 = (Table)document.Sections[6].Tables[6];
            T7.ApplyHorizontalMerge(3, 2, 9);
            Paragraph ph = T7.Rows[3].Cells[2].Paragraphs[0];
            ph.AppendText("Adjusted Migration Result(s) (mg/kg)");
            ph.ApplyStyle(list.Name);
            ph.Format.HorizontalAlignment = HorizontalAlignment.Center;




            //PAGE10

            Table T10 = (Table)document.Sections[9].Tables[6];
            T10.ApplyHorizontalMerge(3, 2, 9);
            Paragraph ph2 = T10.Rows[3].Cells[2].Paragraphs[0];
            ph2.AppendText("Adjusted Migration Result(s) (mg/kg)");
            ph2.ApplyStyle(list.Name);
            ph2.Format.HorizontalAlignment = HorizontalAlignment.Center;

        }

        private void merge_cells2(Document document)
        {

            foreach (Section sec in document.Sections)
            {
                foreach (Table tb in sec.Tables)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        for (int j = 0; j < tb.Rows[i].Cells.Count; j++)
                        {
                            foreach (Paragraph pg in tb.Rows[i].Cells[j].Paragraphs)
                            {
                                if (pg.Text.Contains("Category III : Scrapped-off toy material"))
                                {

                                    sec.Tables[5].ApplyVerticalMerge(6, 0, 1);
                                    sec.Tables[5].ApplyVerticalMerge(7, 0, 1);
                                    sec.Tables[5].ApplyVerticalMerge(0, 0, 1);
                                    sec.Tables[5].ApplyHorizontalMerge(0, 1, 5);

                                }


                            }
                        }


                    }

                }
            }

        }

        private void insertImOnfooter(Document document)
        {
            foreach (Section section in document.Sections)
            {
                if (section.Tables.Count > 0)
                {
                    Table copiedtable = (Table)section.Tables[section.Tables.Count - 1];
                    //insert image on footer
                    for (int i = 0; i < copiedtable.Rows.Count; i++)
                    {
                        for (int j = 0; j < copiedtable.Rows[i].Cells.Count; j++)
                        {
                            //Foreach paragraph in the cell
                            foreach (Paragraph par in copiedtable.Rows[i].Cells[j].Paragraphs)
                            {
                                //Get each document object of paragraph items
                                foreach (DocumentObject docObject in par.ChildObjects)
                                {
                                    //If Type of Document Object is Picture, add it into image list
                                    if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                                    {
                                        //foreach (Section section in document.Sections)
                                        //{
                                        DocPicture picture = docObject as DocPicture;
                                        section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = false;
                                        Paragraph paragraph1 = section.HeadersFooters.Footer.AddParagraph();

                                        Image ima = picture.Image;
                                        paragraph1.AppendPicture(ima);

                                        //}

                                    }
                                }
                            }
                        }
                    }
                }
            }


            //footer로 쓰인 테이블 삭제
            for (int i = 0; i < document.Sections.Count; i++)
            {
                //모든 섹션의 마지막 테이블
                document.Sections[i].Tables.RemoveAt(document.Sections[i].Tables.Count - 1);
            }

        }

        private void insertImOnfoote_renew(Document document)
        {
            foreach (Section section in document.Sections)
            {
                if (section.Tables.Count > 0)
                {
                    Table copiedtable = (Table)section.Tables[section.Tables.Count - 1].Clone();
                    section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = false;
                    section.HeadersFooters.Footer.ChildObjects.Add(copiedtable);
                    Paragraph pr = new Paragraph(document);
                    pr.AppendText("\n");
                    section.HeadersFooters.Footer.ChildObjects.Add(pr);
                }
            }


            //footer로 쓰인 테이블 삭제
            for (int i = 0; i < document.Sections.Count; i++)
            {
                //모든 섹션의 마지막 테이블
                document.Sections[i].Tables.RemoveAt(document.Sections[i].Tables.Count - 1);
            }

        }

        private void insertHeader(Document document)
        {



            foreach (Section section in document.Sections)
            {



                section.PageSetup.HeaderDistance = 42.5197f;
                section.PageSetup.FooterDistance = 13.322835f;
                section.PageSetup.Margins.Top = 28.3465f;
                section.PageSetup.Margins.Bottom = 34.0157f;
                section.PageSetup.Margins.Left = 39.685f;
                section.PageSetup.Margins.Right = 30.33071f;


                List<DocumentObject> headerObjs = new List<DocumentObject>();
                Body textbody = section.Body;
                int stindex = textbody.GetIndex(section.Tables[0]);
                int endindex = textbody.GetIndex(section.Tables[1]);
                for (int i = stindex; i <= endindex; i++)
                {
                    DocumentObject headerObj = textbody.ChildObjects[i];
                    section.HeadersFooters.Header.ChildObjects.Add(headerObj.Clone());
                    headerObjs.Add(headerObj);


                }


                foreach (DocumentObject headerObj in headerObjs)
                {
                    textbody.ChildObjects.Remove(headerObj);

                }





            }

        }

        private Table GetNotes(Section sec)
        {



            foreach (Table tb in sec.Tables)
            {
                foreach (Paragraph pr in tb.Rows[0].Cells[0].Paragraphs)
                {
                    if (pr.Text.Trim().StartsWith("Note"))
                    {
                        return tb;

                    }



                }


            }


            return null;

        }


        private int IsTestRESULT(Section section)
        {

            for (int i = 0; i < section.Tables.Count; i++)
            {

                Table tb = section.Tables[i] as Table;
                if (tb.Rows.Count > 1)
                {
                    if (tb.Rows[1].Cells[0].Paragraphs.Count > 0)
                    {
                        if (tb.Rows[1].Cells[0].Paragraphs[0].Text.Trim() == "Test Item")
                        {
                            return i;


                        }
                    }
                }

            }
            return -1;



        }
        private int IsTin(Section section)
        {
            for (int i = 0; i < section.Tables.Count; i++)
            {

                Table tb = section.Tables[i] as Table;
                if (tb.Rows[0].Cells.Count > 1)
                {
                    if (tb.Rows[0].Cells[1].Paragraphs.Count > 0)
                    {
                        if (tb.Rows[0].Cells[1].Paragraphs[0].Text.Trim().ToString().Contains("Soluble Organic Tin"))
                        {
                            return i;


                        }
                    }
                }

            }
            return -1;


        }

        private static String GetCellText(TableCell cell)
        {
            String txt = null;
            foreach (Paragraph paragraph in cell.Paragraphs)
            {
                for (int j = 0; j < paragraph.ChildObjects.OfType<TextRange>().Count(); j++)
                {
                    TextRange textrange = paragraph.ChildObjects[j] as TextRange;
                    txt = txt + textrange.Text;
                }
            }
            return txt;
        }

        public static void MergeCell(Table table, bool isHorizontalMerge, int index, int start, int end)
        {
            if (isHorizontalMerge)
            {
                //Get a cell from table
                TableCell firstCell = table.Rows[index].Cells[start];
                //Invoke getCellText() method to get the cell’s text
                String firstCellText = GetCellText(firstCell);
                for (int i = start + 1; i <= end; i++)
                {
                    TableCell cell1 = table.Rows[index].Cells[i];
                    //Check if the text is the same as the first cell                
                    if (firstCellText == (GetCellText(cell1)))
                    {
                        //If yes, clear all the paragraphs in the cell
                        cell1.Paragraphs.Clear();
                    }
                }
                //Merge cells horizontally
                table.ApplyHorizontalMerge(index, start, end);
            }
            else
            {
                TableCell firstCell = table.Rows[start].Cells[index];
                String firstCellText = GetCellText(firstCell);
                for (int i = start + 1; i <= end; i++)
                {
                    TableCell cell1 = table.Rows[i].Cells[index];
                    if (firstCellText == (GetCellText(cell1)))
                    {
                        cell1.Paragraphs.Clear();
                    }
                }
                //Merge cells vertically
                table.ApplyVerticalMerge(index, start, end);
            }
        }

        private Table GetSampleDesc(Section sec)
        {



            foreach (Table tb in sec.Tables)
            {
                foreach (Paragraph pr in tb.Rows[0].Cells[0].Paragraphs)
                {
                    if (pr.Text.Trim().Contains("Sample Description"))
                    {
                        return tb;

                    }



                }


            }


            return null;

        }

        private void ConvertToParagraph(TableCell tbc, bool afterTB = false)
        {
            Table tb = tbc.OwnerRow.Owner as Table;
            Body textbody = tb.OwnerTextBody;
            int index = textbody.GetIndex(tb);
            TableRow tbr = tbc.OwnerRow;
            Paragraph pr = (Paragraph)tbc.Paragraphs[0].Clone();
            tb.Rows.Remove(tbr);
            if (afterTB == true)
            {
                index = +1;
            }

            textbody.ChildObjects.Insert(index, pr);

        }

        private void DrawSampleDescTable(Table tb, Document document)
        {
            int numberOFDesc = tb.Rows.Count;
            int numberOFRows = (int)Math.Ceiling((decimal)numberOFDesc / 2);
            int rowIndex = 0;

            Table sampleDescTb = new Table(document);
            sampleDescTb.ResetCells(numberOFRows, 2);
            sampleDescTb.SetColumnWidth(0, (float)241.5033, CellWidthType.Point);
            sampleDescTb.SetColumnWidth(1, (float)241.5033, CellWidthType.Point);



            for (int i = 0; i < numberOFDesc; i++)
            {
                if (i % 2 == 0)//짝수일 때
                {
                    Paragraph pr = (Paragraph)tb.Rows[i].Cells[0].Paragraphs[0].Clone();
                    pr.Format.HorizontalAlignment = HorizontalAlignment.Left;

                    sampleDescTb.Rows[rowIndex].Cells[0].Paragraphs.Add(pr);


                }
                else// 홀수 일 때
                {
                    Paragraph pr = (Paragraph)tb.Rows[i].Cells[0].Paragraphs[0].Clone();
                    pr.Format.HorizontalAlignment = HorizontalAlignment.Left;
                    sampleDescTb.Rows[rowIndex].Cells[1].Paragraphs.Add(pr);
                    rowIndex++;

                }



            }

            //remove original and insert new one
            Body textbody = tb.OwnerTextBody;
            int index = textbody.GetIndex(tb);

            textbody.Tables.Remove(tb);
            textbody.Tables.Insert(index, sampleDescTb);


        }

        private bool Body(Document document)
        {
            ListStyle listStyle = new ListStyle(document, ListType.Numbered);
            listStyle.Name = "levelstyle1";
            listStyle.Levels[0].CharacterFormat.FontName = "Arial";
            listStyle.Levels[0].CharacterFormat.FontSize = 10;
            listStyle.Levels[0].PatternType = ListPatternType.Arabic;

            document.ListStyles.Add(listStyle);

            bool tin = false;
            //result page P2부터 시작.
            foreach (Section sec in document.Sections)
            {
                int resultTbIdx = IsTestRESULT(sec); //이게 TEST RESULT 테이블이 있는 섹션인지 확인
                int tinTbIdx = IsTin(sec); //이게 TEST RESULT 테이블이 있는 섹션인지 확인

                if (resultTbIdx != -1)
                {

                    Table tbOnResult = sec.Tables[resultTbIdx] as Table;
                    int resultColICount = tbOnResult.Rows[1].Cells.Count - 4;

                    if (resultColICount == 1) //multi X 멀티가 아닌 파일
                    {

                        //table cell width 맞추기.

                        PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)99.2);
                        tbOnResult.PreferredWidth = width;
                        tbOnResult.SetColumnWidth(0, (float)34.5, CellWidthType.Percentage);
                        tbOnResult.SetColumnWidth(1, (float)35, CellWidthType.Percentage);
                        tbOnResult.SetColumnWidth(2, (float)11.8, CellWidthType.Percentage);
                        tbOnResult.SetColumnWidth(3, (float)18.3, CellWidthType.Percentage);



                    }
                    else //mullti O  멀티인 파일
                    {
                        //table cell width 맞추기.

                        PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)99.2);
                        tbOnResult.PreferredWidth = width;
                        tbOnResult.SetColumnWidth(0, (float)34.5, CellWidthType.Percentage);

                        float eachWidth = (35 / resultColICount);

                        for (int i = 1; i <= resultColICount; i++)
                        {
                            tbOnResult.SetColumnWidth(i, eachWidth, CellWidthType.Percentage);

                        }

                        tbOnResult.SetColumnWidth(resultColICount + 1, (float)11.8, CellWidthType.Percentage);
                        tbOnResult.SetColumnWidth(resultColICount + 2, (float)18.3, CellWidthType.Percentage);

                        //result 헤더 병합.
                        //tbOnResult.ApplyHorizontalMerge(1, 1, resultColIndex);
                        MergeCell(tbOnResult, true, 1, 1, resultColICount);


                    }


                    //result 뺴고 다른 컬럼 1,2 번 째 로우 병합.

                    tbOnResult.ApplyVerticalMerge(0, 1, 2);//Test Item
                    tbOnResult.ApplyVerticalMerge(resultColICount + 1, 1, 2);//Reporting Limit (mg/kg)
                    tbOnResult.ApplyVerticalMerge(resultColICount + 2, 1, 2); //Permissible Limit  EN71 - 3: 2019 + A1:2021(mg / kg)

                    tbOnResult.Rows[1].Height = 15.59055f;
                    tbOnResult.Rows[1].HeightType = TableRowHeightType.AtLeast;
                    tbOnResult.Rows[2].Height = 15.59055f;
                    tbOnResult.Rows[2].HeightType = TableRowHeightType.AtLeast;


                    //테이블 헤더 셀 중간 정렬.

                    tbOnResult.Rows[1].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;



                    //마지막 샘플 디스크립션 테이블 그냥 새로 만드는게 빠름.
                    Table sampleDescTb = GetSampleDesc(sec);
                    ConvertToParagraph(sampleDescTb.Rows[0].Cells[0]); //smapleDesc 1번째 로우에 있는거 그냥 파라그래프로 빼기.
                    DrawSampleDescTable(sampleDescTb, document);


                }

                if (tinTbIdx != -1)
                {

                    tin = true;
                    Table tbOnTin = sec.Tables[tinTbIdx] as Table;
                    int tinColICount = tbOnTin.Rows[0].Cells.Count - 2;
                    if (tinColICount > 1)
                    {
                        //table cell width 맞추기.

                        PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)83.2);
                        tbOnTin.PreferredWidth = width;

                        tbOnTin.SetColumnWidth(0, (float)38.1, CellWidthType.Percentage);
                        float eachWidth = ((float)(48.6 / tinColICount));

                        for (int i = 1; i <= tinColICount; i++)
                        {
                            tbOnTin.SetColumnWidth(i, eachWidth, CellWidthType.Percentage);

                        }

                        tbOnTin.SetColumnWidth(tinColICount + 1, (float)13.3, CellWidthType.Percentage);


                        //result 헤더 병합.
                        //tbOnResult.ApplyHorizontalMerge(1, 1, resultColIndex);
                        MergeCell(tbOnTin, true, 0, 1, tinColICount);







                    }
                    else
                    {

                        //table cell width 맞추기.


                        PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)83.2);
                        tbOnTin.PreferredWidth = width;
                        tbOnTin.SetColumnWidth(0, (float)38.1, CellWidthType.Percentage);
                        tbOnTin.SetColumnWidth(1, (float)48.6, CellWidthType.Percentage);
                        tbOnTin.SetColumnWidth(2, (float)13.3, CellWidthType.Percentage);



                    }

                    //result 뺴고 다른 컬럼 1,2 번 째 로우 병합.

                    tbOnTin.ApplyVerticalMerge(0, 0, 1);//Test Item
                    tbOnTin.ApplyVerticalMerge(tinColICount + 1, 0, 1);//MDL


                    tbOnTin.Rows[0].Height = 12.75591f;
                    tbOnTin.Rows[0].HeightType = TableRowHeightType.AtLeast;
                    tbOnTin.Rows[1].Height = 12.75591f;
                    tbOnTin.Rows[1].HeightType = TableRowHeightType.AtLeast;


                    //테이블 헤더 셀 중간 정렬.

                    tbOnTin.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    tbOnTin.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Left;


                    //마지막 샘플 디스크립션 테이블 그냥 새로 만드는게 빠름.
                    Table sampleDescTb = GetSampleDesc(sec);
                    ConvertToParagraph(sampleDescTb.Rows[0].Cells[0]); //smapleDesc 1번째 로우에 있는거 그냥 파라그래프로 빼기.
                    DrawSampleDescTable(sampleDescTb, document);





                }


                //notes 부분
                List<Paragraph> prs2 = new List<Paragraph>();
                Table noteTb = GetNotes(sec);

                if (noteTb != null)
                {

                    noteTb.SetColumnWidth(0, (float)29.76378, CellWidthType.Point);
                    noteTb.SetColumnWidth(1, (float)489.54331, CellWidthType.Point);

                    List<string> notes = noteTb.Rows[0].Cells[1].Paragraphs[0].Text.Split(new[] { "" }, StringSplitOptions.None).ToList();
                    foreach (string note in notes)
                    {
                        Paragraph newpr = new Paragraph(document);
                        TextRange tr = newpr.AppendText(note.Trim());

                        if (Regex.IsMatch(tr.Text.Replace(" ", ""), @"([0-9]\.)\w+") == true && tr.Text.Trim() != "")
                        {
                            int idx = tr.Text.IndexOf('.');
                            tr.Text = tr.Text.Substring(idx + 1);
                            tr.CharacterFormat.FontName = "Arial";
                            tr.CharacterFormat.FontSize = 10;


                            newpr.ListFormat.ApplyStyle("levelstyle1");
                            newpr.Format.HorizontalAlignment = HorizontalAlignment.Left;
                            prs2.Add(newpr);

                        }


                        // newpr.AppendBreak(BreakType.LineBreak);

                    }

                    //원래있던 파라그래프 지우고 새로 넣기.
                    noteTb.Rows[0].Cells[1].ChildObjects.Clear();
                    foreach (Paragraph pr in prs2)
                    {
                        noteTb.Rows[0].Cells[1].ChildObjects.Add(pr);



                    }



                }

            }


            return tin;


        }

        private void FooterForASTM_CHECMICAL(Document document)
        {
            foreach (Section section in document.Sections)
            {
                if (section.Tables.Count > 0)
                {
                    Table copiedtable = (Table)section.Tables[section.Tables.Count - 1].Clone();
                    section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = false;
                    section.HeadersFooters.Footer.ChildObjects.Add(copiedtable);
                    //Paragraph pr = new Paragraph(document);
                    //pr.AppendText("\n");
                    //section.HeadersFooters.Footer.ChildObjects.Add(pr);
                }
            }
            //footer로 쓰인 테이블 삭제
            for (int i = 0; i < document.Sections.Count; i++)
            {
                //모든 섹션의 마지막 테이블
                document.Sections[i].Tables.RemoveAt(document.Sections[i].Tables.Count - 1);
            }

        }

        private void BodyForASTM_CHECMICAL(Document document)
        {
            ListStyle listStyle = new ListStyle(document, ListType.Numbered);
            listStyle.Name = "levelstyle1";
            listStyle.Levels[0].CharacterFormat.FontName = "Arial";
            listStyle.Levels[0].CharacterFormat.FontSize = 10;
            listStyle.Levels[0].PatternType = ListPatternType.Arabic;
            document.ListStyles.Add(listStyle);

            //result page P2부터 시작.
            foreach (Section sec in document.Sections)
            {
                //MASS OF TRACE AMOUNT(MG)찾기
                foreach (Table table in sec.Tables)
                {
                    foreach (TableRow tbr in table.Rows)
                    {
                        foreach (TableCell tbrc in tbr.Cells)
                        {
                            foreach (Paragraph pr in tbrc.Paragraphs)
                            {
                                if (pr.Text.Trim() == "Mass of trace amount(mg)")
                                {
                                    int mergingstIdx = tbrc.GetCellIndex() + 1;
                                    int mergingedIdx = tbr.Cells.Count - 2;
                                    MergeCell(table, true, tbr.GetRowIndex(), mergingstIdx, mergingedIdx);
                                    Paragraph newpr = tbr.Cells[mergingstIdx].AddParagraph();
                                    newpr.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    TextRange newtr = newpr.AppendText("Adjusted Migration Result(s) (mg/kg)");
                                    newtr.CharacterFormat.FontName = "Arial";
                                    newtr.CharacterFormat.FontSize = 10;

                                }

                            }

                        }

                    }
                }

                //마지막 샘플 디스크립션 테이블 그냥 새로 만드는게 빠름.
                Table sampleDescTb = GetSampleDesc(sec);
                if (sampleDescTb != null)
                {
                    ConvertToParagraph(sampleDescTb.Rows[0].Cells[0]); //smapleDesc 1번째 로우에 있는거 그냥 파라그래프로 빼기.
                    DrawSampleDescTable(sampleDescTb, document);
                }
            }
        }

        private void removeBlankP(Document document)
        {

            //빈페이지 삭제하기 --마지막이 항상 테이블임. 마지막 테이블 찾아서 그 다음에 있는 빈 줄 다 삭제하기-
            foreach (Section sec in document.Sections)
            {
                int indexpflt = 0;
                for (int i = 0; i < sec.Body.ChildObjects.Count; i++)
                {
                    if (sec.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Table)
                    {
                        indexpflt = i;

                    }
                }


                for (int i = indexpflt + 1; i < sec.Body.ChildObjects.Count; i++)

                {
                    if (sec.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)

                    {
                        if (String.IsNullOrEmpty((sec.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                        {
                            sec.Body.ChildObjects.Remove(sec.Body.ChildObjects[i]);
                            i--;
                        }

                    }

                }

            }


        }

        private void removeBorder(Document document, int o)
        {
            //remove border
            if (o == 1)
            {
                Table tb = (Table)document.Sections[2].Tables[7];
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    TableRow tr = tb.Rows[i];

                    foreach (TableCell cell in tr.Cells)
                    {

                        cell.CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
                        cell.CellFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                        cell.CellFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                        //bottom만 none 안됨 sprire version issue. 그래서 그냥 다 지워버리고 다시 해야함..
                        //cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.None;
                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {

                            if (pr.Text.Trim() == "Clause")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                            }
                        }
                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {
                            if (pr.Text.Trim() == "4")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                            }
                        }

                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {

                            if (pr.Text.Trim() == "5")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                            }
                        }

                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {

                            if (pr.Text.Trim() == "7")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                            }
                        }

                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {

                            if (pr.Text.Trim() == "8")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                            }
                        }


                        //if (i == 0||i==1||i==16||i==18||i==20)
                        //{
                        //    cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;


                        if (i == tb.Rows.Count - 1)
                        {
                            cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                        }

                    }
                }



            }
            if (o == 2)
            {
                Table tb = (Table)document.Sections[2].Tables[5];
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    TableRow tr = tb.Rows[i];



                    foreach (TableCell cell in tr.Cells)
                    {
                        cell.CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.None;
                        cell.CellFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                        cell.CellFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Single;


                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {
                            if (pr.Text.Trim() == "Clause")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                            }
                        }
                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {
                            if (pr.Text.Trim() == "4")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;


                            }
                        }
                        foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                        {
                            if (pr.Text.Trim() == "5")
                            {
                                cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
                            }
                        }

                        //bottom만 none 안됨. 그래서 그냥 다 지워버리고 다시 해야함..
                        //cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.None;

                        //if (i == 0 || i == 1 || i == 6)
                        //{
                        //    cell.CellFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                        //}
                        if (i == tb.Rows.Count - 1)
                        {
                            cell.CellFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;

                        }



                    }


                }

            }

        }


        private void DetailsOnTestConductedforPhEN1(Document document, int sectionidex)
        {
            Table tb = GetTableByFirstCell(document, "Clause", sectionidex);
            string previousClause = "";
            if (tb != null)
            {

                //table height
                foreach (TableRow tbr in tb.Rows)
                {
                    tbr.Height = 15.59055f;
                }



                for (int i = 1; i < tb.Rows.Count; i++)
                {
                    TableRow tr = tb.Rows[i];
                    TableCell cell = tr.Cells[0];
                    Paragraph pr = tr.Cells[0].Paragraphs[0];

                    if (previousClause != pr.Text.Trim().Split('.').First())
                    {

                        SetTableBorderByTableRow(tr, BorderStyle.Single, 1);
                        SetTableBorderByTableRow(tr, BorderStyle.None, 0);
                        pr.Text = " " + pr.Text.Trim();
                        pr.Format.FirstLineIndent = 0F;
                    }
                    else
                    {

                        SetTableBorderByTableRow(tr, BorderStyle.None, 0);
                        SetTableBorderByTableRow(tr, BorderStyle.None, 1);
                        pr.Text = "   " + pr.Text.Trim();
                        pr.Format.FirstLineIndent = 0F;

                    }
                    previousClause = pr.Text.Trim().Split('.').First();
                }

                TableRow firstRow = tb.FirstRow;
                SetTableBorderByTableRow(firstRow, Spire.Doc.Documents.BorderStyle.Double, 0);
                TableRow lastRow = tb.LastRow;
                SetTableBorderByTableRow(lastRow, Spire.Doc.Documents.BorderStyle.Single, 0);


            }


        }

        private void DetailsOnTestConductedforPhEN2(Document document, int sectionidex)
        {

            Table tb = GetTableByFirstCell(document, "Clause", sectionidex);
            if (tb != null)
            {
                //table height
                foreach (TableRow tbr in tb.Rows)
                {
                    tbr.Height = 15.59055f;
                }
                //pr format
                foreach (TableRow tbr in tb.Rows)
                {
                    foreach (Paragraph pr in tbr.Cells[0].Paragraphs)
                    {
                        pr.Format.LeftIndent = 5.66929F;

                    }
                    foreach (Paragraph pr in tbr.Cells[1].Paragraphs)
                    {
                        pr.Format.LeftIndent = 5.66929F;

                    }
                }

                TableRow firstRow = tb.FirstRow;
                SetTableBorderByTableRow(firstRow, Spire.Doc.Documents.BorderStyle.Double, 0);
            }

            //pr format
            List<Table> tb2 = GetTablesByFirstCell(document, "Sample", sectionidex);
            if (tb2.Count > 0)
            {
                foreach (Table sampleTb in tb2)
                {
                    foreach (TableRow tbr in sampleTb.Rows)
                    {
                        foreach (Paragraph pr in tbr.Cells[0].Paragraphs)
                        {
                            pr.Format.LeftIndent = 5.66929F;

                        }

                    }


                }

            }


        }

        private void DetailsOnSeeResult1forPhEN(Document document, int sectionidex)
        {
            ListStyle listStyle = new ListStyle(document, ListType.Numbered);
            listStyle.Name = "levelstyle1";
            listStyle.Levels[0].CharacterFormat.FontName = "Arial";
            listStyle.Levels[0].CharacterFormat.FontSize = 9f;
            listStyle.Levels[0].PatternType = ListPatternType.Arabic;
            document.ListStyles.Add(listStyle);

            Section section = document.Sections[sectionidex];
            section.PageSetup.GridType = GridPitchType.LinesOnly;
            section.PageSetup.LinesPerPage = 43;
            List<TextRange> trs1 = new List<TextRange>();
            List<Paragraph> prs1 = new List<Paragraph>();
            List<TextRange> trs2 = new List<TextRange>();
            List<Paragraph> prs2 = new List<Paragraph>();

            //table column width and height
            Table tb = GetTableByFirstCell(document, "Observation", sectionidex);
            if (tb != null)
            {
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    tb.Rows[i].Height = 18.4252F;
                    tb.Rows[i].Cells[0].Width = 167.244f;
                    tb.Rows[i].Cells[1].Width = 123.0236f;
                    tb.Rows[i].Cells[2].Width = 177.1654f;
                }
                tb.TableFormat.HorizontalAlignment = RowAlignment.Center;

                //Table Pr Style
                for (int i = 1; i < tb.Rows.Count; i++)
                {
                    Paragraph pr = tb.Rows[i].Cells[0].Paragraphs[0];
                    pr.Format.LeftIndent = 5.66929F;
                }

                //note--
                Table noteTb = document.Sections[sectionidex].Tables[document.Sections[sectionidex].Tables.Count - 1] as Table;
                Paragraph newpr = new Paragraph(document);
                foreach (DocumentObject doc in noteTb.FirstRow.Cells[0].ChildObjects)
                {
                    if (doc.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph pr = doc as Paragraph;
                        foreach (DocumentObject prObj in pr.ChildObjects)
                        {
                            if (prObj is TextRange)
                            {
                                TextRange textRange = prObj as TextRange;
                                if (Regex.IsMatch(textRange.Text.Replace(" ", ""), @"([0-9]\.)\w+") == true && newpr.Text.Trim() != "")
                                {
                                    int idx = textRange.Text.IndexOf('.');
                                    textRange.Text = textRange.Text.Substring(idx + 1);
                                    prs2.Add(newpr);
                                    newpr = new Paragraph(document);
                                }
                                textRange.Text = textRange.Text.Trim();
                                newpr.ChildObjects.Add(textRange.Clone());
                            }


                        }

                    }
                }
                prs2.Add(newpr);

                foreach (Paragraph pr in prs2)
                {
                    pr.ListFormat.ApplyStyle("levelstyle1");
                    pr.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                    pr.AppendBreak(BreakType.LineBreak);
                }


                Body body2 = noteTb.OwnerTextBody;
                int index2 = body2.ChildObjects.IndexOf(noteTb);
                body2.ChildObjects.RemoveAt(index2);

                //insert new  paragraph into document--
                foreach (Paragraph pr in prs2)
                {
                    body2.ChildObjects.Insert(index2, pr);
                    index2++;
                }


                //result
                Table result = document.Sections[sectionidex].Tables[3] as Table;
                foreach (DocumentObject doc in result.FirstRow.Cells[0].ChildObjects)
                {
                    if (doc.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph pr = doc as Paragraph;
                        foreach (DocumentObject prObj in pr.ChildObjects)
                        {
                            Debug.WriteLine(prObj.DocumentObjectType);
                            if (prObj is TextRange)
                            {
                                TextRange textRange = prObj as TextRange;
                                trs1.Add(textRange);
                            }

                        }

                    }
                }



                foreach (TextRange tr in trs1)
                {
                    Paragraph pr = new Paragraph(document);
                    pr.ChildObjects.Add(tr);
                    prs1.Add(pr);
                }

                Body body1 = result.OwnerTextBody;
                int index1 = body1.ChildObjects.IndexOf(result);
                body1.ChildObjects.RemoveAt(index1);

                //insert new  paragraph into document--
                foreach (Paragraph pr in prs1)
                {
                    body1.ChildObjects.Insert(index1, pr);
                    index1++;
                }
            }


        }

        private void DetailsOnSeeResult2forPhEN(Document document, int sectionidex)
        {
            //RESULT2 있는지 체크
            Table tb = GetTableByFirstCell(document, "Observation", sectionidex);
            if (tb != null)
            {
                ListStyle listStyle = new ListStyle(document, ListType.Numbered);
                listStyle.Name = "levelstyle2";
                listStyle.Levels[0].CharacterFormat.FontName = "Arial";
                listStyle.Levels[0].CharacterFormat.FontSize = 9f;
                listStyle.Levels[0].PatternType = ListPatternType.Arabic;
                document.ListStyles.Add(listStyle);
                Section section = document.Sections[sectionidex];
                section.PageSetup.GridType = GridPitchType.LinesOnly;
                section.PageSetup.LinesPerPage = 43;
                List<Table> noteTbs = new List<Table>();
                List<TextRange> trs1 = new List<TextRange>();
                List<Paragraph> prs1 = new List<Paragraph>();
                List<TextRange> trs2 = new List<TextRange>();
                List<Paragraph> prs2 = new List<Paragraph>();

                //Note--
                int sttbindex = GetTableIndexByFirstCell(document, "Note :", sectionidex);
                for (int i = sttbindex + 1; i < document.Sections[sectionidex].Tables.Count; i++)
                {
                    Table noteTb = document.Sections[sectionidex].Tables[i] as Table;
                    noteTbs.Add(noteTb);
                    foreach (DocumentObject doc in noteTb.FirstRow.Cells[0].ChildObjects)
                    {
                        if (doc.DocumentObjectType == DocumentObjectType.Paragraph)
                        {
                            Paragraph pr = doc as Paragraph;
                            foreach (DocumentObject prObj in pr.ChildObjects)
                            {
                                Debug.WriteLine(prObj.DocumentObjectType);
                                if (prObj is TextRange)
                                {
                                    TextRange textRange = prObj as TextRange;
                                    trs2.Add(textRange);
                                }

                            }

                        }
                    }

                }
                foreach (TextRange tr in trs2)
                {
                    Paragraph pr = new Paragraph(document);
                    pr.ListFormat.ApplyStyle("levelstyle2");
                    int idx = tr.Text.IndexOf('.');
                    tr.Text = tr.Text.Substring(idx + 1).Trim();
                    pr.ChildObjects.Add(tr);
                    pr.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                    prs2.Add(pr);
                }
                for (int j = 0; j < noteTbs.Count; j++)
                {
                    Body body2 = noteTbs[j].OwnerTextBody;
                    int index2 = body2.ChildObjects.IndexOf(noteTbs[j]);
                    body2.ChildObjects.RemoveAt(index2);
                    body2.ChildObjects.Insert(index2, prs2[j]);

                }


                //table column width and height
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    tb.Rows[i].Height = 18.4252F;
                    tb.Rows[i].Cells[0].Width = 167.244f;
                    tb.Rows[i].Cells[1].Width = 123.0236f;
                    tb.Rows[i].Cells[2].Width = 177.1654f;
                }
                tb.TableFormat.HorizontalAlignment = RowAlignment.Center;

                //Table Pr Style
                for (int i = 1; i < tb.Rows.Count; i++)
                {
                    Paragraph pr = tb.Rows[i].Cells[0].Paragraphs[0];
                    pr.Format.LeftIndent = 5.66929F;
                }


                //Result
                Table result = document.Sections[sectionidex].Tables[4] as Table;
                foreach (DocumentObject doc in result.FirstRow.Cells[0].ChildObjects)
                {
                    if (doc.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        Paragraph pr = doc as Paragraph;
                        foreach (DocumentObject prObj in pr.ChildObjects)
                        {
                            Debug.WriteLine(prObj.DocumentObjectType);
                            if (prObj is TextRange)
                            {
                                TextRange textRange = prObj as TextRange;
                                trs1.Add(textRange);
                            }
                        }
                    }
                }

                foreach (TextRange tr in trs1)
                {
                    Paragraph pr = new Paragraph(document);
                    pr.ChildObjects.Add(tr);
                    prs1.Add(pr);
                }

                Body body1 = result.OwnerTextBody;
                int index1 = body1.ChildObjects.IndexOf(result);
                body1.ChildObjects.RemoveAt(index1);

                //insert new paragraph into document
                foreach (Paragraph pr in prs1)
                {
                    body1.ChildObjects.Insert(index1, pr);
                    index1++;
                }
            }

        }



        private void button1_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog vistaFolderBrowserDialog = new VistaFolderBrowserDialog();
            DialogResult ds = vistaFolderBrowserDialog.ShowDialog();
            if (ds == DialogResult.OK && !string.IsNullOrWhiteSpace(vistaFolderBrowserDialog.SelectedPath))
            {
                filepaths = Directory.GetFiles(vistaFolderBrowserDialog.SelectedPath);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Integr";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);


                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴!
                    bool hasfooterimage = false;
                    foreach (Paragraph par in document.Sections[0].HeadersFooters.Footer.Paragraphs)
                    {
                        foreach (DocumentObject docObject in par.ChildObjects)
                        {
                            //If Type of Document Object is Picture, add it into image list
                            if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                            {
                                hasfooterimage = true;

                            }
                        }
                    }


                    if (hasfooterimage == false)
                    {
                        insertImOnfooter(document);
                        removeBlankP(document);


                        //Details and styles
                        Table table = (Table)document.Sections[0].Tables[1];
                        table.Rows[3].Height = 20;
                        table.Rows[3].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;

                        foreach (Section sec in document.Sections)
                        {
                            Table tb = (Table)sec.Tables[1];
                            TextRange range = (tb.Rows[1].Cells[2].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range.CharacterFormat.FontSize = 9;
                            TextRange range2 = (tb.Rows[1].Cells[4].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range2.CharacterFormat.FontSize = 9;
                        }


                        merge_cells2(document);
                        removeBorder(document, 2);


                        //saveDocument
                        document.SaveToFile(filepath, FileFormat.Docx);
                    }



                }

            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }
        }

        private void button4_Click(object sender, EventArgs e) //PHYSICAL EN
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Physical";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);

                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴.
                    bool hasfooterimage = false;
                    if (document.Sections[0].HeadersFooters.Footer.Tables.Count > 0)
                    {
                        hasfooterimage = true;
                    }
                    if (hasfooterimage == false)
                    {


                        insertImOnfoote_renew(document);
                        removeBlankP(document);
                        //foreach (Section sec in document.Sections)
                        //{
                        //    Table tb = (Table)sec.Tables[1];
                        //    TextRange range = (tb.Rows[1].Cells[2].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                        //    range.CharacterFormat.FontSize = 9;
                        //    TextRange range2 = (tb.Rows[1].Cells[4].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                        //    range2.CharacterFormat.FontSize = 9;
                        //}
                        DetailsOnFirstPageForPhEN(document);
                        DetailsOnResultSummaryForPhEN(document);   //result summary p2
                        DetailsOnTestConductedforPhEN1(document, 2);//test conducted p3
                        DetailsOnTestConductedforPhEN2(document, 3);//test conducted  p4

                        int indexForSR1 = GetSectionindex(document, "See Result 1");
                        if (indexForSR1 != default)
                        {
                            DetailsOnSeeResult1forPhEN(document, indexForSR1);  //See Result 1
                        }
                        int indexForSR2 = GetSectionindex(document, "See Result 2");
                        if (indexForSR2 != default)
                        {
                            DetailsOnSeeResult2forPhEN(document, indexForSR2);  //See Result 2 
                        }
                        //saveDocument
                         document.SaveToFile(filepath, FileFormat.Docx);
                        //document.SaveToFile(@"C:\Users\chaeeun_kim\OneDrive - SGS\문서\work2\휘진대리님테스트\WORDTOWORD\output.docx", FileFormat.Docx);
                    }



                }

            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }

        }

        private void DetailsOnFirstPageForPhEN(Document document)
        {
            //Details and styles
            Table table = (Table)document.Sections[0].Tables[1];
            table.Rows[3].Height = 20;
            table.Rows[3].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
            table.LastRow.Cells[1].Paragraphs[0].Format.LeftIndent = 5.66929F;
            Table table2 = (Table)document.Sections[0].Tables[2];
            foreach (Paragraph pr in table2.Rows[0].Cells[1].Paragraphs)
            {
                pr.Format.LeftIndent = 5.66929F;

            }

            Table detailTb1 = (Table)document.Sections[0].Tables[4];
            foreach (TableRow tr in detailTb1.Rows)
            {
                foreach (Paragraph pr in tr.Cells[0].Paragraphs)
                {
                    pr.Format.LeftIndent = 5.66929F;

                }
                if (tr.Cells.Count > 2)
                {
                    foreach (Paragraph pr in tr.Cells[2].Paragraphs)
                    {
                        pr.Format.LeftIndent = 5.66929F;

                    }

                }
            }
        }

        private void DetailsOnResultSummaryForPhEN(Document document) //DETAIL SUMMARY FOR PHYSICAL EN
        {
            document.Sections[1].Tables[2].Rows[0].Cells[0].Paragraphs[0].Format.LeftIndent = 5.66929F;

            Table ResultSummaryTd = GetTableByFirstCell(document, "Test Requested");
            if (ResultSummaryTd != null)
            {

                foreach (TableRow tbr in ResultSummaryTd.Rows)
                {
                    foreach (Paragraph pr in tbr.Cells[0].Paragraphs)
                    {
                        pr.Format.LeftIndent = 5.66929F;
                    }

                    if (tbr.GetRowIndex() != 0 && tbr.GetRowIndex() != 1)
                    {
                        SetTableBorderByTableRow(tbr, BorderStyle.Dot, 1);

                    }

                    if (tbr.Cells[1].Paragraphs[0].Text.Trim() == "SEE RESULT 1" || tbr.Cells[1].Paragraphs[0].Text.Trim() == "SEE RESULT 2")
                    {
                        tbr.Cells[0].CellFormat.SamePaddingsAsTable = false;
                        tbr.Cells[0].CellFormat.Paddings.Top = 5.6F;
                        tbr.Cells[0].CellFormat.Paddings.Bottom = 5.6F;
                    }

                }


            }

        }


        private void DetailsOnResultSummaryForPhASTM(Document document)
        {
            Table tb = GetTableByFirstCell(document, "Clause", 2);
            string previousClause = "";
            if (tb != null)
            {

                foreach (TableRow tbr in tb.Rows)
                {
                    tbr.Height = 15.59055F;

                }

                for (int i = 1; i < tb.Rows.Count - 1; i++)
                {
                    TableRow tr = tb.Rows[i];
                    TableCell cell = tr.Cells[0];
                    Paragraph pr = tr.Cells[0].Paragraphs[0];
                    if (i == 1)
                    {
                        previousClause = pr.Text.Trim().Split('.').First();
                    }

                    if (previousClause != pr.Text.Trim().Split('.').First() && previousClause != "" && pr.Text.Trim().Split('.').First() != "")
                    {

                        SetTableBorderByTableRow(tr, BorderStyle.Single, 1);
                        SetTableBorderByTableRow(tr, BorderStyle.None, 0);
                    }
                    else
                    {
                        SetTableBorderByTableRow(tr, BorderStyle.None, 0);
                        SetTableBorderByTableRow(tr, BorderStyle.None, 1);

                    }

                    previousClause = pr.Text.Trim().Split('.').First();

                }


                TableRow lastRow = tb.LastRow;
                SetTableBorderByTableRow(lastRow, Spire.Doc.Documents.BorderStyle.None, 1);


            }

        }



        public void SetTableBorderByTableRow(TableRow tbr, BorderStyle borderStyle, int where)
        {

            foreach (TableCell tbrCell in tbr.Cells)
            {

                if (where == 0) //BOTTOM
                {
                    tbrCell.CellFormat.Borders.Bottom.BorderType = borderStyle;
                }
                if (where == 1) //TOP
                {
                    tbrCell.CellFormat.Borders.Top.BorderType = borderStyle;

                }
                if (where == 2) //LEFT
                {
                    tbrCell.CellFormat.Borders.Left.BorderType = borderStyle;

                }
                if (where == 3)//RIGHT
                {
                    tbrCell.CellFormat.Borders.Right.BorderType = borderStyle;

                }

            }
        }
        private void Bodertest(Document document)
        {

            Table tb = document.Sections[0].Tables[1] as Table;
            string previousClause = "";
            if (tb != null)
            {
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    TableRow tr = tb.Rows[i];

                    tr.RowFormat.Borders.BorderType = BorderStyle.None;
                    SetTableBorderByTableRow(tr, BorderStyle.None, 0);
                    SetTableBorderByTableRow(tr, BorderStyle.None, 1);


                }
            }
            //document.SaveToFile(filepath, FileFormat.Docx);
            //document.SaveToFile(@"C:\Users\chaeeun_kim\OneDrive - SGS\문서\work2\휘진대리님테스트\WORDTOWORD\output.docx", FileFormat.Docx);

        }

        private Table GetTableByFirstCell(Document document, string prText, int sectionIndex = default)
        {
            if (sectionIndex != default)
            {
                Section section = document.Sections[sectionIndex];
                foreach (Table table in section.Tables)
                {
                    if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                    {
                        return table;

                    }
                }
            }
            else
            {
                foreach (Section section in document.Sections)
                {
                    foreach (Table table in section.Tables)
                    {
                        if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                        {
                            return table;

                        }
                    }
                }

            }

            return null;

        }

        private List<Table> GetTablesByFirstCell(Document document, string prText, int sectionIndex = default)
        {
            List<Table> tbs = new List<Table>();
            if (sectionIndex != default)
            {
                Section section = document.Sections[sectionIndex];
                foreach (Table table in section.Tables)
                {
                    if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                    {
                        tbs.Add(table);

                    }
                }
            }
            else
            {
                foreach (Section section in document.Sections)
                {
                    foreach (Table table in section.Tables)
                    {
                        if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                        {
                            tbs.Add(table);

                        }
                    }
                }

            }

            return tbs;

        }
        private int GetTableIndexByFirstCell(Document document, string prText, int sectionIndex = default)
        {
            if (sectionIndex != default)
            {
                Section section = document.Sections[sectionIndex];
                for (int i = 0; i < section.Tables.Count; i++)
                {
                    Table table = section.Tables[i] as Table;
                    if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                    {
                        return i;

                    }
                }
            }
            else
            {
                foreach (Section section in document.Sections)
                {
                    for (int i = 0; i < section.Tables.Count; i++)
                    {
                        Table table = section.Tables[i] as Table;
                        if (table.Rows[0].Cells[0].Paragraphs[0].Text.Trim() == prText)
                        {
                            return i;

                        }
                    }
                }

            }

            return -1;

        }

        private int GetSectionindex(Document document, string st)
        {

            for (int i = 0; i < document.Sections.Count; i++)
            {
                Section section = document.Sections[i];
                foreach (Table tb in section.Tables)
                {
                    foreach (TableRow tbr in tb.Rows)
                    {
                        foreach (TableCell tbc in tbr.Cells)
                        {
                            foreach (Paragraph pr in tbc.Paragraphs)
                            {
                                if (pr.Text == st.Trim())
                                {
                                    return i;
                                }


                            }

                        }


                    }



                }


            }


            return default;



        }

        private void button5_Click(object sender, EventArgs e) //PHYSICAL ASTM
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Physical";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    Debug.WriteLine(rptname);

                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);


                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴!
                    bool hasfooterimage = false;

                    if (document.Sections[0].HeadersFooters.Footer.Tables.Count > 0)
                    {
                        hasfooterimage = true;
                    }


                    if (hasfooterimage == false)
                    {
                        insertImOnfoote_renew(document);
                        removeBlankP(document);


                        //Details and styles
                        Table table = (Table)document.Sections[0].Tables[1];
                        table.Rows[3].Height = 20;
                        table.Rows[3].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;

                        foreach (Section sec in document.Sections)
                        {
                            Table tb = (Table)sec.Tables[1];
                            TextRange range = (tb.Rows[1].Cells[2].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range.CharacterFormat.FontSize = 9;
                            TextRange range2 = (tb.Rows[1].Cells[4].ChildObjects[0] as Paragraph).ChildObjects[0] as TextRange;
                            range2.CharacterFormat.FontSize = 9;
                        }

                        DetailsOnResultSummaryForPhASTM(document); //result summary p3
                        //saveDocument
                        document.SaveToFile(filepath, FileFormat.Docx);
                        //document.SaveToFile(@"C:\Users\chaeeun_kim\OneDrive - SGS\문서\work2\휘진대리님테스트\WORDTOWORD\output.docx", FileFormat.Docx);

                    }



                }

            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string folderPath_Integr_astm = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Integr";
            string folderPath_Integr_en = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Integr";
            string folderPath_Physical_astm = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Physical";
            string folderPath_Physical_en = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Physical";
            string folderPath_Chemical_en = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Chemical";
            string folderPath_Chemical_astm = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Chemical";

            DirectoryInfo Path_Integr_Astm = new DirectoryInfo(folderPath_Integr_astm);
            DirectoryInfo Path_Integr_En = new DirectoryInfo(folderPath_Integr_en);
            DirectoryInfo Path_Physical_Astm = new DirectoryInfo(folderPath_Physical_astm);
            DirectoryInfo Path_Physical_En = new DirectoryInfo(folderPath_Physical_en);
            DirectoryInfo Path_Chemical_En = new DirectoryInfo(folderPath_Chemical_en);
            DirectoryInfo Path_Chemical_Astm = new DirectoryInfo(folderPath_Chemical_astm);

            if (Path_Integr_Astm.Exists == false)
            {
                Path_Integr_Astm.Create();
            }

            if (Path_Integr_En.Exists == false)
            {
                Path_Integr_En.Create();
            }

            if (Path_Physical_Astm.Exists == false)
            {
                Path_Physical_Astm.Create();
            }

            if (Path_Physical_En.Exists == false)
            {
                Path_Physical_En.Create();
            }

            if (Path_Chemical_En.Exists == false)
            {
                Path_Chemical_En.Create();
            }

            if (Path_Chemical_Astm.Exists == false)
            {
                Path_Chemical_Astm.Create();
            }

            button2_Click(sender, e);
            button3_Click(sender, e);
            button4_Click(sender, e);
            button5_Click(sender, e);
            button6_Click(sender, e);
            button7_Click(sender, e);
            AutoClosingMessageBox.Show("ASTM, EN 변환 완료!", "알림", 1000);
            Application.Exit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\EN_Chemical";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);

                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴
                    bool hasfooter = false;
                    if (document.Sections[0].HeadersFooters.Header.Tables.Count > 0)
                    {
                        hasfooter = true;
                    }

                    if (hasfooter == false)
                    {
                        insertImOnfoote_renew(document);
                        removeBlankP(document);

                        insertHeader(document);

                        bool isTin = Body(document);

                        //saveDocument
                        document.SaveToFile(filepath, FileFormat.Docx);
                    }

                }
            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = @"C:\Projects\Projects\Sgs\Remote_One\ReportIntegration\ReportIntegration\Bom\ASTM_Chemical";

                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepaths = Directory.GetFiles(fbd.SelectedPath);
                }
            }

            if (filepaths != null)
            {
                foreach (string filepath in filepaths)
                {

                    //get template
                    string rptname = filepath;
                    //Create word document
                    Document document = new Document();
                    //load a document
                    document.LoadFromFile(rptname);

                    //이미 한 번 정렬된 파일을 다시 정렬시키기 않기 위해서. 푸터 이미지 이미 있으면 정렬 안시킴
                    bool hasfooter = false;
                    if (document.Sections[0].HeadersFooters.Header.Tables.Count > 0)
                    {
                        hasfooter = true;
                    }

                    if (hasfooter == false)
                    {
                        FooterForASTM_CHECMICAL(document);
                        removeBlankP(document);
                        insertHeader(document);
                        BodyForASTM_CHECMICAL(document);

                        //saveDocument
                        document.SaveToFile(filepath, FileFormat.Docx);
                    }

                }
            }
            else
            { MessageBox.Show("폴더를 선택하세요"); }
        }
    }
}
