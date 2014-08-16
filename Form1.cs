using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using System.Xml;
using System.Globalization;
using System.Collections;

namespace PublicProject
{
    public partial class Form1 : Form
    {
        #region "DataMembers"

        //ver1.4 change start
        //string mErrorList = string.Empty;
        //ver1.4 change end

        #endregion

        #region "Constructors"

        public Form1()
        {
            InitializeComponent();
            toolStripStatusLabel1.Text = "Select an .XLSX file";
        }

        #endregion

        #region "Control Events"
        
        /// <summary>
        /// Handles the btnSelect Click event.It opens up the file select dialog.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            ofdXlsFile.Filter = "XLS files|*.xlsx|All files (*.*)|*.*";
            DialogResult lDr = ofdXlsFile.ShowDialog();
            if (lDr.ToString().ToUpper() == "OK")
            {
                txtFilePath.Text = ofdXlsFile.FileName;
                txtFilePath.ReadOnly = true;

                toolStripStatusLabel1.Text = "1 file selected.";
            }
        }

        /// <summary>
        /// Handles the btnProcess Click event. It generates the French and the default html files.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (txtFilePath.Text.Trim() != string.Empty)
            {
                //ver1.4 change start
                //mErrorList = string.Empty;
                //ver1.4 change end

                toolStripStatusLabel1.Text = "Working ...";

                string lHtml = CreateTextForHTMLFile(false);
                //ver1.4 change start
                //if (mErrorList == string.Empty)
                //{
                    //ver1.4 change end
                CreateHTMLFile(ReadConfigFile("GeneratedFilePath_EN"), lHtml);

                lHtml = string.Empty;

                lHtml = CreateTextForHTMLFile(true);
                CreateHTMLFile(ReadConfigFile("GeneratedFilePath_FR"), lHtml);

                toolStripStatusLabel1.Text = "Processing Complete.";
                //ver1.4 change start
                //}
                //else
                //{
                //    toolStripStatusLabel1.Text = "Error in Excel Sheet.";
                //    //ver1.3 change start
                //    Error lErrorForm = new Error(mErrorList.Remove(mErrorList.Length - 1, 1));
                //    lErrorForm.ShowDialog();
                //    //MessageBox.Show("Error in these cells : " + mErrorList.Remove(mErrorList.Length-1,1));
                //    //ver1.3 change end
                //}
                //ver1.4 change end
            }
            else
            {
                MessageBox.Show("Select a file first.");
            }
        }

        #endregion

        #region "Private Functions"

        /// <summary>
        /// This functino reads the excel file from the path given in the parameter
        /// </summary>
        /// <param name="pPath"></param>
        private DataTable ReadExcelFile(string pPath)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();            
                object misValue = System.Reflection.Missing.Value;
                
                Excel.Workbook lWorkbook = app.Workbooks.Open(
                                             txtFilePath.Text.Trim(), 0, true, 5,
                                              "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                                              0, true,misValue,misValue);
                Excel.Sheets lSheets = lWorkbook.Worksheets;
                Excel.Worksheet lWorksheet = (Excel.Worksheet)lSheets.get_Item(1);
                                
                //ver1.4 change start
                //added a new parameter in the oledb connection read everything as string
                //OleDbConnection con = new OleDbConnection("Provider= Microsoft.ACE.OLEDB.12.0;Data Source=" + pPath + "; Extended Properties=\"Excel 12.0;HDR=YES;\"");
                OleDbConnection con = new OleDbConnection("Provider= Microsoft.ACE.OLEDB.12.0;Data Source=" + pPath + "; Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;\"");
                //ver1.4 change end
                OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + lWorksheet.Name + "$]" , con);
                
                DataTable ldt = new DataTable();
                da.Fill(ldt);

                //ver1.4 changes start
                //Close running excel app
                con.Close();
                da.Dispose();
                lWorkbook.Close(false, misValue, misValue);

                app.Quit();
                
                //Use the Com Object interop marshall to release the excel object
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lSheets);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                
                app = null;
                lSheets = null;
                lWorksheet = null;
                lWorkbook = null;

                //force a garbage collection 
                System.GC.Collect();
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true); 
                //ver1.4 changes end

                return ldt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// This function reads the app.config. 
        /// </summary>
        /// <param name="pSettingName">Key name for the setting</param>
        /// <returns></returns>
        private string ReadConfigFile(string pSettingName)
        {
            return ConfigurationManager.AppSettings[pSettingName];
        }        

        //ver1.2 change start
        /// <summary>
        /// This function reads the value from Excel file and parses the rows to make the html for the file
        /// </summary>
        /// <param name="pFrench">a booleans check that tells weather we have to translate the text or not</param>
        /// <returns></returns>
        //private string GenerateFiles(bool pFrench)
        //{
        //    DataTable ldt = new DataTable();
        //    StringBuilder lsb = new StringBuilder();
        //    int lEmptyRows = 1;
            
            

        //    ldt = ReadExcelFile(txtFilePath.Text.Trim());

        //    for (int i = 5; i < ldt.Rows.Count; i++)
        //    {
        //        if (ldt.Rows[i][0].ToString() == string.Empty && ldt.Rows[i][1].ToString() == string.Empty && ldt.Rows[i][3].ToString() == string.Empty)
        //        {
        //            lEmptyRows = lEmptyRows + 1;
        //            if (lEmptyRows >= 3)
        //            {   
        //                break;                        
        //            }
        //        }

        //        if (ldt.Rows[i][21].ToString().ToUpper() != "N/A" && ldt.Rows[i][21].ToString().ToUpper() != string.Empty)
        //        {
        //            lEmptyRows = 1;
                    
        //            if (pFrench == false)
        //            {
        //                lsb.Append("<TR><TD>"); lsb.Append( ProcessStringForChanges( ldt.Rows[i][2].ToString() , false) ); lsb.Append("</TD>"); //<TD>#value of column C</TD>
        //                lsb.Append("<TD>"); lsb.Append( ProcessStringForChanges( ldt.Rows[i][3].ToString(), false ) ); lsb.Append("</TD>"); //<TD>#value of column D</TD>
        //                lsb.Append("<TD>"); lsb.Append( ProcessStringForChanges( ldt.Rows[i][5].ToString(), false ) ); lsb.Append("</TD>"); //<TD>#value of column F</TD>
        //                lsb.Append("<TD>"); lsb.Append( ProcessStringForChanges( ldt.Rows[i][6].ToString(), false ) ); lsb.Append("</TD>"); //<TD>#value of column G</TD>
        //                lsb.Append("<TD>"); lsb.Append( ProcessStringForChanges( ldt.Rows[i][7].ToString(), false ) ); lsb.Append("</TD>"); //<TD>#value of column H</TD>
        //                lsb.Append("<TD>"); lsb.Append(ldt.Rows[i][8].ToString()); lsb.Append("</TD>"); //<TD>#value of column I</TD>
        //                lsb.Append("<TD>"); lsb.Append(ldt.Rows[i][9].ToString()); lsb.Append("</TD>"); //<TD>#value of column J</TD>
        //                //ver1.1 date change start
        //                //lsb.Append("<TD>"); lsb.Append(ldt.Rows[i][12].ToString()); lsb.Append("</TD>"); //<TD>#value of column M</TD>
        //                lsb.Append("<TD>"); lsb.Append(Convert.ToDateTime(ldt.Rows[i][12].ToString()).ToShortDateString() ); lsb.Append("</TD>"); //<TD>#value of column M</TD>
        //                //ver1.1 datechange end
        //                lsb.Append("<TD>"); lsb.Append(ldt.Rows[i][10].ToString()); lsb.Append("</TD>"); //<TD>#value of column K</TD>
        //                lsb.Append("<TD>"); lsb.Append(ldt.Rows[i][21].ToString()); lsb.Append("</TD></TR>"); //<TD>#value of column V</TD></TR>
        //            }
        //            else
        //            {
        //                lsb.Append("<TR><TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][2].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column C</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][3].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column D</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][5].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column F</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][6].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column G</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][7].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column H</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][8].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column I</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][9].ToString()) ); lsb.Append("</TD>"); //<TD>#value of column J</TD>
        //                //ver1.1 date change start
        //                //lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][12].ToString())); lsb.Append("</TD>"); //<TD>#value of column M</TD>
        //                lsb.Append("<TD>"); lsb.Append(TranslateToFrench(Convert.ToDateTime(ldt.Rows[i][12].ToString()).ToShortDateString())); lsb.Append("</TD>"); //<TD>#value of column M</TD>
        //                //ver1.1 datechange end
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][10].ToString())); lsb.Append("</TD>"); //<TD>#value of column K</TD>
        //                lsb.Append("<TD>"); lsb.Append( TranslateToFrench( ldt.Rows[i][21].ToString())); lsb.Append("</TD></TR>"); //<TD>#value of column V</TD></TR>
        //            }
        //        }

        //    }
        //    return lsb.ToString();
        //}
        //ver1.2 change end

        /// <summary>
        /// This function creates an HTML file. If the file is already created then it opens it.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <param name="pText"></param>
        private void CreateHTMLFile(string pFileName,string pText)
        {
            StreamWriter sw = null;
            FileStream fs = File.Open(pFileName,
                                    FileMode.OpenOrCreate,
                                    FileAccess.Write);

            // generate a file stream with UTF8 characters
            sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            
            sw.Write(pText);
            
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// This function creates the complete html file. It includes the header,main body and the footer of the html file.
        /// </summary>
        /// <param name="pFrench"></param>
        /// <returns></returns>
        private string CreateTextForHTMLFile(bool pFrench)
        {
            StringBuilder lCompleteHTML = new StringBuilder();            

            if (pFrench == false)
            {
                //English Version
                
                //ver1.3 change start
                //lCompleteHTML.Append(ReadExistingHTMLFile(ReadConfigFile("Header_EN_Path")));
                lCompleteHTML.Append( SearchAndReplaceHeaderFilesForKeywords( ReadExistingHTMLFile( ReadConfigFile("Header_EN_Path") ) ));                
                //ver1.3 change end

                //ver1.2 change start
                //lCompleteHTML.Append(GenerateFiles(pFrench));
                lCompleteHTML.Append(GenerateMainBody(pFrench));                
                //ver1.2 change end

                //ver1.5 change start
                //lCompleteHTML.Append(ReadExistingHTMLFile(ReadConfigFile("Footer_Path")));
                lCompleteHTML.Append(SearchAndReplaceFooterWithKeywords (ReadExistingHTMLFile(ReadConfigFile("Footer_Path")) ));
                //ver1.5 change end
            }
            else
            {
                //French Version

                //ver1.3 change start
                //lCompleteHTML.Append(ReadExistingHTMLFile(ReadConfigFile("Header_FR_Path")));
                lCompleteHTML.Append(SearchAndReplaceHeaderFilesForKeywords( ReadExistingHTMLFile( ReadConfigFile("Header_FR_Path") ) ));
                //ver1.3 change end

                //ver1.2 change start
                //lCompleteHTML.Append(GenerateFiles(pFrench));
                lCompleteHTML.Append(GenerateMainBody(pFrench));
                //ver1.2 change end

                //ver1.5 change start
                //lCompleteHTML.Append(ReadExistingHTMLFile(ReadConfigFile("Footer_Path")));
                lCompleteHTML.Append( SearchAndReplaceFooterWithKeywords( ReadExistingHTMLFile(ReadConfigFile("Footer_Path")) ));
                //ver1.5 change end

                
            }

            return lCompleteHTML.ToString();
        }

        /// <summary>
        /// This function reads the existing HTML files and return the content as string. It reads files like Header and footer files.
        /// </summary>
        /// <param name="pFilePath"></param>
        /// <returns></returns>
        private string ReadExistingHTMLFile(string pFilePath)
        {
            string lHtmlContent = string.Empty;
            
            System.IO.TextReader tr = new StreamReader(pFilePath);
            lHtmlContent = tr.ReadToEnd();
            tr.Close();

            return lHtmlContent;            
        }
        
        /// <summary>
        /// This function translates the words to French by looking into an XML file.
        /// </summary>
        /// <param name="pEnglishWord">The word to translate</param>
        /// <returns>the translated word or the orignal word depending on the fact that it found the value or not</returns>
        private string TranslateToFrench(string pEnglishWord)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(ReadConfigFile("Translate_XML"));            
            //XmlNode node = doc.SelectSingleNode("//ITEMS//ITEM[ENGLISH='" + pEnglishWord + "']");
            //ver1.1 apostrophe bug start
            //XmlNode node = doc.SelectSingleNode("//ITEMS//ITEM[translate(ENGLISH, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='"+ pEnglishWord.ToLower().Trim() +"']");
            XmlNode node = doc.SelectSingleNode("//ITEMS//ITEM[translate(ENGLISH, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')=" + XPathLiteral(pEnglishWord.ToLower().Trim()) + "]");
            //ver1.1 apostrophe bug end
            
            if (node != null)
            {
                //Return the french word
                return node.LastChild.InnerText;
            }
            else
            {
                //Return the same word
                return pEnglishWord;
            }
        }

        #endregion

        #region "ver1.1 change"

        /// <summary>
        /// Produce an XPath literal equal to the value if possible; if not, produce
        /// an XPath expression that will match the value.
        /// 
        /// Note that this function will produce very long XPath expressions if a value
        /// contains a long run of double quotes.
        /// </summary>
        /// <param name="value">The value to match.</param>
        /// <returns>If the value contains only single or double quotes, an XPath
        /// literal equal to the value.  If it contains both, an XPath expression,
        /// using concat(), that evaluates to the value.</returns>
        static string XPathLiteral(string value)
        {
            // if the value contains only single or double quotes, construct
            // an XPath literal
            if (!value.Contains("\""))
            {
                return "\"" + value + "\"";
            }
            if (!value.Contains("'"))
            {
                return "'" + value + "'";
            }

            // if the value contains both single and double quotes, construct an
            // expression that concatenates all non-double-quote substrings with
            // the quotes, e.g.:
            //
            //    concat("foo", '"', "bar")
            StringBuilder sb = new StringBuilder();
            sb.Append("concat(");
            string[] substrings = value.Split('\"');
            for (int i = 0; i < substrings.Length; i++)
            {
                bool needComma = (i > 0);
                if (substrings[i] != "")
                {
                    if (i > 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append("\"");
                    sb.Append(substrings[i]);
                    sb.Append("\"");
                    needComma = true;
                }
                if (i < substrings.Length - 1)
                {
                    if (needComma)
                    {
                        sb.Append(", ");
                    }
                    sb.Append("'\"'");
                }

            }
            sb.Append(")");
            return sb.ToString();
        }

        #endregion

        #region "ver1.2 change"

        /// <summary>
        /// This function verifies the Cell values and does actions according to the cell values
        /// </summary>
        /// <param name="pCellValue">Cell values got from the excel sheet</param>
        /// <param name="pFrench">Is the translation to frech is required or not.If yes then this functino translates that otherwise does nothing</param>
        /// <param name="pExcelIndex">Index of the column whose value is going to be processed</param>
        /// <param name="pRowIndex">Row index of the row in the excel sheet whose cell is under process</param>
        /// <returns>returns the processed cell value</returns>
        private string ProcessStringForChanges(string pCellValue,bool pFrench,int pExcelIndex,int pRowIndex)
        {
            string pRetValue = string.Empty;
            //ver1.4 change start
            //if (pCellValue.Trim() == string.Empty && pExcelIndex != 12)
            //ver1.4 change end
            if (pCellValue.Trim() == string.Empty)
            {
                pRetValue = "&nbsp;&nbsp;";               
            }
            //ver1.3 change start
            //else if (pCellValue.Trim().Contains("#EN_DATE#"))
            //{
            //    //Return the today's date in a specific format
            //    //dd Mon YYYY, hh:mi am" (eg: 13 Feb 2012 10:50pm)
            //    //Output: 15 Feb 2012 11:05 PM
            //    pRetValue = DateTime.Now.ToString("dd MMM yyyy hh:mmtt");
            //}
            //else if (pCellValue.Trim().Contains("#FR_DATE#"))
            //{
            //    //Return the Today's date in French
            //    //dd Mon YYYY, hh24:mi" (eg: 13 Feb 2012 22:50)
            //    //Output: 15 févr. 2012 23:06
            //    CultureInfo culture = new CultureInfo("fr-CA", true);
            //    pRetValue = DateTime.Now.ToString("dd MMM yyyy H:mm", culture);
            //}
            //ver1.3 change end

            //ver1.4 change start
                //no need to parse date any more as whatever is in the column export it as it is
            //If is is the date column then convert the datetime to date string only.
            //else if (pExcelIndex == 12)
            //{
            //    //First validate that the date is correct or not else show the error
            //    DateTime lParsedDateTime;
            //    if (DateTime.TryParse(pCellValue,out lParsedDateTime) == false)                
            //        mErrorList = mErrorList + "M" +(pRowIndex + 2).ToString() + ",";                
            //    else if (Convert.ToDateTime(pCellValue).Year < 2010)                
            //        mErrorList = mErrorList + "M" + (pRowIndex + 2).ToString() + ",";                
            //    else if(pCellValue.Trim() == string.Empty)
            //        mErrorList = mErrorList + "M" + (pRowIndex + 2).ToString() + ",";
            //    else
            //    {
            //        pRetValue = Convert.ToDateTime(pCellValue).ToShortDateString();
            //    }
            //}
                //ver1.4 change end
            else if (pFrench == true)
            {
                pRetValue = TranslateToFrench(pCellValue);
            }
            else
            {
                pRetValue = pCellValue;
            }
            return pRetValue;
        }

        /// <summary>
        /// This function generates the HTML for a row
        /// </summary>
        /// <param name="pdrc">Datarow whose html is needed to be generated</param>
        /// <param name="pFrench">Conversion to french is required or not</param>
        /// <param name="pIndexArr">Splitted array of indexes whose value needs to be appended in the HTML</param>
        /// <param name="pRowIndex">Row index of the row in the excel sheet</param>
        /// <returns>generated html</returns>
        private string GenerateHTMLFromExcelFile(DataRow pdrc, bool pFrench, string[] pIndexArr,int pRowIndex)
        {
            StringBuilder lsb = new StringBuilder();
            int i = 0;

            lsb.Append("<TR>");
            //Loop through all the indexes and generate the HTML
            while (pIndexArr.Length != i)
            {                
                lsb.Append("<TD>");
                lsb.Append(ProcessStringForChanges(pdrc[Convert.ToInt32(pIndexArr[i])].ToString(), pFrench, Convert.ToInt32(pIndexArr[i]), pRowIndex)); 
                lsb.Append("</TD>"); //<TD>#value of column C</TD>                

                i++;
            }

            lsb.Append("</TR>");
            return lsb.ToString();
        }

        /// <summary>
        /// This function generates the complete html main body of the excel file.
        /// </summary>
        /// <param name="pFrench">conversion to french is required or not</param>
        /// <returns>returns the complete generated string</returns>
        public string GenerateMainBody(bool pFrench)
        {   
            DataTable ldt = new DataTable();
            StringBuilder lsb = new StringBuilder();
            int lEmptyRows = 1;

            //Get the Column index from config file
            string lIndex = ReadConfigFile("ExcelSheetIndexNumbers");
            //Split the indexes
            string[] lIndexArr = lIndex.Split(',');


            ldt = ReadExcelFile(txtFilePath.Text.Trim());

            for (int i = 5; i < ldt.Rows.Count; i++)
            {
                if (ldt.Rows[i][0].ToString() == string.Empty && ldt.Rows[i][1].ToString() == string.Empty && ldt.Rows[i][3].ToString() == string.Empty)
                {
                    lEmptyRows = lEmptyRows + 1;
                    if (lEmptyRows >= 3)
                    {   
                        break;                        
                    }
                }

                if (ldt.Rows[i][21].ToString().ToUpper() != "N/A" && ldt.Rows[i][21].ToString().ToUpper() != string.Empty)
                {
                    lEmptyRows = 1;
                    lsb.Append(GenerateHTMLFromExcelFile(ldt.Rows[i], pFrench, lIndexArr,i));
                }
            }
            return lsb.ToString();
        }

        #endregion

        #region "ver1.3 changes"
        
        private string SearchAndReplaceHeaderFilesForKeywords(string pHeaderHtml)
        {
            string pRetValue = string.Empty;
            if (pHeaderHtml.Contains("#EN_DATE#"))
            {                
                //Return the today's date in a specific format
                //dd Mon YYYY, hh:mi am" (eg: 13 Feb 2012 10:50pm)
                //Output: 15 Feb 2012 11:05 PM
                pHeaderHtml = pHeaderHtml.Replace("#EN_DATE#", DateTime.Now.ToString("dd MMM yyyy hh:mmtt"));
            }
            
            if (pHeaderHtml.Contains("#FR_DATE#"))
            {
                //Return the Today's date in French
                //dd Mon YYYY, hh24:mi" (eg: 13 Feb 2012 22:50)
                //Output: 15 févr. 2012 23:06
                CultureInfo culture = new CultureInfo("fr-CA", true);
                pHeaderHtml = pHeaderHtml.Replace("#FR_DATE#", DateTime.Now.ToString("dd MMM yyyy H:mm", culture));
            }

            return pHeaderHtml;
        }

        #endregion

        #region "ver1.5 change"
        /// <summary>
        /// this function replaces the filename and the file date if they exist
        /// </summary>
        /// <param name="pFooterHtml"></param>
        /// <returns></returns>
        private string SearchAndReplaceFooterWithKeywords(string pFooterHtml)
        {
            if (pFooterHtml.Contains("#FILE#"))
            {
                string[] lSplittedPath = txtFilePath.Text.Trim().Split('\\');
                pFooterHtml = pFooterHtml.Replace("#FILE#", lSplittedPath[lSplittedPath.Length - 1]);
            }

            if (pFooterHtml.Contains("#FDATE#"))
            {
                //Mon dd yyyy hh:mi - Englsih only
                //replace with the time stamp of the file   
                //ver1.6 change start
                //pFooterHtml = pFooterHtml.Replace("#FDATE#", DateTime.Now.ToString("MMM dd yyyy hh:mmtt"));
                pFooterHtml = pFooterHtml.Replace("#FDATE#", getFileModifiedDate(txtFilePath.Text.Trim()));
                //ver1.6 change end
            }

            return pFooterHtml;
        }

        #endregion

        #region "ver1.6 change"

        /// <summary>
        /// this function returns the last date modified of the file
        /// </summary>
        /// <param name="pPath"></param>
        /// <returns></returns>
        public string getFileModifiedDate(string pPath)
        {
            DateTime ldt = System.IO.File.GetLastWriteTime(pPath);
            return ldt.ToString("MMM dd yyyy hh:mmtt");
        }

        #endregion
    }
}