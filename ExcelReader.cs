using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ExcelPro
{
    class ExcelReader
    {
        private Timer timer = new Timer();
        string FullFileName { get; set; }
        string SheetName { get; set; }
        int ColumnStartIndex { get; set; }
        int TimeIntervalInMinutes { get; set; }

        public ExcelReader(string filePath, string sheetName, int ncolumnStartIndex, int timeInterval)
        {
            this.FullFileName = filePath;
            this.SheetName = sheetName;
            this.TimeIntervalInMinutes = timeInterval;
            //this.ColumnStartIndex = ncolumnStartIndex;
        }

        public void StartProcessing()
        {
            try
            {

                int interval = TimeIntervalInMinutes * 60000; //Convert to miliseconds

                TimeSpan timeSpan = new TimeSpan(0, 0, 0, 0, interval);


                if (!timer.Enabled)
                {
                    timer.Enabled = true;
                    timer.Interval = interval;
                    timer.Elapsed += Timer_Elapsed;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        object obj = new Object();
        public void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine("Time Elasped");
            DataTable dtExcelData = new DataTable();
            string mailContent = string.Empty;
            lock (obj)
            {
               dtExcelData = ReadExcel();

                if(dtExcelData == null || dtExcelData.Rows.Count ==0)
                {
                    Console.WriteLine("No data row to mail");
                    return;
                }

                mailContent = ExportDatatableToHtml(dtExcelData);

                SendMailThroughGmail(mailContent);
            }
        }

        /// <summary>
        /// Read Excel in xlsx format
        /// Set FullFileName, Sheet name and Start Row Index
        /// </summary>
        private DataTable ReadExcel()
        {
            DataTable dtExcelData = null;

            if (!File.Exists(FullFileName))
            {
                Console.WriteLine("File not found");
                return null;
            }

            DataTable dtDataTable = new DataTable();
            string sConnectionString = string.Empty;

            sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + FullFileName + "; Extended Properties=\"Excel 12.0;HDR=NO;TypeGuessRows=0; MaxScanRows=0; ImportMixedTypes=Text;IMEX = 0\";";

            //Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " +sFileName + ";" Extended Properties = "Excel 12.0 Macro;HDR=YES"; //Macro enable excel
            //StrConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + srcFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\";"; //xlx format 

            OleDbConnection odbcExcelConncetion = new OleDbConnection(sConnectionString);

            OleDbCommand oleDbCommand = new OleDbCommand("Select * from [" + SheetName + "$]", odbcExcelConncetion);
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter();
            oleDbDataAdapter.SelectCommand = oleDbCommand;
            
            try
            {
                oleDbDataAdapter.Fill(dtDataTable);


                if (dtDataTable.Rows.Count > 0)
                {

                    //Change datatable Column name as per Column nam
                    DataRow drExcelFirstRow = dtDataTable.Rows[ColumnStartIndex];
                    int nColumnNo = 0;

                     dtExcelData = dtDataTable.Clone();

                    foreach (DataColumn column in dtExcelData.Columns)
                    {
                        if (column.Table.Columns.Contains(Convert.ToString(drExcelFirstRow[nColumnNo]) == string.Empty ? nColumnNo.ToString() : Convert.ToString(drExcelFirstRow[nColumnNo]).Trim()))
                        {
                            column.ColumnName = nColumnNo.ToString();

                        }

                        else
                        {
                            column.ColumnName = Convert.ToString(drExcelFirstRow[nColumnNo]) == string.Empty ? nColumnNo.ToString() : Convert.ToString(drExcelFirstRow[nColumnNo]).Trim();
                        }
                        nColumnNo++;
                    }

                    nColumnNo = 0;

                    DataRow[] dr = dtDataTable.Select();

                    for (int i = ColumnStartIndex + 1; i < dtDataTable.Rows.Count; i++)
                    {
                        dtExcelData.Rows.Add(dr[i].ItemArray);
                    }
                    dtExcelData.AcceptChanges();


                    Console.WriteLine("Excel Read , Number of rows " + dtExcelData.Rows.Count );
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                dtDataTable = new DataTable();
            }

            return dtExcelData;
        }

        /// <summary>
        /// befor proceed set Allow less secure apps = true in gmail setting
        /// https://myaccount.google.com/lesssecureapps
        /// </summary>
        /// <param name="mailContent"></param>
        private void SendMailThroughGmail(string mailContent )
        {
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com"); 

            try
            {

                MailMessage mail = new MailMessage();
                 SmtpServer = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress("abc@gmail.com"); //From Mail 

                mail.To.Add("sush1223@gmail.com"); //To mail 
                mail.To.Add("xyz@gmail.com"); //To mail

                //mail.Bcc.Add(""); //add  mail to send bcc
                
                mail.Subject = "Subject";
                mail.Body = mailContent;
                mail.IsBodyHtml = true;


                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("sushant@gmail.com", "yourPassword");
                SmtpServer.EnableSsl = true;

                SmtpServer.SendAsync(mail, SmtpServer);
                Console.WriteLine("mail Send");

               // Send(mail, SmtpServer);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
            
        }

        async Task Send(MailMessage message, SmtpClient Client)
        {
           // Client.EnableSsl = EnforceSsl;
            //Client.Credentials = Credentials;

            // Send  
            await Client.SendMailAsync(message);
        }

        /// <summary>
        /// Convert Data Table to Html Table format for mail body
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected string ExportDatatableToHtml(DataTable dt)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");

            strHTMLBuilder.Append("<div>");
            strHTMLBuilder.Append("<p>");
            strHTMLBuilder.Append("Hi sflkasfasfjklsajfdkljasdf");
            strHTMLBuilder.Append("</p>");
            strHTMLBuilder.Append("</div>");


            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='lightyellow' style='font-family:Garamond; font-size:smaller'>");

            strHTMLBuilder.Append("<tr >");
            foreach (DataColumn myColumn in dt.Columns)
            {
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }

            //Close tags.  
            strHTMLBuilder.Append("</table>");

            strHTMLBuilder.Append("<p>");
            strHTMLBuilder.Append("<h5>");
            strHTMLBuilder.Append("Thanks & Regards,"); strHTMLBuilder.Append("<br>");
            strHTMLBuilder.Append("sssssss");
            strHTMLBuilder.Append("</h5>");
            strHTMLBuilder.Append("</p>");

            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");
            string Htmltext = strHTMLBuilder.ToString();
            return Htmltext;
        }
        
    }
}
