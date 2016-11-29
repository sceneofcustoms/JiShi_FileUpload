using MessagingToolkit.Barcode;
using PDFLibNet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JiShi_FileUpload
{
    public partial class FileMonitor : Form
    {
        bool working = false;
        bool working_del = false;
        string guid = string.Empty;
        string UserName = ConfigurationManager.AppSettings["FTPUserName"];
        string Password = ConfigurationManager.AppSettings["FTPPassword"];
        System.Uri Uri = new Uri("ftp://" + ConfigurationManager.AppSettings["FTPServer"] + ":" + ConfigurationManager.AppSettings["FTPPortNO"]);
        string direc_img = ConfigurationManager.AppSettings["ImagePath"];
        public enum Definition
        {
            One = 1, Two = 2, Three = 3, Four = 4, Five = 5, Six = 6, Seven = 7, Eight = 8, Nine = 9, Ten = 10
        }
        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath, string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, Definition definition)
        {
            PDFWrapper pdfWrapper = new PDFWrapper();
            pdfWrapper.LoadPDF(pdfInputPath);
            if (!System.IO.Directory.Exists(imageOutputPath))
            {
                Directory.CreateDirectory(imageOutputPath);
            }
            // validate pageNum
            if (startPageNum <= 0)
            {
                startPageNum = 1;
            }
            if (endPageNum > pdfWrapper.PageCount)
            {
                endPageNum = pdfWrapper.PageCount;
            }
            if (startPageNum > endPageNum)
            {
                int tempPageNum = startPageNum;
                startPageNum = endPageNum;
                endPageNum = startPageNum;
            }
            // start to convert each page
            for (int i = startPageNum; i <= endPageNum; i++)
            {
                pdfWrapper.ExportJpg(imageOutputPath + imageName, i, i, 300, 90);//这里可以设置输出图片的页数、大小和图片质量从默认的150--》300
                while (pdfWrapper.IsJpgBusy == true)
                {
                    System.Threading.Thread.Sleep(500);
                }
            }
            pdfWrapper.Dispose();
        }
        public FileMonitor()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.timer1.Enabled = true;
            this.timer2.Enabled = true;
            this.button1.Text = "程序运行中...";
            this.label1.Text = ((Login)this.Owner).ID;
            this.label2.Text = ((Login)this.Owner).NAME;
            this.label2.Visible = true;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!working)//如果timer当前空闲  防止处理任务的时间大于轮循的时间
            {
                working = true;
                string direc_pdf = ConfigurationManager.AppSettings["PdfPath"];
                string direc_upload = DateTime.Now.ToString("yyyy-MM-dd");
                string barcode = string.Empty;
                string sql = string.Empty;
                FtpHelper ftp = new FtpHelper(Uri, UserName, Password);
                DirectoryInfo dc = new DirectoryInfo(direc_pdf);
                foreach (FileInfo file in dc.GetFiles("*.pdf", SearchOption.AllDirectories))
                {
                    //1 对pdf文件进行条码识别
                    if (!FileIsUsed(file.FullName))//如果文件未被占用
                    {
                        try
                        {
                            guid = Guid.NewGuid().ToString(); //将文件重命名
                            file.MoveTo(direc_pdf + guid + ".pdf");
                            //将PDF文件的第一页转图片
                            ConvertPDF2Image(direc_pdf + guid + ".pdf", direc_img, guid + ".jpg", 1, 1, ImageFormat.Jpeg, Definition.Ten);
                            //在图片文件中读取条形码信息
                            BarcodeDecoder barcodeDecoder = new BarcodeDecoder();
                            if (File.Exists(direc_img + guid + ".jpg"))
                            {
                                System.Drawing.Bitmap image = new System.Drawing.Bitmap(direc_img + guid + ".jpg");
                                Dictionary<DecodeOptions, object> decodingOptions = new Dictionary<DecodeOptions, object>();
                                List<BarcodeFormat> possibleFormats = new List<BarcodeFormat>(10);
                                //possibleFormats.Add(BarcodeFormat.DataMatrix);
                                //possibleFormats.Add(BarcodeFormat.QRCode);
                                //possibleFormats.Add(BarcodeFormat.PDF417);
                                //possibleFormats.Add(BarcodeFormat.Aztec);
                                //possibleFormats.Add(BarcodeFormat.UPCE);
                                //possibleFormats.Add(BarcodeFormat.UPCA);
                                possibleFormats.Add(BarcodeFormat.Code128);
                                //possibleFormats.Add(BarcodeFormat.Code39);
                                //possibleFormats.Add(BarcodeFormat.ITF14);
                                //possibleFormats.Add(BarcodeFormat.EAN8);
                                possibleFormats.Add(BarcodeFormat.EAN13);
                                //possibleFormats.Add(BarcodeFormat.RSS14);
                                //possibleFormats.Add(BarcodeFormat.RSSExpanded);
                                //possibleFormats.Add(BarcodeFormat.Codabar);
                                //possibleFormats.Add(BarcodeFormat.MaxiCode);
                                decodingOptions.Add(DecodeOptions.TryHarder, true);
                                decodingOptions.Add(DecodeOptions.PossibleFormats, possibleFormats);
                                Result decodedResult = barcodeDecoder.Decode(image, decodingOptions);
                                while (decodedResult == null)
                                {
                                    System.Threading.Thread.Sleep(500);
                                }
                                barcode = decodedResult.Text;
                            }
                        }
                        catch (Exception ex)
                        {
                            barcode = string.Empty;
                            working = false;
                        }
                        //上传文件
                        ftp.CreateDirectory(direc_upload, true);
                        bool res = ftp.UploadFile(direc_pdf + guid + ".pdf", @"\" + direc_upload + @"\" + guid + ".pdf");
                        //如果上传成功!插入数据库记录
                        if (res)
                        {
                            OracleConnection conn = new OracleConnection(ConfigurationManager.AppSettings["strconn"]);
                            conn.Open();
                            OracleCommand com = null;
                            if (!string.IsNullOrEmpty(barcode)) //如果条形码识别成功
                            {
                                sql = "select * from list_attachment where ordercode='" + barcode + "'";
                                com = new OracleCommand(sql, conn);
                                OracleDataAdapter da = new OracleDataAdapter(com);
                                DataTable dt_att = new DataTable();
                                da.Fill(dt_att);
                                if (dt_att.Rows.Count > 0)//如果该订单的文件已经存在 删除表记录和文件
                                {
                                    sql = "delete from list_attachment where ordercode='" + barcode + "'";
                                    com = new OracleCommand(sql, conn);
                                    com.ExecuteNonQuery();
                                    ftp.DeleteFile(dt_att.Rows[0]["FILEPATH"] + "");
                                    //如果该文件曾经被复制到其他的订单号下  
                                    if (!string.IsNullOrEmpty(dt_att.Rows[0]["COPYORDERCODE"] + ""))
                                    {
                                        sql = "select * from list_attachment where ordercode='" + dt_att.Rows[0]["COPYORDERCODE"] + "'";
                                        com = new OracleCommand(sql, conn);
                                        da = new OracleDataAdapter(com);
                                        DataTable dt_gl = new DataTable();
                                        da.Fill(dt_gl);
                                        sql = "delete from list_attachment where ordercode='" + dt_att.Rows[0]["COPYORDERCODE"] + "'";
                                        com = new OracleCommand(sql, conn);
                                        com.ExecuteNonQuery();
                                        ftp.DeleteFile(dt_gl.Rows[0]["FILEPATH"] + "");
                                    }
                                }
                            }
                            sql = @"insert into list_attachment (ID,FILEPATH,FILENAME,FILESIZE,ORDERCODE,CREATETIME,CREATEUSERID) 
                                  VALUES (LIST_ATTACHMENT_ID.NEXTVAL,'{0}','{1}','{2}','{3}',sysdate,'{4}')";
                            sql = string.Format(sql, @"/" + direc_upload + @"/" + guid + ".pdf", guid + ".pdf", file.Length, barcode, this.label1.Text);
                            com = new OracleCommand(sql, conn);
                            com.ExecuteNonQuery();
                            if (!string.IsNullOrEmpty(barcode))
                            {
                                sql = "update list_order set FILERELATE='1' where code='" + barcode + "'";
                                com = new OracleCommand(sql, conn);
                                com.ExecuteNonQuery();
                            }
                            if (this.checkBox1.Checked)
                            {
                                //DJRI161100579 DJRE161100579 DJCI161100579 DJCE161100579  GJI161100553 GJE161100553  
                                string newcode = string.Empty;
                                string preffix = barcode.Substring(0, 3);
                                if (preffix == "DJR" || preffix == "DJC" || preffix == "GJI" || preffix == "GJE")
                                {
                                    if (barcode.IndexOf("I") > 0)
                                    {
                                        newcode = barcode.Replace("I", "E");
                                    }
                                    if (barcode.IndexOf("E") > 0)
                                    {
                                        newcode = barcode.Replace("E", "I");
                                    }
                                    string guid2 = Guid.NewGuid().ToString(); //复制并重命名文件
                                    file.CopyTo(direc_pdf + guid2 + ".pdf");
                                    sql = @"insert into list_attachment (ID,FILEPATH,FILENAME,FILESIZE,ORDERCODE,CREATETIME,CREATEUSERID,CREATENAME)
                                    VALUES (LIST_ATTACHMENT_ID.NEXTVAL,'{0}','{1}','{2}','{3}',sysdate,'{4}','{5}')";
                                    sql = string.Format(sql, @"/" + direc_upload + @"/" + guid2 + ".pdf", guid2 + ".pdf", file.Length, newcode, this.label1.Text, this.label2.Text);
                                    com = new OracleCommand(sql, conn);
                                    com.ExecuteNonQuery();
                                    bool result = ftp.UploadFile(direc_pdf + guid2 + ".pdf", @"\" + direc_upload + @"\" + guid2 + ".pdf");
                                    sql = "update list_attachment set copyordercode='" + newcode + "' where ordercode='" + barcode + "'";
                                    com = new OracleCommand(sql, conn);
                                    com.ExecuteNonQuery();
                                    if (result)
                                    {
                                        FileInfo fi_new = new FileInfo(direc_pdf + guid2 + ".pdf");
                                        fi_new.MoveTo(direc_img + guid2 + ".pdf");
                                        sql = "update list_order set FILERELATE='1' where code='" + newcode + "'";
                                        com = new OracleCommand(sql, conn);
                                        com.ExecuteNonQuery();
                                    }
                                }
                            }
                            conn.Close();
                        }
                        //将文件搬移至已上传的目录
                        file.MoveTo(direc_img + guid + ".pdf");
                    }
                }
                working = false;
            }
        }
        public static Boolean FileIsUsed(String fileFullName)
        {
            Boolean result = false;
            //判断文件是否存在，如果不存在，直接返回 false
            if (!System.IO.File.Exists(fileFullName))
            {
                result = false;
            }//end: 如果文件不存在的处理逻辑
            else
            {//如果文件存在，则继续判断文件是否已被其它程序使用
                //逻辑：尝试执行打开文件的操作，如果文件已经被其它程序使用，则打开失败，抛出异常，根据此类异常可以判断文件是否已被其它程序使用。
                System.IO.FileStream fileStream = null;
                try
                {
                    fileStream = System.IO.File.Open(fileFullName, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None);
                    result = false;
                }
                catch (System.IO.IOException ioEx)
                {
                    result = true;
                }
                catch (System.Exception ex)
                {
                    result = true;
                }
                finally
                {
                    if (fileStream != null)
                    {
                        fileStream.Close();
                    }
                }

            }//end: 如果文件存在的处理逻辑
            //返回指示文件是否已被其它程序使用的值
            return result;
        }
        //删除历史文件
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (!working_del)//如果timer当前空闲 
            {
                working_del = true;
                foreach (string str in Directory.GetFiles(direc_img))
                {
                    FileInfo fi = new FileInfo(str);
                    fi.Delete();
                }
                working_del = false;
            }
        }
    }
}
