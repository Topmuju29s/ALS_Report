using Access_data.Model;
using Access_data.Utilities;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Access_data.DatasetImg.ImageSet;

namespace Access_data.Service
{
    public class ReportService
    {
        private string DirectoryPath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", ""), "Template");
        public string testGetItem()
        {
            return "report is OK";
        }
        public Stream Report01(string PACKAGENO)
        {
            try
            {
                using (masterEntities context = new masterEntities())
                {
                    // Value Process
                    var resultss = context.TB_R_PACKAGE.ToList();
                    var result = context.TB_R_PACKAGE.Where(x => x.Package_No == PACKAGENO).FirstOrDefault();
                    if (result == null)
                    {
                        throw new Exception("CB_Package_No Not Found !");
                    }
                    var AfterJoin = (from trp in context.TB_R_PACKAGE
                                     join trb in context.TB_R_BA on trp.Package_ID equals trb.Package_ID
                                     join tri in context.TB_R_INVOICE on trb.BA_ID equals tri.BA_ID
                                     where trp.Package_No == PACKAGENO
                                     select new
                                     {
                                         trp.Package_No,
                                         trb.BA_No,
                                         tri.Invoice_No,
                                         tri.Invoice_Issue_Date,
                                         trb.BA_ID
                                     }).OrderBy(x => x.Package_No).ThenBy(p => p.BA_No).ToList();
                    ReportDocument rpt = new ReportDocument();
                    rpt.Load(Path.Combine(DirectoryPath, "report01.rpt"));
                    // Generate BARCODE
                    BarcodeLib.Barcode b = new BarcodeLib.Barcode();
                    Image img = b.Encode(BarcodeLib.TYPE.CODE128, AfterJoin.FirstOrDefault().Package_No, Color.Black, Color.White, 375, 120);
                    MemoryStream fs = new MemoryStream();
                    ((Bitmap)img).Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                    byte[] data = fs.ToArray();
                    fs.Dispose();


                    var groupBA = AfterJoin.GroupBy(x => x.BA_No).ToList();
                    List<Report01> resultofGroupCustomModel = new List<Report01>();
                    var BillingAppointment = "Billing Appoint no.";
                    int rowStart = 0;
                    var dt = new Access_data.DatasetImg.ImageSet();
                    dt.DataTable1.Rows.Add(data);
                    foreach (var BA_Group in groupBA)
                    {
                        var TEMPCHECKGROUP = string.Empty;
                        foreach (var item in BA_Group)
                        {
                            Report01 itemonKey = new Report01();
                            if (rowStart == 0)
                            {
                                itemonKey.BillingappointmentnoText = BillingAppointment;
                            }
                            else
                            {
                                itemonKey.BillingappointmentnoText = "";
                            }
                            if (string.IsNullOrEmpty(TEMPCHECKGROUP))
                            {
                                itemonKey.BANO = item.BA_No;
                                itemonKey.InvoiceNoText = "Invoice No.";
                            }
                            else
                            {
                                itemonKey.BANO = "";
                                itemonKey.InvoiceNoText = "";
                            }
                            itemonKey.InvoiceNo = item.Invoice_No;
                            itemonKey.InvoiceDate = item.Invoice_Issue_Date?.ToString("dd-MMMM-yyyy");
                            itemonKey.InvoiceDateText = "Invoice Date";
                            resultofGroupCustomModel.Add(itemonKey);
                            dt.DataTable2.Rows.Add(new Object[] { itemonKey.BANO, itemonKey.InvoiceNo, itemonKey.InvoiceDate, itemonKey.BillingappointmentnoText, itemonKey.InvoiceDateText, itemonKey.InvoiceNoText });
                            rowStart++;
                        }
                    }
                    rpt.SetDataSource(dt);
                    rpt.SetParameterValue("CompanyName", result.Invoice_To_Company);
                    rpt.SetParameterValue("CustCode", result.Invoice_To_Cust_Code);
                    rpt.SetParameterValue("DeliveryAddress", result.Invoice_To_Address);
                    rpt.SetParameterValue("ReceiverDocument", result.Invoice_To_Person);
                    rpt.SetParameterValue("PackingNo", result.Package_No);
                    rpt.SetParameterValue("InvoiceAddress", "Topmuju"); //groupBA.FirstOrDefault();
                    rpt.SetParameterValue("InvoiceProcess", "Topmuju"); //groupBA.FirstOrDefault();
                    return rpt.ExportToStream(ExportFormatType.PortableDocFormat);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Stream Report02(string BANO)
        {
            try
            {
                using (masterEntities context = new masterEntities())
                {
                    // Value Process
                    var result = context.TB_R_BA.Where(x => x.BA_No == BANO).FirstOrDefault();
                    if (result == null)
                    {
                        throw new Exception("BANO Not Found !");
                    }
                    var AfterJoin = (from trb in context.TB_R_BA
                                     join trp in context.TB_R_PACKAGE on trb.Package_ID equals trp.Package_ID
                                     join tri in context.TB_R_INVOICE on trb.BA_ID equals tri.BA_ID
                                     where trb.BA_No == BANO
                                     select new
                                     {
                                         trb.BA_ID,
                                         trb.BA_No,
                                         trp.Package_No,
                                         trb.Quote_No,
                                         tri.CreditTerm,
                                         trb.Report_To_Comany,
                                         trb.Reports_To_Address,
                                         trb.Invoice_To_Comany,
                                         trb.Invoice_Cust_Code,
                                         trb.Invoice_To_Person,
                                         trb.Invoice_To_Address,
                                         trb.Invoice_To_Tel,
                                         tri.Invoice_Issue_Date
                                     }).OrderBy(x => x.Package_No)
                                     .ThenBy(p => p.BA_No)
                                     .ToList();

                    var HeaderText = AfterJoin.FirstOrDefault();
                    var invoice_List = context.TB_R_INVOICE.Where(x => x.BA_ID == HeaderText.BA_ID).ToList();

                    // Generate BARCODE
                    BarcodeLib.Barcode b1 = new BarcodeLib.Barcode();
                    Image img1 = b1.Encode(BarcodeLib.TYPE.CODE128, AfterJoin.FirstOrDefault().Package_No, Color.Black, Color.White, 375, 120);
                    MemoryStream fs1 = new MemoryStream();
                    ((Bitmap)img1).Save(fs1, System.Drawing.Imaging.ImageFormat.Jpeg);
                    byte[] Barcode01 = fs1.ToArray();
                    fs1.Dispose();

                    BarcodeLib.Barcode b2 = new BarcodeLib.Barcode();
                    Image img2 = b2.Encode(BarcodeLib.TYPE.CODE128, AfterJoin.FirstOrDefault().BA_No, Color.Black, Color.White, 375, 120);
                    MemoryStream fs2 = new MemoryStream();
                    ((Bitmap)img2).Save(fs2, System.Drawing.Imaging.ImageFormat.Jpeg);
                    byte[] Barcode02 = fs2.ToArray();
                    fs2.Dispose();

                    ReportDocument rpt = new ReportDocument();
                    rpt.Load(Path.Combine(DirectoryPath, "report02.rpt"));
                    var dt = new Access_data.DatasetImg.ImageSet();
                    dt.DataTable1.Rows.Add(new Object[] { Barcode01, Barcode02 });

                    var index = 1;
                    foreach (var inviceItem in invoice_List)
                    {
                        dt.DataTable3.Rows.Add(new Object[] {
                            index,
                            inviceItem.Invoice_Issue_Date != null ? inviceItem.Invoice_Issue_Date?.ToString("dd-MM-yyyy") : "",
                            inviceItem.Invoice_No,
                            inviceItem.Total_Invoice_Amount_Inc_Vat
                        });
                        index++;
                    }


                    rpt.SetDataSource(dt);

                    rpt.SetParameterValue("QueteNo", HeaderText.Quote_No);
                    rpt.SetParameterValue("CreditTerm", HeaderText.CreditTerm);
                    rpt.SetParameterValue("ReportToName", HeaderText.Report_To_Comany);
                    rpt.SetParameterValue("ReportToAddress", HeaderText.Reports_To_Address);

                    rpt.SetParameterValue("CompanyName", HeaderText.Invoice_To_Comany);
                    rpt.SetParameterValue("CustCode", HeaderText.Invoice_Cust_Code);
                    rpt.SetParameterValue("DeliverToName", HeaderText.Invoice_To_Person);
                    rpt.SetParameterValue("DeliverToAddress", HeaderText.Invoice_To_Address);
                    rpt.SetParameterValue("DeliverToTelphone", HeaderText.Invoice_To_Tel);

                    rpt.SetParameterValue("PackingNo", HeaderText.Package_No);
                    rpt.SetParameterValue("BillingAppointmentNo", HeaderText.BA_No);
                    rpt.SetParameterValue("IssueDate", HeaderText.Invoice_Issue_Date?.ToString("MMMM dd, yyyy / HH:mm"));
                    return rpt.ExportToStream(ExportFormatType.PortableDocFormat);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
