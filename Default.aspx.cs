using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ApexExercise {
  public partial class _Default : Page {
    private const string dateFormat = "yyyy-MM-dd";
    private const int count = 15;

    private AdventureWorks2008DataContext dataContext;
  
    protected void Page_Load(object sender, EventArgs e) {
      dataContext = new AdventureWorks2008DataContext();

      export.Click += export_Click;
      reset.Click += reset_Click;
      submit.Click += submit_Click;

      if (!Page.IsPostBack) {
        setDates();
        startDate.Focus();
      }
    }

    void export_Click(object sender, EventArgs e) {
      var pkg = new ExcelPackage();
      var wbk = pkg.Workbook;
      var sheet = wbk.Worksheets.Add("Invoice Data");

      var normalStyle = "Normal";
      var acctStyle = wbk.CreateAccountingFormat();

      var data = from soh in dataContext.SalesOrderHeaders
                 join sod in dataContext.SalesOrderDetails on soh.SalesOrderID equals sod.SalesOrderID
                 join p   in dataContext.Products on sod.ProductID equals p.ProductID
                 join pod in dataContext.PurchaseOrderDetails on p.ProductID equals pod.ProductID
                 join poh in dataContext.PurchaseOrderHeaders on pod.PurchaseOrderID equals poh.PurchaseOrderID
                 where soh.DueDate > DateTime.Parse(startDate.Text) && soh.DueDate < DateTime.Parse(endDate.Text)
                 select new CustomerInvoice { InvoiceNumber = poh.PurchaseOrderID.ToString(),  InvoiceTotal = poh.TotalDue };
      
      var columns = new []
      {
        new Column { Title = "Invoice Number", Style = normalStyle, Action = i => i.InvoiceNumber, },
        new Column { Title = "Invoice Total", Style = acctStyle, Action = i => i.InvoiceTotal, TotalAction = () => data.Sum(x=>x.InvoiceTotal), },
      };
      
      sheet.SaveData(columns, data.Take(count));
      
      var bytes = pkg.GetAsByteArray();
      File.WriteAllBytes(Path.Combine(Environment.GetEnvironmentVariable("TEMP"), "ApexExample.xlsx"), bytes);
    }

    void submit_Click(object sender, EventArgs e) {
      var q = from soh in dataContext.SalesOrderHeaders
              join c in dataContext.Customers on soh.CustomerID equals c.CustomerID
              join p in dataContext.Persons on c.CustomerID equals p.BusinessEntityID
              join sod in dataContext.SalesOrderDetails on soh.SalesOrderID equals sod.SalesOrderID
              join pr in dataContext.Products on sod.ProductID equals pr.ProductID
              // join s in dataContext.Stores on c.StoreID equals s.BusinessEntityID
              join pod in dataContext.PurchaseOrderDetails on pr.ProductID equals pod.ProductID
              join poh in dataContext.PurchaseOrderHeaders on pod.PurchaseOrderID equals poh.PurchaseOrderID
              where soh.DueDate > DateTime.Parse(startDate.Text) && soh.DueDate < DateTime.Parse(endDate.Text)
              select new { /* s.Name */ c.StoreID, p.FirstName, p.LastName, soh.AccountNumber, poh.PurchaseOrderID, soh.PurchaseOrderNumber, soh.OrderDate, soh.DueDate, poh.TotalDue, pr.ProductNumber, sod.OrderQty, sod.UnitPrice, sod.LineTotal };

      var sb = new System.Text.StringBuilder();
      sb.Append("<table>");
      sb.Append("<tr><th>Sold At</th><th>Sold To</th><th>Account Number</th><th>Invoice #</th><th>Customer PO #</th><th>Order Date</th><th>Due Date</th><th>Invoice Total</th><th>Product Number</th><th>Order Qty</th><th>Unit Net</th><th>Line Total</th></tr>");

      foreach (var r in q.Take(count))
        sb.Append(String.Format("<tr><td>{0}</td><td>{1} {2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6:d}</td><td>{7:d}</td><td>{8:C2}</td><td>{9}</td><td>{10}</td><td>{11:C2}</td><td>{12:C2}</td></tr>", r.StoreID, r.FirstName, r.LastName, r.AccountNumber, r.PurchaseOrderID, r.PurchaseOrderNumber, r.OrderDate, r.DueDate, r.TotalDue, r.ProductNumber, r.OrderQty, r.UnitPrice, r.LineTotal));

      sb.Append("</table>");

      results.Text = sb.ToString();
      results.Visible = true;
    }

    void reset_Click(object sender, EventArgs e) {
      setDates();
      results.Visible = false;
      startDate.Focus();
    }

    private void setDates() {
      var lastMonth = DateTime.Now.AddMonths(-1);
      startDate.Text = firstDay(lastMonth).ToString(dateFormat);
      endDate.Text = lastDay(lastMonth).ToString(dateFormat);
    }

    private DateTime lastDay(DateTime d) {
      var nextMonth = d.AddMonths(1);
      return firstDay(nextMonth).AddDays(-1d);
    }

    private DateTime firstDay(DateTime d) {
      var y = d.Year;
      var m = d.Month;
      return new DateTime(y, m, 1);
    }
  }
}