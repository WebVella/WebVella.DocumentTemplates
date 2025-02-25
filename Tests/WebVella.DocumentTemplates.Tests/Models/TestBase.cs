using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Engines.ExcelFile;

namespace WebVella.DocumentTemplates.Tests.Models;
public class TestBase
{
	public DataTable SampleData;
	public DataTable TypedData;
	public DataTable EmailData;
	public CultureInfo DefaultCulture = new CultureInfo("en-US");
	public TestBase()
	{
		//Data
		{
			var ds = new DataTable();
			ds.Columns.Add("position", typeof(int));
			ds.Columns.Add("sku", typeof(string));
			ds.Columns.Add("name", typeof(string));
			ds.Columns.Add("price", typeof(decimal));

			for (int i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["position"] = position;
				dsrow["sku"] = $"sku{position}";
				dsrow["name"] = $"item{position}";
				dsrow["price"] = (decimal)(position * 0.33);
				ds.Rows.Add(dsrow);
			}
			SampleData = ds;
		}
		//TypedData
		{
			var ds = new DataTable();
			ds.Columns.Add("short", typeof(short));
			ds.Columns.Add("int", typeof(int));
			ds.Columns.Add("long", typeof(long));
			ds.Columns.Add("number", typeof(decimal));
			ds.Columns.Add("date", typeof(DateOnly));
			ds.Columns.Add("datetime", typeof(DateTime));
			ds.Columns.Add("shorttext", typeof(string));
			ds.Columns.Add("text", typeof(string));
			ds.Columns.Add("guid", typeof(Guid));
			for (short i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["short"] = (short)position;
				dsrow["int"] = (int)(position + 100);
				dsrow["long"] = (long)(position + 1000);
				dsrow["number"] = (decimal)(position * 10);
				dsrow["date"] = DateOnly.FromDateTime(DateTime.Now.AddDays(i));
				dsrow["datetime"] = DateTime.Now.AddDays(i + 10);
				dsrow["shorttext"] = $"short text {i}";
				dsrow["text"] = $"text {i}";
				dsrow["guid"] = Guid.NewGuid();
				ds.Rows.Add(dsrow);
			}
			TypedData = ds;
		}
		//Email Data
		{
			var ds = new DataTable();
			ds.Columns.Add("sender_email", typeof(string));
			ds.Columns.Add("recipient_email", typeof(string));
			ds.Columns.Add("subject", typeof(string));
			for (short i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["sender_email"] = $"sender{position}@test.com";
				dsrow["recipient_email"] = $"recipient{position}@test.com";
				dsrow["subject"] = $"subject{position}";

				ds.Rows.Add(dsrow);
			}

			EmailData = ds;

		}
	}

}
