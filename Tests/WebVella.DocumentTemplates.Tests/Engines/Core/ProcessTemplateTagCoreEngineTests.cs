using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class ProcessTemplateTagCoreEngineTests : TestBase
{
	private static readonly object locker = new();
	public ProcessTemplateTagCoreEngineTests() : base() { }

	[Fact]
	public void TemplateProcessShouldReturnNoResultsIfEmpty()
	{
		//Given
		string template = "";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.True(String.IsNullOrWhiteSpace((string?)result.Values[0]));
	}

	[Fact]
	public void TemplateProcessShouldReturnNoResultsIfNoTag()
	{
		//Given
		string template = "sometext test";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal(template, result.Values[0]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCannotBeProcessed()
	{
		//Given
		string template = "{{}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal(template, result.Values[0]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessed()
	{
		//Given
		string template = "{{name}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["name"]?.ToString(), result.Values[0]);
		Assert.Equal(ds.Rows[1]["name"]?.ToString(), result.Values[1]);
		Assert.Equal(ds.Rows[2]["name"]?.ToString(), result.Values[2]);
		Assert.Equal(ds.Rows[3]["name"]?.ToString(), result.Values[3]);
		Assert.Equal(ds.Rows[4]["name"]?.ToString(), result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMulti()
	{
		//Given
		string template = "{{position}}.{{name}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["position"] + "." + ds.Rows[0]["name"], result.Values[0]);
		Assert.Equal(ds.Rows[1]["position"] + "." + ds.Rows[1]["name"], result.Values[1]);
		Assert.Equal(ds.Rows[2]["position"] + "." + ds.Rows[2]["name"], result.Values[2]);
		Assert.Equal(ds.Rows[3]["position"] + "." + ds.Rows[3]["name"], result.Values[3]);
		Assert.Equal(ds.Rows[4]["position"] + "." + ds.Rows[4]["name"], result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndex()
	{
		//Given
		string template = "{{position}}.{{name[0]}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["position"] + "." + ds.Rows[0]["name"], result.Values[0]);
		Assert.Equal(ds.Rows[1]["position"] + "." + ds.Rows[0]["name"], result.Values[1]);
		Assert.Equal(ds.Rows[2]["position"] + "." + ds.Rows[0]["name"], result.Values[2]);
		Assert.Equal(ds.Rows[3]["position"] + "." + ds.Rows[0]["name"], result.Values[3]);
		Assert.Equal(ds.Rows[4]["position"] + "." + ds.Rows[0]["name"], result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndexSingle1()
	{
		//Given
		string template = "{{name[0]}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal((string)ds.Rows[0]["name"], result.Values[0]);
		Assert.Single(result.Tags);
		Assert.Equal(WvTemplateTagType.Data,result.Tags[0].Type);
	}
    [Fact]
    public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndexSingle2()
    {
        //Given
        string template = "{{=name}}";
        DataTable ds = SampleData;
        //When
        WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
        //Then
        Assert.NotNull(result);
        Assert.Single(result.Tags);
        Assert.Equal(WvTemplateTagType.Function, result.Tags[0].Type);
    }
    [Fact]
    public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndexSingle3()
    {
        //Given
        string template = "{{==name}}";
        DataTable ds = SampleData;
        //When
        WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, DefaultCulture);
        //Then
        Assert.NotNull(result);
        Assert.Single(result.Tags);
        Assert.Equal(WvTemplateTagType.SpreadsheetFunction, result.Tags[0].Type);
    }


    [Fact]
	public void ProcessValueShouldBeCorrect()
	{
		//Given
		var columnTypeNames = new List<string>{
			"short","int","long","number","date","datetime",
			"shorttext","text","guid"
		};
		var culture = new CultureInfo("en-US");
		var currentCulture = Thread.CurrentThread.CurrentCulture;
		var currentUICulture = Thread.CurrentThread.CurrentUICulture;
		try
		{
			Thread.CurrentThread.CurrentCulture = culture;
			Thread.CurrentThread.CurrentCulture = culture;
			foreach (var columnName in columnTypeNames)
			{
				string template = "{{" + columnName + "[0]}}";
				DataTable ds = TypedData;
				//When
				WvTemplateTagResultList result = new WvTemplateUtility().ProcessTemplateTag(template, ds, culture);
				//Then
				Assert.NotNull(result);
				Assert.Single(result.Values);
				Assert.Equal(result.Values[0], ds.Rows[0][columnName]);
			}
		}
		finally
		{
			Thread.CurrentThread.CurrentCulture = currentCulture;
			Thread.CurrentThread.CurrentUICulture = currentUICulture;
		}
	}
}
