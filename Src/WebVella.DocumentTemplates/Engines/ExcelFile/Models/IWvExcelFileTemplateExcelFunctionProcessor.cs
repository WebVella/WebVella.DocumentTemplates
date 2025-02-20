﻿using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.ExcelFile.Models;
public interface IWvExcelFileTemplateExcelFunctionProcessor
{
	public string Name { get; }
	public int Priority { get; }
	public string? FormulaA1 { get; set; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
		object? value,
		WvTemplateTag tag,
		WvExcelFileTemplateContext templateContext,
		int expandPosition,
		int expandPositionMax,
		WvExcelFileTemplateProcessResult result,
		WvExcelFileTemplateProcessResultItem resultItem,
		IXLWorksheet processedWorksheet
	);

}
