﻿using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvWordFileTemplate : WvTemplateBase
{
	public string? Template { get; set; }
	public WvWordFileTemplateResult? Process(DataTable dataSource, CultureInfo? culture = null)
	{
		return null;
	}
}