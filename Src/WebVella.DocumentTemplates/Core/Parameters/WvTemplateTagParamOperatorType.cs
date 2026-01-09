using System.ComponentModel;

namespace WebVella.DocumentTemplates.Core;
public enum WvTemplateTagParamOperatorType
{
	[Description("")]
	Unknown = 0,
	
	[Description("=")]
	Equals = 1,
	
	[Description("!=")]
	NotEquals = 2 ,
	
	[Description(">")]
	GreaterThan = 3,
	
	[Description(">=")]
	GreaterOrEqualThan = 4,
	
	[Description("<")]
	LessThan = 5,
	
	[Description("<=")]
	LessOrEqualThan = 6,
	
	[Description("*=")]
	Contains = 7,
	
	[Description("!*=")]
	NotContains = 8,	
	
	[Description("^=")]
	StartsWith = 9,
	
	[Description("!^=")]
	NotStartsWith = 10,
	
	[Description("$=")]
	EndsWith = 11,
	
	[Description("!$=")]
	NotEndsWith = 12,		
}
