using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Services;
public class WvSpreadsheetFileMetaService
{
	private static List<IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor>? _spreadsheetFunctionsList = null;
	private static List<IWvSpreadsheetFileTemplateFunctionProcessor>? _functionsList = null;

	static WvSpreadsheetFileMetaService()
	{
		_spreadsheetFunctionsList = new();
		_functionsList = new();

		var assemblies = AppDomain.CurrentDomain.GetAssemblies()
							.Where(a => !(a.FullName!.ToLowerInvariant().StartsWith("microsoft.")
							   || a.FullName.ToLowerInvariant().StartsWith("system.")));

		foreach (var assembly in assemblies)
		{
			foreach (Type type in assembly.GetTypes())
			{
				if (type.IsAbstract || type.IsInterface)
					continue;

				var defaultConstructor = type.GetConstructor(Type.EmptyTypes);
				if (defaultConstructor is null)
					continue;

				if (type.ImplementsInterface(typeof(IWvSpreadsheetFileTemplateFunctionProcessor)))
				{
					IWvSpreadsheetFileTemplateFunctionProcessor? instance = Activator.CreateInstance(type) as IWvSpreadsheetFileTemplateFunctionProcessor;
					if (instance is not null)
					{
						_functionsList.Add(instance);
					}
				}

				if (type.ImplementsInterface(typeof(IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor)))
				{
					IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor? instance = Activator.CreateInstance(type) as IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor;
					if (instance is not null)
					{
						_spreadsheetFunctionsList.Add(instance);
					}
				}
			}
		}
	}

	public Type? GetFunctionProcessorByName(string name)
	{
		if (String.IsNullOrWhiteSpace(name)) return null;
		var matches = _functionsList!.Where(a => a.Name.ToLowerInvariant() == name.Trim().ToLowerInvariant());
		if (matches.Any()) return matches.OrderByDescending(x => x.Priority).First().GetType();
		return null;
	}

	public Type? GetSpreadsheetFunctionProcessorByName(string name)
	{
		if (String.IsNullOrWhiteSpace(name)) return null;
		var matches = _spreadsheetFunctionsList!.Where(a => a.Name.ToLowerInvariant() == name.Trim().ToLowerInvariant());
		if (matches.Any()) return matches.OrderByDescending(x => x.Priority).First().GetType();
		return null;
	}

}
