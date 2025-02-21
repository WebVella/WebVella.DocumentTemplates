using WebVella.DocumentTemplates.Engines.ExcelFile.Models;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.ExcelFile.Services;
public class WvExcelFileMetaService
{
	private static List<IWvExcelFileTemplateExcelFunctionProcessor>? _excelFunctionsList = null;
	private static List<IWvExcelFileTemplateFunctionProcessor>? _functionsList = null;

	static WvExcelFileMetaService()
	{
		_excelFunctionsList = new();
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

				if (type.ImplementsInterface(typeof(IWvExcelFileTemplateFunctionProcessor)))
				{
					IWvExcelFileTemplateFunctionProcessor? instance = Activator.CreateInstance(type) as IWvExcelFileTemplateFunctionProcessor;
					if (instance is not null)
					{
						_functionsList.Add(instance);
					}
				}

				if (type.ImplementsInterface(typeof(IWvExcelFileTemplateExcelFunctionProcessor)))
				{
					IWvExcelFileTemplateExcelFunctionProcessor? instance = Activator.CreateInstance(type) as IWvExcelFileTemplateExcelFunctionProcessor;
					if (instance is not null)
					{
						_excelFunctionsList.Add(instance);
					}
				}
			}
		}
	}

	public List<IWvExcelFileTemplateFunctionProcessor> GetRegisteredFunctionProcessors()
	{
		return _functionsList!;
	}

	public IWvExcelFileTemplateFunctionProcessor? GetFunctionProcessorByName(string name)
	{
		if (String.IsNullOrWhiteSpace(name)) return null;
		var matches = _functionsList!.Where(a => a.Name.ToLowerInvariant() == name.Trim().ToLowerInvariant());
		if (matches.Any()) return matches.OrderByDescending(x => x.Priority).First();
		return null;
	}

	public List<IWvExcelFileTemplateExcelFunctionProcessor> GetRegisteredExcelFunctionProcessors()
	{
		return _excelFunctionsList!;
	}

	public IWvExcelFileTemplateExcelFunctionProcessor? GetExcelFunctionProcessorByName(string name)
	{
		if (String.IsNullOrWhiteSpace(name)) return null;
		var matches = _excelFunctionsList!.Where(a => a.Name.ToLowerInvariant() == name.Trim().ToLowerInvariant());
		if (matches.Any()) return matches.OrderByDescending(x => x.Priority).First();
		return null;
	}

}
