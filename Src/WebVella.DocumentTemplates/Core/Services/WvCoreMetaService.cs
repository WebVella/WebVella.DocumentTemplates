using WebVella.DocumentTemplates.Engines.ExcelFile.Models;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core.Services;
public class WvCoreMetaService
{
	private static List<IWvTemplateTagParameterProcessorBase>? _tagParameterList = null;

	static WvCoreMetaService()
	{
		_tagParameterList = new();

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

				if (type.ImplementsInterface(typeof(IWvTemplateTagParameterProcessorBase)))
				{
					IWvTemplateTagParameterProcessorBase? instance = Activator.CreateInstance(type) as IWvTemplateTagParameterProcessorBase;
					if (instance is not null)
					{
						_tagParameterList.Add(instance);
					}
				}
			}
		}
	}

	public Type? GetRegisteredTagParameterProcessorByName(string name)
	{
		if (String.IsNullOrWhiteSpace(name)) return null;
		var matches = _tagParameterList!.Where(a => a.Name.ToLowerInvariant() == name.Trim().ToLowerInvariant());
		if (matches.Any()) return matches.OrderByDescending(x => x.Priority).First().GetType();
		return null;
	}

}
