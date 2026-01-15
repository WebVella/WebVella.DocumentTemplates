namespace WebVella.DocumentTemplates.Core;

public class WvTemplateTagIndexGroup
{
    public WvTemplateTagIndexGroup() { }

    public WvTemplateTagIndexGroup(List<int> indexes)
    {
        Indexes = new();
        Indexes.AddRange(indexes);
    }
    public List<int> Indexes { get; set; } = new();
}
