using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebVella.DocumentTemplates.Engines.Email.Models;
public class WvEmailAttachment
{
	public string? Filename { get; set; }
	public byte[]? Content { get; set; }
}
