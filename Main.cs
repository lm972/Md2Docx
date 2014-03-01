using System;
using System.IO;

namespace Md2Docx
{
	class MainClass
	{
		static void Main (string[] args)
		{
			string source = "";
			string dest = "";

			if (args.Length > 0) {
				source = args [0];
				if (args.Length > 1) {
					dest = args [1];
				}
			} else {
				source = "/Users/ZTH/Projects/Md2Docx/usage.md";
			}

			if (dest == "") dest = source.Substring(0, source.LastIndexOf('.')) + ".docx";

			MdParser parser = new MdParser (File.ReadAllText (source));
			parser.RenderDocx (dest);
			System.Diagnostics.Process.Start (new System.Diagnostics.ProcessStartInfo (dest) {
				UseShellExecute = true
			});
		}
	}
}

