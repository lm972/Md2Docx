using System;
using System.Text.RegularExpressions;

namespace Md2Docx
{
	/// <summary>
	/// Md 的元素
	/// </summary>
	public class MdElement {
		public MdElementType ElementType { get; set; }

		public string Text { get; set; }

		public string[] Args { get; set; }

		public int Line { get; set; }

		public int ColStart { get; set; }
		public int ColEnd { get; set; }

		public MdElement() {
			ElementType = MdElementType.Text;
			Text = null;
			Args = new string[0];
		}
	}

	/// <summary>
	/// Md element type
	/// </summary>
	public enum MdElementType {
		Text,
		Command,
		Headline,
		BlockQuote,
		Ul,
		Ol,
		Image,
		Table
	}

}

