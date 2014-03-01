using System;
using System.Reflection;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;

namespace Md2Docx
{
	public class MdParser
	{
		private string[] _lines = null;

		private int _line = 0;

		private List<MdElement> _elements = new List<MdElement>();

		private Dictionary<string, string> _commands = new Dictionary<string, string>();

		private KeyCounter<string> _counters = new KeyCounter<string>();

		private Document _docx;

		private DocumentBuilder _builder;

		public MdParser (string mdtext)
		{
			// 确定换行符
			if (mdtext.Contains ("\r\n"))
				_lines = Regex.Split (mdtext, @"\r\n");
			else if (mdtext.Contains ("\r"))
				_lines = mdtext.Split ('\r');
			else
				_lines = mdtext.Split ('\n');

			// 解析成段落
			parseLines ();

			// 定义命令并扩展
			for (int i = 0; i < _elements.Count; ++i)
				_elements [i] = expandCommand (_elements [i]);

			#if DEBUG
			printElements ();
			#endif
		}

		private void commit (ref MdElement ele){
			if (ele.Text != null)
				_elements.Add (ele);
			ele = new MdElement ();
			ele.Line = _line + 1;
		}

		private void warn (string message) {
			Console.WriteLine ("Line {0}: {1}", _line, message);
		}

		private void parseLines() {
			MdElement ele = new MdElement ();
			ele.Line = 1;
			_elements = new List<MdElement> ();

			string line = null;
			for(_line = 1; _line <= _lines.Length; ++_line) {
				line = _lines [_line - 1].Trim();
				if (string.IsNullOrWhiteSpace (line)) {
					commit (ref ele);
					continue;
				}

				if (line.Length == 1) {
					ele.Text = (ele.Text ?? "") + line;
					continue;
				}

				switch (line [0]) {
				case '#': // 带编号标题
					int headlineLevel = 1;
					for (; headlineLevel < line.Length && line [headlineLevel] == '#'; ++headlineLevel)
						;

					ele.Args = new string[]{ "" + headlineLevel };
					ele.ElementType = MdElementType.Headline;
					ele.Text = line.Substring (headlineLevel).Trim();
					if (ele.Text.EndsWith ("#"))
						ele.Text = ele.Text.Substring (0, ele.Text.IndexOf ('#')).Trim ();
					commit (ref ele);
					break;
				case '-':
					if (line.All (c => c == '-')) { // 副标题
						ele.ElementType = MdElementType.Headline;
						ele.Args = new string[]{ "-1" };
						commit (ref ele);
					} else { // 无编号列举项目
						ele.ElementType = MdElementType.Ul;
						ele.Text = (ele.Text ?? "") + line.Substring (1).Trim () + "\n";
					}

					break;
				case '=':
					if (line.All (c => c == '=')) { // 主标题
						ele.ElementType = MdElementType.Headline;
						ele.Args = new string[]{ "0" };
					}
					break;
				case '>': // 引用
					ele.ElementType = MdElementType.BlockQuote;
					ele.Text = (ele.Text ?? "") + line.Substring (1).Trim ();
					break;
				case '!': // 图像或表格
					ele.ElementType = MdElementType.Image;
					var mBang = Regex.Match (line, @"\!\[(\w+)\]\((.*?)\)");
					if (mBang == null || string.IsNullOrEmpty (mBang.Groups [2].Value)) {
						warn ("Wrong arguments for `!` direction.");
						continue;
					}
					ele.Text = mBang.Groups [1].Value;
					ele.Args = mBang.Groups [2].Value.Split (' ');
					if (ele.Args [0] == "") { // 表格
						ele.ElementType = MdElementType.Table;
					} else { // 图像的话就直接提交了
						commit (ref ele);
						continue;
					}
					break;
				default: // 普通的文本
					if (line.StartsWith ("%=") || line.StartsWith ("%%")) { // 定义命令
						if (line.Length <= 2 || !line.Contains (":")) {
							warn ("Insufficient arguments for command definition.");
							continue;
						}
						string commandName = line.Substring (2);
						commandName = commandName.Substring (0, commandName.IndexOf (':'));
						string commandReplacement = line.Substring (line.IndexOf (':') + 1);
						if (_commands.ContainsKey (commandName)) {
							if (line.StartsWith ("%%"))
								_commands [commandName] = commandReplacement;
							else
								warn ("Command `" + commandName + "` already defined.");
							continue;
						}
						_commands.Add (commandName, commandReplacement);
						continue;
					}
					if (line.StartsWith ("1.")) { // Ol
						ele.ElementType = MdElementType.Ol;
						ele.Text = (ele.Text ?? "") + line.Substring (2) + "\n";
						break;
					}
					if (ele.Text == null)
						ele.Text = "";
					else if (!string.IsNullOrEmpty (ele.Text) && !ele.Text.EndsWith ("\n"))
						ele.Text += "\n";
					ele.Text += line;
					break;
				}
			}
		}

		/// <summary>
		/// 对命令进行扩展，返回 true 如果这次过程中有命令被展开或定义。
		/// </summary>
		/// <returns></returns>
		private MdElement expandCommand(MdElement ele) {
			int col = 0;

			_line = ele.Line;
			string text = null;

			var commands = findCommandElements (ele.Text).ToArray ();

			foreach (var cmd in commands) {
				if (!_commands.ContainsKey (cmd.Text))
					continue;
				if (cmd.ColStart > col + 1)
					text += ele.Text.Substring (col, cmd.ColStart - 1 - col);
				col = cmd.ColEnd + 1;
				try {
					for(int argi = 0; argi < cmd.Args.Length; ++argi) {
						MdElement tmp = new MdElement();
						tmp.ElementType = MdElementType.Text; tmp.Line = -1; tmp.Text = cmd.Args[argi];
						tmp = expandCommand(tmp);
						cmd.Args[argi] = tmp.Text;
					}
					string fmt = string.Format (_commands [cmd.Text], cmd.Args);
					text += fmt;
				} catch {
					warn ("Arguments failure while expanding command `" + cmd.Text + "`.");
				}
			}

			if (col < ele.Text.Length)
				text += ele.Text.Substring (col);

			ele.Text = text;

			return ele;
		}

		private IEnumerable<MdElement> findCommandElements(string text) {
			MdElement ele = null;
			string argstr = "";
			int balanced = 0;
			List<string> argList = new List<string>();
			for (int col = 0; col < text.Length; ++col) {
				switch (text [col]) {
				case '\\':
					++col;
					break;
				case '%': // 命令
					if (balanced == 0) {
						ele = new MdElement ();
						ele.ElementType = MdElementType.Command;
						ele.ColStart = col;
						var mName = Regex.Match (text.Substring (col), @"\%(\w+)");
						if (mName != null) {
							ele.Text = mName.Groups [1].Value;
							col += ele.Text.Length;
						}
						else
							argstr += "%";
					} else {
						argstr += '%';
					}
					break;
				case '@':
					ele = new MdElement {
						ElementType = MdElementType.Command,
						Text = "footnote"
					};
					break;
				case '^':
					ele = new MdElement {
						ElementType = MdElementType.Command,
						Text = "endnote"
					};
					break;
				case '{':
					if (ele != null) {
						++balanced;
					}
					if (balanced > 1) {
						argstr += '{';
					}
					break;
				default:
					if (ele != null) {
						if (char.IsWhiteSpace (text [col]) && balanced == 1) {
							argList.Add (argstr);
							argstr = "";
							continue;
						}
						if (text [col] == '}' && balanced > 0)
							balanced--;
						if (balanced == 0) {
							argList.Add (argstr);
							ele.Args = argList.ToArray ();
							ele.ColEnd = col;
							yield return ele;
							ele = null; argstr = ""; argList.Clear ();
						}
						else argstr += text [col];
					}
					break;
				}
			}
			if (ele != null) {
				ele.ColEnd = text.Length - 1;
				yield return ele;
			}
		}

		private void printElements() {
			foreach (var ele in _elements)
				Console.WriteLine ("{0} #{1} T={2} A={3}", ele.ElementType, ele.Line, ele.Text.Replace ("\n", @"\n"), 
					ele.Args.Length > 0 ? ele.Args.Aggregate ((p, q) => p + " " + q) : "");
		}

		public void RenderDocx(string filename) {
			_docx = new Document ();
			_docx.Styles.Add (StyleType.Paragraph, "Text");

			_builder = new DocumentBuilder (_docx);
			_builder.Font.NameFarEast = "宋体";
			buildCurrentStyle ("");
			Console.WriteLine (_commands ["style-"]);
			foreach (var ele in _elements) {
				renderElement (ele);
			}
			_docx.Save (filename);

			var docx2 = Novacode.DocX.Load (filename);
			docx2.Paragraphs.First().Remove (false);
			docx2.Save ();
		}

		private void renderText(string txt) {
			var currnode = _builder.CurrentNode;
			var cmds = findCommandElements (txt).ToList();
			cmds.Add (new MdElement {
				ElementType = MdElementType.Command,
				Text = "$",
				ColStart = txt.Length,
				ColEnd = txt.Length
			});
			int col = 0;
			foreach (var cmd in cmds) {
				for (; col < cmd.ColStart; ++col) {
					char c = txt [col];
					switch (c) {
					case '\\':
						col++;
						if (col < txt.Length)
							_builder.Write ("" + txt [col]);
						break;
					case '_':
						_builder.Underline = _builder.Underline == Underline.None ? Underline.Single : Underline.None;
						break;
					case '*':
						_builder.Bold = !_builder.Bold;
						break;
					case '~':
						_builder.Font.StrikeThrough = !_builder.Font.StrikeThrough;
						break;
					case '\n':
						break;
					default:
						_builder.Write ("" + c);
						break;
					}
				}
				switch (cmd.Text) {
				case "":
					_builder.Write (cmd.Args.Aggregate ((a, b) => a + " " + b).Trim ());
					break;
				case "$":
					// 行末指示符
					break;
				case "C":
				case "counter":
					if (cmd.Args.Length == 0) {
						warn ("Command `counter` needs more than 0 arguments.");
					} else {
						++_counters [cmd.Args [0]];
						_builder.Write ("" + _counters [cmd.Args [0]]);
					}
					break;
				case "concat":
					if (cmd.Args.Length > 0)
							_builder.Write(cmd.Args.Aggregate((a,b)=>a+b));
					break;
				case "center":
					_builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
					break;
				case "justified":
					_builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
					break;
				case "right":
					_builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
					break;
				case "header":
					currnode = _builder.CurrentNode;
					_builder.MoveToHeaderFooter (HeaderFooterType.HeaderFirst);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder.MoveToHeaderFooter (HeaderFooterType.HeaderEven);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder.MoveToHeaderFooter (HeaderFooterType.HeaderPrimary);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder = new DocumentBuilder (_docx);
					break;
				case "footer":
					currnode = _builder.CurrentNode;
					_builder.MoveToHeaderFooter (HeaderFooterType.FooterFirst);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder.MoveToHeaderFooter (HeaderFooterType.FooterEven);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder.MoveToHeaderFooter (HeaderFooterType.FooterPrimary);
					foreach (var arg in cmd.Args)
						renderText (arg + " ");
					_builder = new DocumentBuilder (_docx);
					break;
				case "wordfield":
					if (cmd.Args.Length > 1)
						_builder.InsertField (cmd.Args [0].ToUpper(), cmd.Args[1].ToUpper());
					break;
				case "ref":
					if (cmd.Args.Length > 0)
						_builder.Write ("" + _counters [cmd.Args [0]]);
					break;
				case "style":
					if (cmd.Args.Length > 0)
						applyStyle (cmd.Args [0]);
					else
						applyStyle ("");
				case "pagenum":
					_builder.InsertField ("PAGE", "");
					break;
				case "numpages":
					_builder.InsertField ("NUMPAGES", "");
					break;
				case "footnote":
					_builder.InsertFootnote (FootnoteType.Footnote,
						cmd.Args.Aggregate ((a, b) => a + " " + b));
					break;
				case "endnote":
					_builder.InsertFootnote (FootnoteType.Endnote,
						cmd.Args.Aggregate ((a, b) => a + " " + b));
					break;
				default:
					warn ("Unrecognized command `" + cmd.Text + "`.");
					break;
				}
				col = cmd.ColEnd + 1;
			}
		}

		private void renderElement(MdElement ele) {
			applyStyle ("");
			_line = ele.Line;
			try {
				switch (ele.ElementType) {
				case MdElementType.BlockQuote:
					_builder.ParagraphFormat.LeftIndent = 10;
					_builder.ParagraphFormat.SpaceBefore = 10;
					_builder.ParagraphFormat.SpaceAfter = 10;
					applyStyle("blockquote");
					renderText (ele.Text);
					break;
				case MdElementType.Headline:
					applyStyle("headline");
					_builder.ParagraphFormat.SpaceBefore = 10;
					if (int.Parse (ele.Args [0]) <= 0) {
						_builder.Font.Size = ele.Args[0] == "0" ? 24 : 18;
						_builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
					} else {
						try {
							_builder.ParagraphFormat.OutlineLevel = (OutlineLevel)Enum.Parse(typeof(OutlineLevel), "Level" + ele.Args[0]);
						}catch {}
						_builder.Font.Bold = true;
					}
					applyStyle("headline" + ele.Args[0]);
					renderText (ele.Text);
					break;
				case MdElementType.Image:
					if (ele.Args.Length == 0) {
						warn("Image instructive needs more than 0 arguments.");
					} else {
						var img = System.Drawing.Image.FromFile(ele.Args[0]);
						applyStyle("image");
						_builder.InsertImage(img, img.Width, img.Height);
						_builder.InsertParagraph();
						if (ele.Args.Length > 1) {
							_counters[ele.Args[1]]++;
							_counters[ele.Text] = _counters[ele.Args[1]];
							applyStyle("caption");
							_builder.Write(ele.Args[1] + _counters[ele.Args[1]]);
							_builder.Write(ele.Args.Length > 2 ? " " + ele.Args[2] : "");
						}
					}
					break;
				case MdElementType.Text:
					renderText (ele.Text);
					break;
				case MdElementType.Command:
					warn ("Skipped command `" + ele.Text + "`; invalid position.");
					break;
				case MdElementType.Ol:
					_builder.CurrentParagraph.ListFormat.List = _docx.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);
					var olitems = ele.Text.Split('\n');
					for(int i =0; i < olitems.Length - 1; ++i) {
						_builder.Write(olitems[i]);
						_builder.InsertParagraph();
					}
					_builder.CurrentParagraph.ListFormat.List = null;
					break;
				case MdElementType.Ul:
					_builder.CurrentParagraph.ListFormat.List = _docx.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletDefault);
					var ulitems = ele.Text.Split('\n');
					for(int i =0; i < ulitems.Length - 1; ++i) {
						_builder.Write(ulitems[i]);
							_builder.InsertParagraph();
					}
					_builder.CurrentParagraph.ListFormat.List = null;
					break;
				case MdElementType.Table:
					_builder.CellFormat.Borders.LineStyle = LineStyle.Single;

					if (ele.Args.Length > 2) {
						_counters[ele.Text] = ++_counters[ele.Args[1]];
						applyStyle("caption");
						_builder.Write(ele.Args[1] + _counters[ele.Args[1]]);
						_builder.Write(ele.Args.Length > 3 ? " " + ele.Args[2] : "");
						_builder.InsertParagraph();
					}

					applyStyle(""); applyStyle("table");

					_builder.StartTable();

					string[] tableRows = ele.Text.Split('\n');

					//Insert some table
					for (int i = 1; i < tableRows.Length - 1; i++)
					{
						string row = tableRows[i];
						foreach(var vp in Regex.Split(row, @"\t+"))
						{
							_builder.InsertCell();
							_builder.Write(vp);
						}
						_builder.EndRow();
					}
					_builder.EndTable();
					break;
				}
			} catch (Exception e) {
				warn (e.Message);
			}

			_builder.InsertParagraph ();
		}

		private string getSettings(object o, string prefix = "") {
			string val = "";
			var props = o.GetType ().GetProperties ();
			foreach (var p in props) {
				if (p.PropertyType.Equals(typeof(bool)) ||
					p.PropertyType.Equals(typeof(int)) ||
					p.PropertyType.Equals(typeof(double)) ||
					p.PropertyType.Equals(typeof(string)) ||
					p.PropertyType.IsEnum)
					val += prefix + p.Name + "=" + p.GetValue (o, null) + ",";
			}
			return val;
		}

		private void buildCurrentStyle(string name) {
			if (!_commands.ContainsKey ("style-" + name))
				_commands.Add ("style-" + name, "");
			string val = getSettings (_builder.ParagraphFormat, "PF:") + getSettings (_builder.Font, "FT:");

			_commands ["style-" + name] = val;
		}

		private void applyStyle(string name) {
			if (!_commands.ContainsKey ("style-" + name))
				return;

			string[] st = _commands ["style-" + name].Split(',');
			foreach (string s in st) {
				try {
					string[] kv = s.Split ('=');
					switch (kv [0]) {
					case "bold":
						_builder.Bold = true;
						break;
					case "italic":
						_builder.Italic = true;
						break;
					case "size":
						_builder.Font.Size = double.Parse (kv [1]);
						break;
					case "font":
						_builder.Font.Name = kv [1];
						break;
					case "l-indent":
						_builder.ParagraphFormat.LeftIndent = double.Parse(kv[1]);
						break;
					case "r-indent":
						_builder.ParagraphFormat.RightIndent = double.Parse(kv[1]);
						break;
					case "before":
						_builder.ParagraphFormat.SpaceBefore = double.Parse(kv[1]);
						break;
					case "after":
						_builder.ParagraphFormat.SpaceAfter = double.Parse(kv[1]);
						break;
					case "indent":
						_builder.ParagraphFormat.FirstLineIndent = double.Parse(kv[1]);
						break;
					case "underline":
						_builder.Underline = Underline.Single;
						break;
					case "center":
						_builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
						break;
					case "justified":
					case "left":
						_builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
						break;
					case "right":
						_builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
						break;
					default:
						PropertyInfo pty = null;
						object obj = null;
						if (kv[0].StartsWith("PF:")) {
							pty = _builder.ParagraphFormat.GetType().GetProperty(kv[0].Substring(3));
							obj =_builder.ParagraphFormat;
						} else {
							pty = _builder.Font.GetType().GetProperty(kv[0].Substring(3));
							obj = _builder.Font;
						}
						if (pty == null) continue;
						object value = null;
						if (pty.PropertyType.IsEnum)
							value = Enum.Parse(pty.PropertyType, kv[1]);
						else if (pty.PropertyType.Equals(typeof(bool)))
							value = kv[1] == "True";
						else if (pty.PropertyType.Equals(typeof(int)))
							value = int.Parse(kv[1]);
						else if (pty.PropertyType.Equals(typeof(double)))
							value = double.Parse(kv[1]);
						else value = kv[1];
						pty.SetValue(
							obj,
							value,
							null);
						break;
					}
				} catch {
				}
			}
		}
	}
}

