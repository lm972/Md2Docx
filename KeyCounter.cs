using System;
using System.Collections.Generic;

namespace Md2Docx
{
	public class KeyCounter<TKey>
	{
		private Dictionary<TKey, int> _dict = new Dictionary<TKey, int>();

		public KeyCounter ()
		{
		}

		public int this[TKey key] {
			get {
				if (!_dict.ContainsKey (key))
					return 0;
				return _dict [key];
			}
			set {
				if (!_dict.ContainsKey (key))
					_dict.Add (key, 0);
				_dict [key] = value;
			}
		}
	}
}

