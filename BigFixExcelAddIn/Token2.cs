/********************************************************
 *	Author: Andrew Deren
 *	Date: July, 2004
 *	http://www.adersoftware.com
 * 
 *	StringTokenizer class. You can use this class in any way you want
 * as long as this header remains in this file.
 * 
 **********************************************************/

using System;

namespace Ader.Text
{
	public enum TokenKind2
	{
		Unknown,
		Word,
		Number,
		QuotedString,
		WhiteSpace,
		Symbol,
		EOL,
		EOF
	}

	public class Token2
	{
		int line;
		int column;
		string value;
		TokenKind2 kind;

		public Token2(TokenKind2 kind, string value, int line, int column)
		{
			this.kind = kind;
			this.value = value;
			this.line = line;
			this.column = column;
		}

		public int Column
		{
			get { return this.column; }
		}

		public TokenKind2 Kind
		{
			get { return this.kind; }
		}

		public int Line
		{
			get { return this.line; }
		}

		public string Value
		{
			get { return this.value; }
		}
	}

}
