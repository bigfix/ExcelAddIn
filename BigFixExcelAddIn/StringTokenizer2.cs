/********************************************************8
 *	Author: Andrew Deren
 *	Date: July, 2004
 *	http://www.adersoftware.com
 * 
 *	StringTokenizer class. You can use this class in any way you want
 *  as long as this header remains in this file.
 * 
 **********************************************************/
using System;
using System.IO;
using System.Text;

namespace Ader.Text
{
    /// <summary>
    /// StringTokenizer tokenized string (or stream) into tokens.
    /// </summary>
    public class StringTokenizer2
    {
        const char EOF = (char)0;

        int line;
        int column;
        int pos;	// position within data

        string data;

        bool ignoreWhiteSpace;
        char[] symbolChars;

        int saveLine;
        int saveCol;
        int savePos;

        bool hasOpenParen = false;
        int openParenCount = 0;

        public StringTokenizer2(TextReader reader)
        {
            if (reader == null)
                throw new ArgumentNullException("reader");

            data = reader.ReadToEnd();

            Reset();
        }

        public StringTokenizer2(string data)
        {
            if (data == null)
                throw new ArgumentNullException("data");

            this.data = data;

            Reset();
        }

        /// <summary>
        /// gets or sets which characters are part of TokenKind2.Symbol
        /// </summary>
        public char[] SymbolChars
        {
            get { return this.symbolChars; }
            set { this.symbolChars = value; }
        }

        /// <summary>
        /// if set to true, white space characters will be ignored,
        /// but EOL and whitespace inside of string will still be tokenized
        /// </summary>
        public bool IgnoreWhiteSpace
        {
            get { return this.ignoreWhiteSpace; }
            set { this.ignoreWhiteSpace = value; }
        }

        private void Reset()
        {
            this.ignoreWhiteSpace = false;
            // this.symbolChars = new char[]{'=', '+', '-', '/', ',', '.', '*', '~', '!', '@', '#', '$', '%', '^', '&', '(', ')', '{', '}', '[', ']', ':', ';', '<', '>', '?', '|', '\\'};
            this.symbolChars = new char[] { };
            line = 1;
            column = 1;
            pos = 0;
        }

        protected char LA(int count)
        {
            if (pos + count >= data.Length)
                return EOF;
            else
                return data[pos + count];
        }

        protected char Consume()
        {
            char ret = data[pos];
            pos++;
            column++;

            return ret;
        }

        protected Token2 CreateToken(TokenKind2 kind, string value)
        {
            return new Token2(kind, value, line, column);
        }

        protected Token2 CreateToken(TokenKind2 kind)
        {
            string tokenData = data.Substring(savePos, pos - savePos);
            return new Token2(kind, tokenData, saveLine, saveCol);
        }

        public Token2 Next()
        {
        ReadToken:

            char ch = LA(0);
            switch (ch)
            {
                case EOF:
                    return CreateToken(TokenKind2.EOF, string.Empty);

                case ' ':
                case '\t':
                    {
                        if (this.ignoreWhiteSpace)
                        {
                            Consume();
                            goto ReadToken;
                        }
                        else
                            return ReadWhitespace();
                    }
                case '0':
                case '1':
                case '2':
                case '3':
                case '4':
                case '5':
                case '6':
                case '7':
                case '8':
                case '9':
                    return ReadString();
                //return ReadNumber();

                case '\r':
                    {
                        StartRead();
                        Consume();
                        if (LA(0) == '\n')
                            Consume();	// on DOS/Windows we have \r\n for new line

                        line++;
                        column = 1;

                        return CreateToken(TokenKind2.EOL);
                    }
                case '\n':
                    {
                        StartRead();
                        Consume();
                        line++;
                        column = 1;

                        return CreateToken(TokenKind2.EOL);
                    }

                case '"':
                    {
                        return ReadString();
                    }

                case '(':
                    {
                        hasOpenParen = true;
                        openParenCount++;
                        return ReadString();
                    }

                default:
                    {
                        return ReadString();
                    }

            }
        }

        /// <summary>
        /// save read point positions so that CreateToken can use those
        /// </summary>
        private void StartRead()
        {
            saveLine = line;
            saveCol = column;
            savePos = pos;
        }

        /// <summary>
        /// reads all whitespace characters (does not include newline)
        /// </summary>
        /// <returns></returns>
        protected Token2 ReadWhitespace()
        {
            StartRead();

            Consume(); // consume the looked-ahead whitespace char

            while (true)
            {
                char ch = LA(0);
                if (ch == '\t' || ch == ' ')
                    Consume();
                else
                    break;
            }

            return CreateToken(TokenKind2.WhiteSpace);

        }

        /// <summary>
        /// reads number. Number is: DIGIT+ ("." DIGIT*)?
        /// </summary>
        /// <returns></returns>
        protected Token2 ReadNumber()
        {
            StartRead();

            bool hadDot = false;

            Consume(); // read first digit

            while (true)
            {
                char ch = LA(0);
                if (Char.IsDigit(ch))
                    Consume();
                else if (ch == '.' && !hadDot)
                {
                    hadDot = true;
                    Consume();
                }
                else
                    break;
            }

            return CreateToken(TokenKind2.Number);
        }

        /// <summary>
        /// reads word. Word contains any alpha character or _
        /// </summary>
        protected Token2 ReadWord()
        {
            StartRead();

            Consume(); // consume first character of the word

            while (true)
            {
                char ch = LA(0);
                if (Char.IsLetter(ch) || ch == '_')
                    Consume();
                else
                    break;
            }

            return CreateToken(TokenKind2.Word);
        }

        /// <summary>
        /// reads all characters until next " is found.
        /// If "" (2 quotes) are found, then they are consumed as
        /// part of the string
        /// </summary>
        /// <returns></returns>
        protected Token2 ReadString()
        {
            StartRead();

            Consume(); // read "
            int positionCount = 0;

            while (true)
            {
                char ch = LA(0);
                if (ch == EOF)
                    break;
                else if (ch == '\r')	// handle CR in strings
                {
                    Consume();
                    if (LA(0) == '\n')	// for DOS & windows
                        Consume();

                    line++;
                    column = 1;
                }
                else if (ch == '\n')	// new line in quoted string
                {
                    Consume();

                    line++;
                    column = 1;
                }

                else if (ch == '(')	// Checks for nested open parens
                {
                    Consume();
                    hasOpenParen = true;
                    openParenCount++;
                }

                else if (ch == ')')	// Detect close parenthesis
                {
                    Consume();
                    openParenCount--;
                    if (openParenCount == 0) //Checks for nested parenthesis
                        hasOpenParen = false;
                }

                else if (ch == ',')
                {
                    if (hasOpenParen == true)
                    {
                        Consume();
                    }
                    else
                    {
                        Consume();
                        break;	// done reading, and this quotes does not have escape character
                    }
                }
                else
                    Consume();

                positionCount++;

            }

            return CreateToken(TokenKind2.QuotedString);
        }

        /// <summary>
        /// checks whether c is a symbol character.
        /// </summary>
        protected bool IsSymbol(char c)
        {
            for (int i = 0; i < symbolChars.Length; i++)
                if (symbolChars[i] == c)
                    return true;

            return false;
        }
    }
}
