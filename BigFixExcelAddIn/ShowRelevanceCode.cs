using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AddinExpress.MSO;
using Ader.Text;
using QWhale.Editor.TextSource;

namespace BigFixExcelConnector
{
    public partial class ShowRelevanceCode : Form
    {
        String RelevanceWithLWIndent = "";

        public ShowRelevanceCode()
        {
            InitializeComponent();

            // Lee Wei 2013-03-25 commented out for testing
            /*
            Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
            Excel.Range relevanceCell = hiddenWorksheet.get_Range("A1", "A1");
            RelevanceWithLWIndent = relevanceCell.Value.ToString().Replace("\\n", "\n");
            syntaxEditBES.Text = RelevanceWithLWIndent;
            */

            syntaxEditBES.Braces.BracesOptions = BracesOptions.Highlight;
            syntaxEditBES.Braces.BackColor = Color.Orange;

            buttonIndentLW.Enabled = false;

        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonCopy_Click(object sender, EventArgs e)
        {
            syntaxEditBES.Selection.SelectAll();
            syntaxEditBES.Selection.Copy();
            syntaxEditBES.Selection.Clear();
        }

        private void buttonIndent_Click(object sender, EventArgs e)
        {
            syntaxEditBES.SyntaxPaint.DrawColumnsIndent = true;

            StringTokenizer tok = new StringTokenizer(syntaxEditBES.Text);
            tok.IgnoreWhiteSpace = true;
            tok.SymbolChars = new char[] { '(', ')' };

            syntaxEditBES.Text = "";

            int sp = 0;
            String space = "";
            String firstIndent = "\t";
            String oneTab = "\t";
            Boolean crlf = false;
            Boolean newline = true;
            Boolean justElse = false;
            Boolean justElse2 = false;
            Boolean justElseNoParen = false;
            String parsedResults = "";
            Boolean previousTokenIsSymbol = false;

            Token token;
            do
            {
                token = tok.Next();
                // Console.WriteLine(token.Kind.ToString() + ": " + token.Value);

                // ================================= ( ===========================================
                if (token.Kind.ToString() == "EOL")
                {
                    // nothing, kill the EOLs
                }
                else if (token.Kind.ToString() == "Symbol" && token.Value == "(")
                {
                    if (crlf == true)
                    {
                        parsedResults = parsedResults + "\r\n";
                    }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value + "\r\n";

                    crlf = false;
                    newline = true;
                    justElseNoParen = false;
                    sp++;
                }
                // ================================= ) ===========================================
                else if (token.Kind.ToString() == "Symbol" && token.Value == ")")
                {

                    if (justElse == true && justElse2 == true)
                    {
                        sp--;
                        sp--;
                        justElse = false;
                        justElse2 = false;
                    }
                    else
                    {
                        sp--;
                    }

                    if (justElse == true && justElse2 == false)
                    {
                        justElse2 = true;
                        // sp--;
                    }

                    if (justElseNoParen == true)
                    {
                        sp--;
                        justElseNoParen = false;
                        justElse = false;
                        justElse2 = false;
                    }

                    if (crlf == true)
                    {
                        parsedResults = parsedResults + "\r\n";
                    }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value + "\r\n";

                    crlf = false;
                    // crlf = true;
                    newline = true;
                }
                // ================================= then/else ==========================================
                else if (token.Kind.ToString() == "Word" && (token.Value.ToLower() == "then" ||
                                                             token.Value.ToLower() == "else"
                                                                ))
                {
                    if (justElseNoParen == true)
                    {
                        sp--;
                        justElseNoParen = false;
                        justElse = false;
                        justElse2 = false;
                    }

                    if (token.Value.ToLower() == "else")
                    {
                        justElse = true;
                        justElseNoParen = true;
                    }

                    if (crlf == true)
                    {
                        parsedResults = parsedResults + "\r\n";
                    }

                    sp--;

                    if (sp < 0) { sp = 0; }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value + "\r\n";

                    sp++;
                    crlf = false;
                    newline = true;
                }
                // ================================= and/or ===========================================
                else if (token.Kind.ToString() == "Word" && (token.Value.ToLower() == "and" ||
                                                             token.Value.ToLower() == "or"))
                {
                    if (token.Value.ToLower() == "else")
                    {
                        justElse = true;
                    }

                    if (crlf == true)
                    {
                        parsedResults = parsedResults + "\r\n";
                    }

                    // sp--;
                    // MessageBox.Show(sp.ToString());

                    // if (sp < 0) { sp = 0; }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    // space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value + "\r\n";

                    // sp++;
                    crlf = false;
                    newline = true;
                }
                // ================================= whose ===========================================
                else if (token.Kind.ToString() == "Word" && (token.Value.ToLower() == "whose"))
                {
                    if (crlf == true)
                    {
                        parsedResults = parsedResults + "\r\n";
                    }

                    // sp--;

                    if (sp < 0) { sp = 0; }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value;
                    // MessageBox.Show("crlf is: " + crlf.ToString());
                    // sp++;
                    crlf = true;
                }
                // ================================= if ===========================================
                else if (token.Kind.ToString() == "Word" && (token.Value.ToLower() == "if"))
                {
                    if (crlf == true)
                    {
                    }

                    // sp--;

                    if (sp < 0) { sp = 0; }

                    space = "";
                    for (int i = 0; i < sp; i++)
                    {
                        space = space + oneTab;
                    }
                    space = firstIndent + space;

                    parsedResults = parsedResults + space + token.Value + "\r\n";

                    sp++;
                    crlf = false;
                    justElse = false;
                    justElse2 = false;
                }
                // ================================= others ===========================================
                else
                {
                    space = "";

                    if (newline)
                    {
                        for (int i = 0; i < sp; i++)
                        {
                            space = space + oneTab;
                        }
                        space = firstIndent + space;
                    }
                    else
                    {
                        // space = firstIndent + space;
                    }

                    // Added June 22nd 2010 to fix extra space problem between >=
                    if ((token.Kind.ToString() == "Unknown") && previousTokenIsSymbol)
                    {
                        parsedResults = parsedResults.TrimEnd() + space + token.Value;
                    }
                    else
                        parsedResults = parsedResults + space + token.Value + " ";
                    // ==============================================================

                    crlf = true;
                    newline = false;

                    // Added June 22nd 2010 to fix extra space problem between >=
                    if (token.Kind.ToString() == "Unknown")
                        previousTokenIsSymbol = true;
                    else
                        previousTokenIsSymbol = false;
                    // ==============================================================
                }

            } while (token.Kind != TokenKind.EOF);

            syntaxEditBES.Text = parsedResults;
            buttonIndentLW.Enabled = true;
            buttonIndent.Enabled = false;
            buttonFlatten.Enabled = true;

        }

        private void buttonFlatten_Click(object sender, EventArgs e)
        {
            syntaxEditBES.SyntaxPaint.DrawColumnsIndent = false;

            StringTokenizer tok = new StringTokenizer(syntaxEditBES.Text);
            tok.IgnoreWhiteSpace = true;
            tok.SymbolChars = new char[] { '(', ')' };

            syntaxEditBES.Text = "";
            Token token;

            String parsedResults = "";

            Boolean previousTokenIsSymbol = false;

            do
            {
                token = tok.Next();
                // Console.WriteLine(token.Kind.ToString() + ": " + token.Value);

                // ================================= ( ===========================================
                if (token.Kind.ToString() == "EOL")
                {
                }
                else if (token.Kind.ToString() == "Symbol" && token.Value == "(")
                {
                    parsedResults = parsedResults + token.Value;
                }
                else if (token.Kind.ToString() == "Symbol" && token.Value == ")")
                {
                    parsedResults = parsedResults.TrimEnd(' ') + token.Value + " ";
                }
                else
                {
                    // Added June 22nd 2010 to fix extra space problem between >=
                    if ((token.Kind.ToString() == "Unknown") && previousTokenIsSymbol)
                    {
                        parsedResults = parsedResults.TrimEnd() + token.Value;
                    }
                    else
                        parsedResults = parsedResults + token.Value + " ";
                    // ==============================================================

                    // Added June 22nd 2010 to fix extra space problem between >=
                    if (token.Kind.ToString() == "Unknown")
                        previousTokenIsSymbol = true;
                    else
                        previousTokenIsSymbol = false;
                    // ==============================================================
                }

            } while (token.Kind != TokenKind.EOF);

            syntaxEditBES.Text = parsedResults;
            buttonIndentLW.Enabled = true;
            buttonIndent.Enabled = true;
            buttonFlatten.Enabled = false;
        }

        private void buttonIndentLW_Click(object sender, EventArgs e)
        {
            syntaxEditBES.Text = RelevanceWithLWIndent;
            buttonIndentLW.Enabled = false;
            buttonIndent.Enabled = true;
            buttonFlatten.Enabled = true;
        }
    }
}
