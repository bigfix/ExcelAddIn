using System;
using System.IO;
using System.Text;
using System.Web.Services.Protocols;
using System.Xml;
using BigFixExcelConnector;

namespace BigFixExcelConnector
{
    internal class RelevanceBindingEx : RelevanceService
    {
        protected override XmlReader GetReaderForMessage(SoapClientMessage message, int bufferSize)
        {
            var reader = new XmlTextReader(new XmlSanitizingStreamReader(message.Stream))
                             {
                                 ProhibitDtd = true,
                                 Normalization = true,
                                 XmlResolver = null
                             };
            return reader;
        }

        internal class XmlSanitizingStreamReader : StreamReader
        {
            private const int Eof = -1;

            public XmlSanitizingStreamReader(Stream streamToSanitize) : base(streamToSanitize, true)
            {
            }

            public static bool IsLegalXmlChar(int character)
            {
                return (
                           character == 0x9 ||
                           character == 0xA ||
                           character == 0xD ||
                           (character >= 0x20 && character <= 0xD7FF) ||
                           (character >= 0xE000 && character <= 0xFFFD) ||
                           (character >= 0x10000 && character <= 0x10FFFF)
                       );
            }

            public override int Read()
            {
                // read each char, skipping ones XML has prohibited
                int nextCharacter;

                do
                {
                    // read a character
                    if ((nextCharacter = base.Read()) == Eof)
                    {
                        // end of file
                        break;
                    }
                } while (!IsLegalXmlChar(nextCharacter));

                return nextCharacter;
            }

            public override int Read(char[] buffer, int index, int count)
            {
                if (buffer == null)
                {
                    throw new ArgumentNullException("buffer", @"buffer is null");
                }
                if (index < 0)
                {
                    throw new ArgumentOutOfRangeException("index", @"index is out of bounds");
                }
                if (count < 0)
                {
                    throw new ArgumentOutOfRangeException("count", @"count cannot be zero");
                }
                if ((buffer.Length - index) < count)
                {
                    throw new ArgumentException("Invalid offset.");
                }

                var num = 0;
                do
                {
                    var num2 = Read();
                    if (num2 == -1)
                    {
                        return num;
                    }
                    buffer[index + num++] = (char) num2;
                } while (num < count);
                return num;
            }

            public override int ReadBlock(char[] buffer, int index, int count)
            {
                int num;
                var num2 = 0;
                do
                {
                    num2 += num = Read(buffer, index + num2, count - num2);
                } while ((num > 0) && (num2 < count));
                return num2;
            }

            public override string ReadLine()
            {
                var builder = new StringBuilder();
                while (true)
                {
                    var num = Read();
                    switch (num)
                    {
                        case -1:
                            if (builder.Length > 0)
                            {
                                return builder.ToString();
                            }
                            return null;

                        case 13:
                        case 10:
                            if ((num == 13) && (Peek() == 10))
                            {
                                Read();
                            }
                            return builder.ToString();
                    }
                    builder.Append((char) num);
                }
            }

            public override string ReadToEnd()
            {
                int num;
                var buffer = new char[0x1000];
                var builder = new StringBuilder(0x1000);
                while ((num = Read(buffer, 0, buffer.Length)) != 0)
                {
                    builder.Append(buffer, 0, num);
                }
                return builder.ToString();
            }

            public override int Peek()
            {
                int nextCharacter;

                do
                {
                    nextCharacter = base.Peek();
                } while (!IsLegalXmlChar(nextCharacter) && (nextCharacter = base.Read()) != Eof);

                return nextCharacter;
            }
        }
    }
}