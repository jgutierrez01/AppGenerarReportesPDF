#region This file is part of the iTextSharp.extended library
//
// PdfContentModifier.cs
//
// Author: Josip Habjan (habjan@gmail.com, http://www.linkedin.com/in/habjan) 
// Copyright (c) 2013 Josip Habjan. All rights reserved.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>. 
//
#endregion

using System;
using System.Collections.Generic;
using iTextSharp.text.pdf.parser;

namespace iTextSharp.text.pdf.modifier
{
	public class PdfContentModifier
    {

        #region Private variables

        private Dictionary<string, PdfContentOperatorHandler> _operators = new Dictionary<string, PdfContentOperatorHandler>();
        private Stack<PdfContentStreamBuilder> _contentStreamBuilderStack = new Stack<PdfContentStreamBuilder>();
        private PdfResourceDictionary _resourceDictionaryStack = new PdfResourceDictionary();

        #endregion

        #region Modify

        public void Modify(PdfReader reader, int pageNumber)
        {
            PdfDictionary pageDictionary = reader.GetPageN(pageNumber);
            PdfDictionary resourcesDictionary = pageDictionary.GetAsDict(PdfName.RESOURCES);

            byte[] contentBytes = reader.GetPageContent(pageNumber);

            contentBytes = this.Modify(contentBytes, resourcesDictionary);

            reader.SetPageContent(pageNumber, contentBytes);
        }

        #endregion

        #region Modify

        public byte[] Modify(byte[] contentBytes, PdfDictionary resourcesDictionary)
        {
            _contentStreamBuilderStack.Push(new PdfContentStreamBuilder());
            _resourceDictionaryStack.Push(resourcesDictionary);
            PRTokeniser tokeniser = new PRTokeniser(new RandomAccessFileOrArray(contentBytes));
            PdfContentParser ps = new PdfContentParser(tokeniser);

            List<PdfObject> operands = new List<PdfObject>();

            while (ps.Parse(operands).Count > 0)
            {
                PdfLiteral oper = (PdfLiteral)operands[operands.Count - 1];

                PdfContentOperatorHandler operHandler = null;

                if (_operators.TryGetValue(oper.ToString(), out operHandler))
                {
                    operands = operHandler(oper, operands);
                }

                _contentStreamBuilderStack.Peek().Push(operands);
            }

            _resourceDictionaryStack.Pop();
            return _contentStreamBuilderStack.Pop().GetBytes();
        }

        #endregion

        #region RegisterOperator

        public void RegisterOperator(string oper, PdfContentOperatorHandler callback)
        {
            _operators.Add(oper, callback);
        }

        #endregion

        #region ResourceDictionary

        public PdfResourceDictionary ResourceDictionary
        {
            get { return _resourceDictionaryStack; }
        }

        #endregion

    }

}
