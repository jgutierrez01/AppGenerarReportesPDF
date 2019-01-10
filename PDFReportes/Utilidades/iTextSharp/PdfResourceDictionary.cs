#region This file is part of the iTextSharp.extended library
//
// PdfResourceDictionary.cs
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
using iTextSharp.text.pdf;

namespace iTextSharp.text.pdf
{
	public class PdfResourceDictionary : PdfDictionary
    {
        
        #region Private variables

        private List<PdfDictionary> _resourceStack = new List<PdfDictionary>();

        #endregion

        #region Push

        public void Push(PdfDictionary resources)
        {
            _resourceStack.Add(resources);
        }

        #endregion

        #region Pop

        public void Pop()
        {
            _resourceStack.RemoveAt(_resourceStack.Count - 1);
        }

        #endregion

        #region GetDirectObject

        public override PdfObject GetDirectObject(PdfName key)
        {
            for (int index = _resourceStack.Count - 1; index >= 0; index--)
            {
                PdfDictionary subResource = _resourceStack[index];

                if (subResource != null)
                {
                    PdfObject obj = subResource.GetDirectObject(key);
                    if (obj != null)
                    {
                        return obj;
                    }
                }
            }
            return base.GetDirectObject(key);
        }

        #endregion

    }
}
