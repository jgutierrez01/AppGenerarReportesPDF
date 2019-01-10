#region This file is part of the iTextSharp.extended library
//
// PdfContentOperatorHandler.cs
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
using System.Text;

namespace iTextSharp.text.pdf.modifier
{
	public delegate List<PdfObject> PdfContentOperatorHandler(PdfLiteral oper, List<PdfObject> operands);
}
