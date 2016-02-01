/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System.Collections.Generic;
using System;
using NPOI.XSSF.Util;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using System.Text.RegularExpressions;
using NPOI.OpenXmlFormats.Vml;
using System.Xml.Serialization;
using System.Xml;
using System.Collections;
using NPOI.OpenXmlFormats.Vml.Office;
using NPOI.OpenXmlFormats.Vml.Spreadsheet;
using System.Text;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;

namespace NPOI.XSSF.UserModel
{


    /**
     * Represents a SpreadsheetML VML Drawing.
     *
     * <p>
     * In Excel 2007 VML Drawings are used to describe properties of cell comments,
     * although the spec says that VML is deprecated:
     * </p>
     * <p>
     * The VML format is a legacy format originally introduced with Office 2000 and is included and fully defined
     * in this Standard for backwards compatibility reasons. The DrawingML format is a newer and richer format
     * Created with the goal of eventually replacing any uses of VML in the Office Open XML formats. VML should be
     * considered a deprecated format included in Office Open XML for legacy reasons only and new applications that
     * need a file format for Drawings are strongly encouraged to use preferentially DrawingML
     * </p>
     * 
     * <p>
     * Warning - Excel is known to Put invalid XML into these files!
     *  For example, &gt;br&lt; without being closed or escaped crops up.
     * </p>
     *
     * See 6.4 VML - SpreadsheetML Drawing in Office Open XML Part 4 - Markup Language Reference.pdf
     *
     * @author Yegor Kozlov
     */
    public class XSSFCtrlProp : POIXMLDocumentPart
    {
        private string fmlaLinkField;
        /**
         * Create a new SpreadsheetML Drawing
         *
         * @see XSSFSheet#CreateDrawingPatriarch()
         */
        public XSSFCtrlProp()
            : base()
        {

        }

        internal void Read(Stream is1)
        {
            
        }

        internal void Write(Stream out1)
        {
            using (StreamWriter sw = new StreamWriter(out1))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sw.Write(" <formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" ");
                if(fmlaLink != null){
                    sw.Write(String.Format("fmlaLink=\"{0}\"", fmlaLink));
                }
                sw.Write(" lockText=\"1\" noThreeD=\"1\"/>");

            }
        }

        public string fmlaLink
        {
            get { return this.fmlaLinkField; }
            set { this.fmlaLinkField = value; }
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            Write(out1);
            out1.Close();
        }

    }
}

