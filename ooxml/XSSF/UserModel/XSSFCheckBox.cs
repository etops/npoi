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

using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a text box in a SpreadsheetML Drawing.
     *
     * @author Yegor Kozlov
     */
    public class XSSFCheckBox : XSSFSimpleShape //, ITextbox
    {
        NPOI.OpenXmlFormats.Vml.CT_Shape _vmlShape;
        XSSFCtrlProp _ctrlProp;
        internal XSSFCheckBox(XSSFDrawing drawing, CT_Shape ctShape, NPOI.OpenXmlFormats.Vml.CT_Shape vmlShape, XSSFCtrlProp ctrlProp)
            : base(drawing, ctShape)
        {
            _vmlShape = vmlShape;
            _ctrlProp = ctrlProp;
        }

        public NPOI.OpenXmlFormats.Vml.CT_Shape GetVmlShape()
        {
            return _vmlShape;
        }

        // set a cell e.g. $K$10
        public void SetFmlaLink(string value)
        {
            _vmlShape.ClientData[0].fmlaLink = null;
            _vmlShape.ClientData[0].AddNewFmlaLink(value);
            _ctrlProp.fmlaLink = value;
        }

        public void SetText(string value)
        {
            _vmlShape.textbox.ItemXml = "<div style='text-align:left'>" +
                                            "<font face=\"Segoe UI\" size=\"160\" color=\"auto\">" +
                                              value +
                                            "</font>" +
                                        "</div>";
        }
        
    }
}


