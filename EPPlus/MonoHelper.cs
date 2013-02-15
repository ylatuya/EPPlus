/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						         Date
 * ******************************************************************************
 * Timotheus Pokorra    Add workaround for Mono      2012-06-27
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.IO.Packaging;
using System.Globalization;

namespace OfficeOpenXml
{
    /// <summary>
    /// provide a solution for the problem that Microsoft Excel 2010 is not able to open the xlsx file produced on Mono,
    /// due to the way the relationship ID is formatted.
    /// Mono only produces numbers like 0, 1, 2, etc.
    /// Microsoft Office 2010 does not like just a number starting with 0, that Mono would produce
    /// see https://github.com/mono/mono/blob/master/mcs/class/WindowsBase/System.IO.Packaging/PackagePart.cs
    /// Open XML SDK 2.0 Productivity Tool for Microsoft Office error message: 
    /// Cannot open the file: Die ID "0" ist keine gültige XSD-ID.
    /// </summary>
    public class PackagePartForMono
    {
        private static Int32 FNextRelationshipID = 1;
        
        /// <summary>
        /// create a relationship ID that works on Mono and for Microsoft Office
        /// </summary>
        public static string NextRelationshipID
        {
            get
            {
                string result = "rID" + FNextRelationshipID.ToString();
                FNextRelationshipID++;
                return result;
            }
        }
    }

    /// <summary>
    /// workaround for wrongly implemented GetRelativeUri in Mono
    /// see https://bugzilla.xamarin.com/show_bug.cgi?id=2527
    /// Mono always returns the sourcePartUri.
    /// </summary>
    public class PackUriHelperMonoSafe
    {
        /// <summary>
        /// get the relative Uri
        /// </summary>
        public static Uri GetRelativeUri(Uri sourcePartUri, Uri targetPartUri)
        {
            Uri result = PackUriHelper.GetRelativeUri(sourcePartUri, targetPartUri);
            
            if (result.ToString().StartsWith("/"))
            {
                // bug in Mono, see https://bugzilla.xamarin.com/show_bug.cgi?id=2527
                // Mono always returns the sourcePartUri.
                string source = sourcePartUri.ToString();
                string target = targetPartUri.ToString();
    
                int countSame = 0;
    
                while (countSame < target.Length
                       && countSame < source.Length
                       && target[countSame] == source[countSame])
                {
                    countSame++;
                }
    
                // go back to the last separator
                countSame = target.Substring(0, countSame).LastIndexOf("/") + 1;
                string ResultString = target.Substring(countSame);
    
                if (countSame > 0)
                {
                    // how many directories do we need to go up from the working Directory
                    while (countSame < source.Length)
                    {
                        if (source[countSame] == '/')
                        {
                            ResultString = "../" + ResultString;
                        }
    
                        countSame++;
                    }
                }
                
                result = new Uri(ResultString, UriKind.Relative);
            }
            
            return result;
        }
    }
}
