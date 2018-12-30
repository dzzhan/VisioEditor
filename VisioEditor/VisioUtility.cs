// VisioUtility.cs
// compile with: /doc:VisioUtility.xml
// <copyright>Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>This file contains utility classes and methods that perform
// operations commonly performed on Visio objects.</summary>

using System;
using System.Diagnostics;
using System.Resources;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using System.Diagnostics.CodeAnalysis;

namespace VisioEditor
{

    /// <summary>The Utility class contains various utility methods.</summary>
    public sealed class Utility {

        /// <summary>Create a global resource manager for the application
        /// to retrieve localized strings.</summary>
        private static ResourceManager theResourceManager = new ResourceManager(
            Strings.ResourceFilename,
            System.Reflection.Assembly.GetExecutingAssembly());

        private Utility() {
        }

        /// <summary>The MapVisioToWindows method converts Visio coordinates to
        /// Windows coordinates.</summary>
        /// <remarks>The conversion includes changing from Visio's coordinate
        /// system (inches) to the Windows coordinate system (pixels). The
        /// conversion of the Y coordinate must also take into the account the
        /// different locations of the origin, since Visio's origin is at the
        /// lower-left corner and the Windows origin is at the upper-left
        /// corner.</remarks>
        /// <param name="drawingControl">Drawing control with the Visio window
        ///  to use.</param>
        /// <param name="visioX">X position in the Visio coordinate system.
        /// </param>
        /// <param name="visioY">Y position in the Visio coordinate system.
        /// </param>
        /// <returns>Point containing the given Visio coordinate in Windows
        /// coordinates.</returns>
        [CLSCompliant(false)]
        public static System.Drawing.Point MapVisioToWindows(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            double visioX,
            double visioY) {

            // The drawing control object must be valid.
            if (drawingControl == null) {

                // Throw a meaningful error.
                throw new ArgumentNullException("drawingControl");
            }

            int windowsX = 0;
            int windowsY = 0;
            double visioLeft;
            double visioTop;
            double visioWidth;
            double visioHeight;
            int pixelLeft;
            int pixelTop;
            int pixelWidth;
            int pixelHeight;
            Window referenceWindow;

            referenceWindow = (Window)drawingControl.Window;
            if (referenceWindow == null) {
                return System.Drawing.Point.Empty;
            }
            // Get the window coordinates in Visio units.
            referenceWindow.GetViewRect(out visioLeft, out visioTop,
                out visioWidth, out visioHeight);

            // Get the window coordinates in pixels.
            referenceWindow.GetWindowRect(out pixelLeft, out pixelTop,
                out pixelWidth, out pixelHeight);

            // Convert the X coordinate by using pixels per inch from the
            // width values.
            windowsX = (int)(pixelLeft +
                ((pixelWidth / visioWidth) * (visioX - visioLeft)));

            // Convert the Y coordinate by using pixels per inch from the
            // height values and transform from a top-left origin (windows
            // coordinates) to a bottom-left origin (Visio coordinates).
            windowsY = (int)(pixelTop +
                ((pixelHeight / visioHeight) * (visioTop - visioY)));

            return new System.Drawing.Point(windowsX, windowsY);        
        }

        /// <summary>The GetClickedShape method finds a shape at the specified
        /// location within a default tolerance.</summary>
        /// <param name="drawingControl">Drawing control with the Visio page
        ///  containing the location.</param>
        /// <param name="clickLocationX">The X coordinate of the location in
        /// Visio page units (inches).</param>
        /// <param name="clickLocationY">The Y coordinate of the location in
        /// Visio page units (inches).</param>
        /// <returns>The Visio shape at the location or null if no shape is
        /// found.</returns>
        [CLSCompliant(false)]
        public static IVShape GetClickedShape(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            double clickLocationX,
            double clickLocationY) {

            const double Tolerance = .0001;

            return GetClickedShape(
                drawingControl, clickLocationX, clickLocationY, Tolerance);
        }

        /// <summary>The GetClickedShape method finds a shape at the specified
        /// location.</summary>
        /// <remarks>If there are more than one shape at the location the top
        /// shape in Z order is returned.</remarks>
        /// <param name="drawingControl">Drawing control with the Visio page
        ///  containing the location.</param>
        /// <param name="clickLocationX">The X coordinate of the location in
        /// Visio page units (inches).</param>
        /// <param name="clickLocationY">The Y coordinate of the location in
        /// Visio page units (inches).</param>
        /// <param name="tolerance">The distance from the location to an
        /// included shape.</param>
        /// <returns>The Visio shape at the location or null if no shape is
        /// found.</returns>
        [CLSCompliant(false)]
        public static IVShape GetClickedShape(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            double clickLocationX,
            double clickLocationY,
            double tolerance) {

            if (drawingControl == null)
                return null;

            // The drawing control object must be valid.
            if (drawingControl.Window == null) {
            // do nothing...and nothing can be done as well.
                return null;
            }

            IVShape foundShape = null;

            Page currentPage;

            currentPage = drawingControl.Window.PageAsObj;

            // Use the spatial search method to return a list of
            // shapes at the location.
            Selection foundShapes = currentPage.get_SpatialSearch(
                clickLocationX, clickLocationY,
                (short)VisSpatialRelationCodes.visSpatialContainedIn,
                tolerance, 
                (short)VisSpatialRelationFlags.visSpatialFrontToBack);

            if (foundShapes.Count > 0) {

                // The selection collection of shapes is 1-based, so
                // the first shape is at index 1.
                foundShape = foundShapes[1];
            }

            return foundShape;
        }

        /// <summary>The GetShapeFromArguments method returns a reference to
        /// a shape given the command line arguments.</summary>
        /// <param name="visioApplication">The Visio application.</param>
        /// <param name="arguments">The command line arguments string containing:
        ///  /doc=id /page=id /shape=sheet.id.</param>
        /// <returns>The Visio shape or null.</returns>
        [CLSCompliant(false)]
        public static Shape GetShapeFromArguments(
            Microsoft.Office.Interop.Visio.IVApplication visioApplication,
            string arguments) {

            if (visioApplication == null) 
                return null;

            if (arguments == null)
                return null;

            const char equal = '=';
            const char argumentDelimiter = '/';

            // Standard Visio add-on command line arguments.
            const string commandLineArgumentDoc = "doc";
            const string commandLineArgumentPage = "page";
            const string commandLineArgumentShape = "shape";

            int index;
            int docId = -1;
            int pageId = -1;
            string shapeId = "";
            string[] contextParts;
            string contextPart;
            string[] argumentParts;
            Document document = null;
            Page page = null;
            Shape targetShape = null;

            // Parse the command line arguments.
            contextParts = arguments.Trim().Split(argumentDelimiter);

            for (index = contextParts.GetLowerBound(0); index <= 
                contextParts.GetUpperBound(0); index++) {

                contextPart = contextParts[index].Trim();

                if (contextPart.Length > 0) {

                    // Separate the parameter from the parameter value.
                    argumentParts = contextPart.Split(equal);

                    if (argumentParts.GetUpperBound(0) == 1) {

                        // Get the doc, page, and shape argument values.
                        if (commandLineArgumentDoc.Equals(argumentParts[0])) {

                            docId = Convert.ToInt16(argumentParts[1],
                                System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (commandLineArgumentPage.Equals(argumentParts[0])) {

                            pageId = Convert.ToInt16(argumentParts[1],
                                System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else  if (commandLineArgumentShape.Equals(argumentParts[0])) {

                            shapeId = argumentParts[1];
                        }
                    }
                }
            }

            // If the command line arguments contains document, page, and shape
            // then look up the shape.
            if ((docId > 0) && (pageId > 0) && (shapeId.Length > 0)) {

                document = visioApplication.Documents[docId];
                page = document.Pages[pageId];
                targetShape = page.Shapes[shapeId];
            }

            return targetShape;
        }

        /// <summary>The OpenStencil method opens the specified Visio document
        ///  stencil.</summary>
        /// <param name="drawingControl">Drawing control with the collection of
        ///  Visio documents to add the new stencil to.</param>
        /// <param name="stencilPath">The stencil path\filename to open.</param>
        /// <returns>Document object if the stencil is opened successfully.
        /// A COMException is thrown if the stencil cannot be opened.</returns>
        [CLSCompliant(false)]
        public static Document OpenStencil(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            string stencilPath) {

            if (drawingControl == null) {
                throw new ArgumentNullException("drawingControl");
            }

            // The drawing control object must be valid.
            if (drawingControl.Window == null || drawingControl.Window.Application == null) {

                // Throw a meaningful error.
                throw new ArgumentNullException("drawingControl");
            }
            Document targetStencil = null;

            Documents targetDocuments;
            targetDocuments = (Documents)drawingControl.Window.Application.Documents;

            // Open the stencil file invisibly.
            // A COMException will be thrown if the open fails.
            // If the stencil is currently open Visio will return a 
            // reference to the open stencil.
            targetStencil = targetDocuments.OpenEx(stencilPath,    
                (short)VisOpenSaveArgs.visOpenRO |
                (short)VisOpenSaveArgs.visOpenHidden |
                (short)VisOpenSaveArgs.visOpenMinimized |
                (short)VisOpenSaveArgs.visOpenNoWorkspace);

            return(targetStencil);
        }

        /// <summary>The GetMaster method gets the master by name.</summary>
        /// <param name="drawingControl">Drawing control with the collection of
        ///  Visio documents that contain the stencil and masters.</param>
        /// <param name="stencilPath">The stencil path\filename.</param>
        /// <param name="masterNameU">The universal name of the master.</param>
        /// <returns>Master object if found. A COMException is thrown if not found.</returns>
        [CLSCompliant(false)]
        public static Master GetMaster(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            string stencilPath,
            string masterNameU) {

            // The drawing control object must be valid.
            if (drawingControl == null) {

                // Throw a meaningful error.
                throw new ArgumentNullException("drawingControl");
            }

            Master targetMaster = null;

            Document targetDocument;
            Masters targetMasters;

            targetDocument = OpenStencil(drawingControl, stencilPath);
            targetMasters = targetDocument.Masters;
            targetMaster = targetMasters.get_ItemU(masterNameU);

            return(targetMaster);
        }

        /// <summary>The SaveDrawing method prompts the user to save changes
        /// to the Visio document.</summary>
        /// <param name="drawingControl">Drawing control with the Visio
        ///  document to save.</param>
        /// <param name="promptFirst">Display "save changes" prompt.</param>
        /// <returns>The id of the message box button that dismissed
        /// the dialog.</returns>
        [CLSCompliant(false)]
        public static DialogResult SaveDrawing(
            AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl drawingControl,
            bool promptFirst) {

                if (drawingControl == null) {
                    return DialogResult.Cancel;
                }


            SaveFileDialog saveFileDialog = null;
            DialogResult result = DialogResult.No;
            string targetFilename = string.Empty;
            Document targetDocument;
            
            try {
                targetFilename = drawingControl.Src;
                targetDocument = (Document)drawingControl.Document;

                // Prompt to save changes.
                if (promptFirst == true) {

                    string prompt = string.Empty;
                    string title = string.Empty;

                    title = Utility.GetResourceString(Strings.SaveDialogTitle);

                    if (targetFilename == null) {
                        return DialogResult.Cancel;
                    }

                    // Save changes to the existing drawing.
                    if (targetFilename.Length > 0) {
                        prompt = Utility.GetResourceString(Strings.SavePrompt);
                        prompt += Environment.NewLine;
                        prompt += targetFilename;
                    }
                    else {

                        // Save changes as new drawing.
                        prompt = Utility.GetResourceString(Strings.SaveAsPrompt);
                    }
                    result = Utility.RtlAwareMessageBoxShow(null, prompt, title,
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
                        drawingControl.Window.Application.AlertResponse);
                }
                else {
                    result = DialogResult.Yes;
                }

                // Display a file browse dialog to select path and filename.
                if ((DialogResult.Yes == result) && (targetFilename.Length == 0)) {

                    // Set up the save file dialog and let the user specify the
                    // name to save the document to.
                    saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Title =
                        Utility.GetResourceString(Strings.SaveDialogTitle);
                    saveFileDialog.Filter =
                        Utility.GetResourceString(Strings.SaveDialogFilter);
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.InitialDirectory =
                        System.Windows.Forms.Application.StartupPath;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK) {

                        targetFilename = saveFileDialog.FileName;
                    }
                }
                // Save the document to the filename specified by
                // the end user in the save file dialog, or the existing file name.
                if ((DialogResult.Yes == result) && (targetFilename.Length > 0)) {

                    if (targetDocument != null) {
                        targetDocument.SaveAs(targetFilename);
                    }
                    drawingControl.Src = targetFilename;
                    drawingControl.Document.Saved = true;
                }
                
            }
            finally {

                // Make sure the dialog is cleaned up.
                if (saveFileDialog != null) {
                    saveFileDialog.Dispose();
                }
            }
            return result;
        }

        /// <summary>This method displays the exception based on the
        /// boolean value.</summary>
        /// <param name="messageResourceName">Resource name of message to dispay.</param>
        /// <param name="exception">Exception object.</param>
        /// <param name="showErrorDialog">Display the dialog when true.</param>
        public static void DisplayException(string messageResourceName,
            Exception exception, bool showErrorDialog) {

            int alertResponse;
            alertResponse = showErrorDialog ? 0 : -1;

            DisplayException(messageResourceName, exception, alertResponse);
        }

        /// <summary>This method displays the exception based on the
        /// AlertResponse value.</summary>
        /// <param name="messageResourceName">Resource name of message to dispay.</param>
        /// <param name="exception">Exception object.</param>
        /// <param name="alertResponse">AlertResponse value of the running
        /// Visio instance.</param>
        public static void DisplayException(string messageResourceName,
            Exception exception, int alertResponse) {

            if (exception == null) {
                return;
            }

            string message;
            message = GetResourceString(messageResourceName);

#if DEBUG
            // Include call stack infomation when running a debug build.
            message += Environment.NewLine + Environment.NewLine;
            message += exception.ToString();
#endif
            Utility.RtlAwareMessageBoxShow(null, message, exception.Source,
                MessageBoxButtons.OK, MessageBoxIcon.Error, alertResponse);
            
            if (exception.TargetSite != null) {
                Debug.WriteLine(exception.Message, exception.TargetSite.Name);
            }
        }

        /// <summary>This method loads the string from the embedded resource.
        /// </summary>
        /// <param name="resourceName">Name of the resource to be loaded.</param>
        /// <returns>Loaded resource string if successful, otherwise empty
        /// string.</returns>
        public static string GetResourceString(string resourceName) {

            string resourceValue = "";

            resourceValue = theResourceManager.GetString(resourceName, 
                System.Globalization.CultureInfo.CurrentUICulture);

            return resourceValue;
        }
        
        /// <summary>class for right-to-left aware message box</summary>
        public static DialogResult RtlAwareMessageBoxShow(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, int alertResponse)
        {
            if (alertResponse != 0)
                return (DialogResult)alertResponse;

            Control control = owner as Control;
            bool bRightToLeft = false;
            MessageBoxOptions options = (MessageBoxOptions)0;
            if (control != null)
            {
                bRightToLeft = (control.RightToLeft == RightToLeft.Yes);
            }
            // Ask the CurrentUICulture if we are running under right-to-left.
            else
                bRightToLeft = System.Globalization.CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft;

            if (bRightToLeft)
            {
                options |= MessageBoxOptions.RtlReading |
                MessageBoxOptions.RightAlign;
            }

            return MessageBox.Show(owner, text, caption,
                                    buttons, icon, MessageBoxDefaultButton.Button1, options);
        }
    }
}
