// Strings.cs
// compile with: /doc:Strings.xml
// <copyright>Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>This file contains string constants used in the Office Plan sample.</summary>
using System.Diagnostics.CodeAnalysis;

namespace VisioEditor
{

    /// <summary>The Strings class provides all the string constants that are
    /// shared between the class modules in the project.</summary>
    public sealed class Strings
    {
        // <summary>The string resource names are used to read the strings
        // from the resource file.</summary>

        /// <summary>Disable constructor for class with only static members.</summary>
        private Strings() { }

        /// <summary>Resource file name.</summary>
        public const string ResourceFilename = "OfficePlanSample.OfficePlanSample";
        /// <summary>Configuration file name.</summary>
        public const string ConfigurationFilename = "ConfigurationFilename";
        /// <summary>Product data file name.</summary>
        public const string ProductDataFilename = "ProductDataFilename";
        /// <summary>Visio stencil file name.</summary>
        public const string ProductStencilFilename = "ProductStencilFilename";

        /// <summary>Grid control description column header text.</summary>
        public const string GridDescriptionHeader = "GridDescriptionHeader";
        /// <summary>Grid control product id column header text.</summary>
        public const string GridProductIdHeader = "GridProductIdHeader";
        /// <summary>Grid control quantity column header text.</summary>
        public const string GridQuantityHeader = "GridQuantityHeader";
        /// <summary>Grid control range column header text.</summary>
        public const string GridRangeHeader = "GridRangeHeader";
        /// <summary>Grid control retail price column header text.</summary>
        public const string GridRetailPriceHeader = "GridRetailPriceHeader";
        /// <summary>Grid control wholesale price column header text.</summary>
        public const string GridWholesalePriceHeader = "GridWholesalePriceHeader";

        /// <summary>Visio application object null error message.</summary>
        public const string NullApplicationError = "NullApplicationError";
        /// <summary>Visio document object null error message.</summary>
        public const string NullDocumentError = "NullDocumentError";
        /// <summary>Missing stencil error message.</summary>
        public const string MissingStencilError = "MissingStencilError";
        /// <summary>Missing master shape error message.</summary>
        public const string MissingMasterError = "MissingMasterError";
        /// <summary>Database connect error message.</summary>
        public const string DatabaseConnectErrorMessage = "DatabaseConnectErrorMessage";
        /// <summary>File not found error message.</summary>
        public const string FileNotFoundErrorMessage = "FileNotFoundErrorMessage";
        /// <summary>Security error message.</summary>
        public const string SecurityErrorMessage = "SecurityErrorMessage";
        /// <summary>COM error message.</summary>
        public const string ComErrorMessage = "ComErrorMessage";
        /// <summary>Unknown error message.</summary>
        public const string UnknownErrorMessage = "UnknownErrorMessage";
        /// <summary>Application error message.</summary>
        public const string ApplicationErrorMessage = "ApplicationErrorMessage";
        /// <summary>Product data not available error message.</summary>
        public const string ProductDataNotAvailableErrorMessage = "ProductDataNotAvailableErrorMessage";
        /// <summary>Startup error message.</summary>
        public const string VisioStartupError = "VisioStartupError";

        /// <summary>File open dialog caption text.</summary>
        public const string OpenDialogTitle = "OpenDialogTitle";
        /// <summary>File open dialog file types filter.</summary>
        public const string OpenDialogFilter = "OpenDialogFilter";
        /// <summary>File save dialog caption text.</summary>
        public const string SaveDialogTitle = "SaveDialogTitle";
        /// <summary>File save dialog file types filter.</summary>
        public const string SaveDialogFilter = "SaveDialogFilter";
        /// <summary>Save changes prompt.</summary>
        public const string SavePrompt = "SavePrompt";
        /// <summary>Save changes as prompt.</summary>
        public const string SaveAsPrompt = "SaveAsPrompt";

        // <summary>The file name information includes the name of the stencil
        // file with the furniture shapes.  Also included are the suffix
        // values for Visio stencil and document files.</summary>

        /// <summary>Visio stencil suffix VSS.</summary>
        public const string VssSuffix = ".VSS";
        /// <summary>Visio 2013 stencil suffix VSSX.</summary>
        public const string VssxSuffix = ".VSSX";
        /// <summary>Visio 2013 macro enabled stencil suffix VSSM.</summary>
        public const string VssmSuffix = ".VSSM";
        /// <summary>Visio xml stencil suffix VSX.</summary>
        public const string VsxSuffix = ".VSX";
        /// <summary>Visio drawing suffix VSD.</summary>
        public const string VsdSuffix = ".VSD";
        /// <summary>Visio 2013 drawing suffix VSDX.</summary>
        public const string VsdxSuffix = ".VSDX";
        /// <summary>Visio 2013 macro enabled drawing suffix VSDM.</summary>
        public const string VsdmSuffix = ".VSDM";
        /// <summary>Visio xml drawing suffix VDX.</summary>
        public const string VdxSuffix = ".VDX";

        /// <summary> The only cell that needs to be read from the shape is
        /// the product ID.</summary>
        public const string ProductIdCellName = "Prop.ProductId";
    }
}