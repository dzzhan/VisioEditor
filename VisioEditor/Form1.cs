using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using AxVisOcx = Microsoft.Office.Interop.VisOcx;
//using System.Web.UI.WebControls;
using Cell = Microsoft.Office.Interop.Visio.Cell;

namespace VisioEditor
{
    public partial class GraphEditor : Form
    {
        public GraphEditor()
        {
            InitializeComponent();
        }

        private static void ConnectShapes(Visio.Shape shape1, Visio.Shape shape2, Visio.Shape connector)
        {
            // get the cell from the source side of the connector

            Cell beginXCell = connector.get_CellsSRC(

            (short)Visio.VisSectionIndices.visSectionObject,


            (short)Visio.VisRowIndices.visRowXForm1D,

            (short)Visio.VisCellIndices.vis1DBeginX);

            // glue the source side of the connector to the first shape

            beginXCell.GlueTo(shape1.get_CellsSRC(

            (short)Visio.VisSectionIndices.visSectionObject,

            (short)Visio.VisRowIndices.visRowXFormOut,

            (short)Visio.VisCellIndices.visXFormPinX));

            // get the cell from the destination side of the connector

            Cell endXCell = connector.get_CellsSRC(

            (short)Visio.VisSectionIndices.visSectionObject,

            (short)Visio.VisRowIndices.visRowXForm1D,

            (short)Visio.VisCellIndices.vis1DEndX);

            // glue the destination side of the connector to the second shape

            endXCell.GlueTo(shape2.get_CellsSRC(

            (short)Visio.VisSectionIndices.visSectionObject,

            (short)Visio.VisRowIndices.visRowXFormOut,

            (short)Visio.VisCellIndices.visXFormPinX));

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            axDrawingControl1.Window.Zoom = 0.5;
            axDrawingControl1.Window.ShowScrollBars = (short)Visio.VisScrollbarStates.visScrollBarBoth;
            axDrawingControl1.Window.ShowRulers = 0;
            axDrawingControl1.Window.BackgroundColor = (uint)ColorTranslator.ToOle(Color.Red);
            axDrawingControl1.Window.BackgroundColorGradient = (uint)ColorTranslator.ToOle(Color.Red);
            axDrawingControl1.Window.ZoomBehavior = Visio.VisZoomBehavior.visZoomVisioExact;
            Visio.Page currentPage = axDrawingControl1.Document.Pages[1];
            Visio.Document currentStencil = axDrawingControl1.Document.Application.Documents.OpenEx("Basic_M.vss",
                                                                        (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Document auditStencil = axDrawingControl1.Document.Application.Documents.OpenEx("AUDIT_M.VSSX",
                                                                         (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Shape shape1 = currentPage.Drop(currentStencil.Masters["矩形"], 1.50, 1.50);
            Visio.Shape shape2 = currentPage.Drop(currentStencil.Masters["正方形"], 2.50, 3.50);
            Visio.Shape connector = currentPage.Drop(auditStencil.Masters["动态连接线"], 4.50, 4.50);
            ConnectShapes(shape1, shape2, connector);
            Cell arrowCell = connector.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineEndArrow);
            arrowCell.FormulaU = "5";
            connector.get_Cells("EndArrow").Formula = "=5";
        }
    }
}
