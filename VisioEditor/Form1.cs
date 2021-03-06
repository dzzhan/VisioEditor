﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.VisOcx;
using System.Xml;

namespace VisioEditor
{
    public partial class GraphEditor : Form
    {
        Page    m_stCurrentPage;
        Document m_stBasicMaster;
        Document m_stAuditMaster;
        Document m_stClassMaster;

        public GraphEditor()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
        }

        private static void ConnectShapes(Shape shape1, Shape shape2, Shape connector)
        {
            // get the cell from the source side of the connector
            Cell beginXCell = connector.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                                                     (short)VisRowIndices.visRowXForm1D,
                                                     (short)VisCellIndices.vis1DBeginX);

            // glue the source side of the connector to the first shape
            beginXCell.GlueTo(shape1.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                                                  (short)VisRowIndices.visRowXFormOut,
                                                  (short)VisCellIndices.visXFormPinX));

            // get the cell from the destination side of the connector
            Cell endXCell = connector.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                                                    (short)VisRowIndices.visRowXForm1D,
                                                    (short)VisCellIndices.vis1DEndX);

            // glue the destination side of the connector to the second shape
            endXCell.GlueTo(shape2.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                                                (short)VisRowIndices.visRowXFormOut,
                                                (short)VisCellIndices.visXFormPinX));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            axDrawingControl1.Window.ShowScrollBars = (short)VisScrollbarStates.visScrollBarBoth;
            axDrawingControl1.Window.ShowRulers = 0;
            axDrawingControl1.Window.ZoomBehavior = VisZoomBehavior.visZoomVisioExact;
            m_stBasicMaster = axDrawingControl1.Document.Application.Documents.OpenEx("Basic_M.vss",
                                                                        (short)VisOpenSaveArgs.visOpenDocked);
            m_stAuditMaster = axDrawingControl1.Document.Application.Documents.OpenEx("AUDIT_M.VSSX",
                                                                         (short)VisOpenSaveArgs.visOpenDocked);
            m_stClassMaster = axDrawingControl1.Document.Application.Documents.OpenEx("USTRME_M.VSSX",
                                                                           (short)VisOpenSaveArgs.visAddDocked);
            m_stCurrentPage = axDrawingControl1.Document.Pages[1];
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "";
            dlg.Filter = "Visio文件 (*.vsdx)|*.vsdx|所有文件(*.*)|*.*";
            dlg.FilterIndex = 1;
            dlg.RestoreDirectory = true;
            dlg.CheckFileExists = true;
            dlg.CheckPathExists = true;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.Trim() != string.Empty)
                {
                    m_stBasicMaster.Close();
                    m_stAuditMaster.Close();
                    m_stClassMaster.Close();
                    axDrawingControl1.Src = dlg.FileName;
                    m_stBasicMaster = axDrawingControl1.Document.Application.Documents.OpenEx("Basic_M.vss",
                                                                                (short)VisOpenSaveArgs.visOpenDocked);
                    m_stAuditMaster = axDrawingControl1.Document.Application.Documents.OpenEx("AUDIT_M.VSSX",
                                                                                 (short)VisOpenSaveArgs.visOpenDocked);
                    m_stClassMaster = axDrawingControl1.Document.Application.Documents.OpenEx("USTRME_M.VSSX",
                                                                                   (short)VisOpenSaveArgs.visAddDocked);
                    m_stCurrentPage = axDrawingControl1.Document.Pages[1];
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "";
            dlg.Filter = "Visio文件 (*.vsdx)|*.vsdx|所有文件(*.*)|*.*";
            dlg.FilterIndex = 1;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.Trim() != string.Empty)
                {
                    Document visDoc = axDrawingControl1.Document;
                    visDoc.SaveAs(dlg.FileName);
                }
            }
        }

        private void addTaskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Shape shape1 = m_stCurrentPage.Drop(m_stBasicMaster.Masters["矩形"], 1, 11);
            shape1.Text = "DSP\n191";
            shape1.get_CellsU("LineColor").ResultIU = (double)VisDefaultColors.visDarkGreen;//有电（绿色）
            shape1.Cells["FillForegnd"].Formula = "RGB(192,255,206)";
        }

        private void CheckShapes()
        {
            for (int i = 1; i <= m_stCurrentPage.Shapes.Count; i++)
            {
                Shape sp = m_stCurrentPage.Shapes[i];
                if (sp.Master.Name == "动态连接线")
                {
                    FindConnection(ref sp);
                }
                string text = sp.Text;
                Cell ce = sp.get_CellsSRC((short)VisSectionIndices.visSectionProp, 
                                          (short)VisRowIndices.visRowProp,
                                          (short)VisCellIndices.visUserValue);
                int row = ce.Row;
                int col = ce.Column;
                string sid = ce.Formula;
            }
        }

        private void FindConnection(ref Shape sp)
        {
            List<string> vShapes = new List<string>();
            foreach (Connect conn in sp.Connects)
            {
                Cell fromCell = conn.FromCell;
                Cell toCell = conn.ToCell;
                if (fromCell.Shape.Master.Name != "动态连接线")
                {
                    vShapes.Add(fromCell.Shape.Text);
                }
                else if (toCell.Shape.Master.Name == "动态连接线")
                {
                    vShapes.Add(toCell.Shape.Text);
                }
            }
            string strConnect = vShapes[0] + " -> " + vShapes[1];
        }

        private void checkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckShapes();
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "";
            dlg.Filter = "XML文件 (*.xml)|*.xml|所有文件(*.*)|*.*";
            dlg.FilterIndex = 1;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.Trim() != string.Empty)
                {
                    ExportXML(dlg.FileName);
                }
            }
        }

        private void onDrawingControlMouseUp(object sender, AxMicrosoft.Office.Interop.VisOcx.EVisOcx_MouseUpEvent eventData)
        {
            ShowMessage("MouseUp!");
        }

        private void ShowMessage(string strMsg)
        {
            dbgMsgText.Text += (strMsg + "\r\n");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void ExportXML(string strXmlFileName)
        {
            //创建XmlDocument对象
            XmlDocument xmlDoc = new XmlDocument();

            //XML的声明<?xml version="1.0" encoding="gb2312"?> 
            XmlDeclaration xmlSM = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            
            //追加xmldecl位置
            xmlDoc.AppendChild(xmlSM);
            
            //添加一个名为Gen的根节点
            XmlElement xml = xmlDoc.CreateElement("", "task_graph", "");
            
            //追加Gen的根节点位置
            xmlDoc.AppendChild(xml);
            
            //添加另一个节点,与Gen所匹配，查找<Gen>
            XmlNode root = xmlDoc.SelectSingleNode("task_graph");
            XmlNode taskList = xmlDoc.CreateElement("task_list");
            XmlNode connList = xmlDoc.CreateElement("connections");
            XmlNode submitList = xmlDoc.CreateElement("submits");
            root.AppendChild(taskList);
            root.AppendChild(connList);
            root.AppendChild(submitList);
            XmlElement taskItem = null;
            for (int i = 1; i <= m_stCurrentPage.Shapes.Count; i++)
            {
                Shape sp = m_stCurrentPage.Shapes[i];
                if ((sp.Master == null) && (sp.Type == (short)VisShapeTypes.visTypeGroup))
                {
                    taskItem = xmlDoc.CreateElement("task");
                    SetGroupSelectMode(sp, 1);
                    foreach (Shape subShape in sp.Shapes)
                    {
                        if (subShape.Master.NameU == "Class")
                        {
                            taskItem.SetAttribute("name", subShape.Text);
                        }
                        else
                        {
                            XmlElement subTask = xmlDoc.CreateElement("properties");
                            subTask.SetAttribute("name", subShape.Text);
                            subTask.SetAttribute("shape", subShape.Master.Name);
                            subTask.SetAttribute("pinx", subShape.Cells["PinX"].Formula);
                            subTask.SetAttribute("piny", subShape.Cells["PinY"].Formula);
                            taskItem.AppendChild(subTask);
                        }
                    }
                    SetGroupSelectMode(sp, 2);
                    taskItem.SetAttribute("pinx", sp.Cells["PinX"].Formula);
                    taskItem.SetAttribute("piny", sp.Cells["PinY"].Formula);
                    taskItem.SetAttribute("width", sp.Cells["Width"].Formula);
                    taskList.AppendChild(taskItem);
                }
                else if (sp.Master.NameU == "Inheritance")
                {
                    string strNameU = sp.Master.NameU;
                    List<string> vShapes = new List<string>();
                    foreach (Connect cn in sp.Connects)
                    {
                        Cell fromCell = cn.FromCell;
                        Cell toCell = cn.ToCell;
                        if (fromCell.Shape.Master.NameU != "Inheritance")
                        {
                            vShapes.Add(fromCell.Shape.Text);
                        }
                        else if (toCell.Shape.Master.NameU != "Inheritance")
                        {
                            vShapes.Add(toCell.Shape.Text);
                        }
                    }

                    if (vShapes.Count == 2)
                    {
                        if (sp.Text.ToLower() != "submit")
                        {
                            XmlElement conn = xmlDoc.CreateElement("conn");
                            conn.SetAttribute("from", vShapes[0]);
                            conn.SetAttribute("to", vShapes[1]);
                            connList.AppendChild(conn);
                        }
                        else
                        {
                            XmlElement submit = xmlDoc.CreateElement("submit");
                            submit.SetAttribute("source", vShapes[0]);
                            submit.SetAttribute("target", vShapes[1]);
                            submitList.AppendChild(submit);
                        }
                    }
                }
                /*
                if (sp.Master.Name == "类")
                {
                    taskItem = xmlDoc.CreateElement("task");
                    taskItem.SetAttribute("name", sp.Text);
                    taskItem.SetAttribute("shape", sp.Master.Name);
                    //taskItem.SetAttribute("pinx", strPinX);
                    //taskItem.SetAttribute("piny", strPinY);
                    //task.SetAttribute("fill_color", sp.Cells["FillForegnd"].Formula);
                    taskList.AppendChild(taskItem);
                    SetGroupSelectMode(sp, 0);
                    foreach(Shape subShape in sp.Shapes)
                    {
                        XmlElement subTask = xmlDoc.CreateElement("properties");
                        subTask.SetAttribute("name", subShape.Text);
                        subTask.SetAttribute("shape", subShape.Master.Name);
                        taskItem.AppendChild(subTask);
                    }
                    SetGroupSelectMode(sp, 2);
                }
                else if ((taskItem != null) && ((sp.Master.Name == "成员") || (sp.Master.Name == "分隔符")))
                {
                    XmlElement taskProp = xmlDoc.CreateElement("properties");
                    taskProp.SetAttribute("name", sp.Text);
                    taskProp.SetAttribute("shape", sp.Master.Name);
                    //taskProp.SetAttribute("pinx", strPinX);
                    //taskProp.SetAttribute("piny", strPinY);
                    taskItem.AppendChild(taskProp);
                }
                else if (sp.Master.Name != "动态连接线")
                {
                    XmlElement task = xmlDoc.CreateElement("task");
                    task.SetAttribute("name", sp.Text);
                    task.SetAttribute("shape", sp.Master.Name);
                    //task.SetAttribute("fill_color", sp.Cells["FillForegnd"].Formula);
                    taskList.AppendChild(task);
                }
                else
                {
                    List<string> vShapes = new List<string>();
                    foreach (Connect cn in sp.Connects)
                    {
                        Cell fromCell = cn.FromCell;
                        Cell toCell = cn.ToCell;
                        if (fromCell.Shape.Master.Name != "动态连接线")
                        {
                            vShapes.Add(fromCell.Shape.Text);
                        }
                        else if (toCell.Shape.Master.Name != "动态连接线")
                        {
                            vShapes.Add(toCell.Shape.Text);
                        }
                    }

                    if (vShapes.Count == 2)
                    {
                        if (sp.Text.ToLower() != "submit")
                        {
                            XmlElement conn = xmlDoc.CreateElement("conn");
                            conn.SetAttribute("from", vShapes[0]);
                            conn.SetAttribute("to", vShapes[1]);
                            connList.AppendChild(conn);
                        }
                        else
                        {
                            XmlElement submit = xmlDoc.CreateElement("submit");
                            submit.SetAttribute("source", vShapes[0]);
                            submit.SetAttribute("target",vShapes[1]);
                            submitList.AppendChild(submit);
                        }
                    }
                }*/

            }
            xmlDoc.Save(strXmlFileName);
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "";
            dlg.Filter = "XML文件 (*.xml)|*.xml|所有文件(*.*)|*.*";
            dlg.FilterIndex = 1;
            dlg.RestoreDirectory = true;
            dlg.CheckFileExists = true;
            dlg.CheckPathExists = true;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                if (dlg.FileName.Trim() != string.Empty)
                {
                    m_stCurrentPage = axDrawingControl1.Document.Pages[1];
                    ImportXML(dlg.FileName);
                }
            }
        }

        private void RemoveUnusedShapes()
        {
            List<Shape> vUnusedShapes = new List<Shape>();
            foreach (Shape sp in m_stCurrentPage.Shapes)
            {
                if ((sp.Master.Name == "成员") || (sp.Master.Name == "分隔符"))
                {
                    vUnusedShapes.Add(sp);
                }
            }

            foreach (Shape sp in vUnusedShapes)
            {
                sp.Delete();
            }
        }

        private void ImportXML(string strXmlFile)
        {
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(strXmlFile);

            XmlNode xmlRoot = xmldoc.SelectSingleNode("task_graph");
            if (null == xmlRoot)
            {
                return;
            }

            Dictionary<string, Shape> vTaskShapes = new Dictionary<string, Shape>();
            Dictionary<string, Dictionary<string, string>> vTaskProps = new Dictionary<string, Dictionary<string, string>>();
            XmlNode taskList = xmlRoot.SelectSingleNode("task_list");
            if (taskList != null)
            {
                foreach (XmlElement elem in taskList.ChildNodes)
                {
                    if (elem.Name.ToLower() == "task")
                    {
                        string strTaskName = elem.Attributes["name"].Value;
                        string strPinX = elem.Attributes["pinx"].Value;
                        string strPinY = elem.Attributes["piny"].Value;
                        string strWidth = elem.Attributes["width"].Value;

                        Dictionary<string, string> vPropties = new Dictionary<string, string>();
                        foreach (XmlElement prop in elem.ChildNodes)
                        {
                            vPropties.Add(prop.Attributes["name"].Value, prop.Attributes["shape"].Value);
                        }

                        Shape sp = DrawTask(strTaskName, strPinX, strPinY, strWidth);
                        vTaskShapes.Add(strTaskName, sp);
                        vTaskProps.Add(strTaskName, vPropties);
                    }
                }

                RemoveUnusedShapes();

                foreach (string strTaskName in vTaskShapes.Keys)
                {
                    if (vTaskProps.ContainsKey(strTaskName))
                    {
                        DrawTaskProp(vTaskShapes[strTaskName], vTaskProps[strTaskName]);
                    }
                }
            }

            XmlNode connList = xmlRoot.SelectSingleNode("connections");
            if (connList != null)
            {
                foreach (XmlElement elem in connList.ChildNodes)
                {
                    if (elem.Name.ToLower() == "conn")
                    {
                        string strSource = elem.Attributes["from"].Value;
                        string strTarget = elem.Attributes["to"].Value;
                        ConnectTask(vTaskShapes[strSource], vTaskShapes[strTarget]);
                    }
                }
            }

            XmlNode submitList = xmlRoot.SelectSingleNode("submits");
            if (submitList != null)
            {
                foreach (XmlElement elem in submitList.ChildNodes)
                {
                    if (elem.Name.ToLower() == "submit")
                    {
                        string strSource = elem.Attributes["source"].Value;
                        string strTarget = elem.Attributes["target"].Value;
                        ConnectTask(vTaskShapes[strSource], vTaskShapes[strTarget], true);
                    }
                }
            }
        }

        private Shape DrawTask(string strTaskName, string strPinX, string strPinY, string strWidth)
        {
            Shape sp = m_stCurrentPage.Drop(m_stClassMaster.Masters["类"], 1, 11);
            sp.Cells["PinX"].Formula = strPinX;
            sp.Cells["PinY"].Formula = strPinY;
            sp.Cells["Width"].Formula = strWidth;
            sp.Text = strTaskName;
            return sp;
        }

        private void DrawTaskProp(Shape sp, Dictionary<string, string> vTaskProp)
        {
            foreach (string strPropName in vTaskProp.Keys)
            {
                Shape stPropShape = m_stCurrentPage.DropIntoList(m_stClassMaster.Masters[vTaskProp[strPropName]], sp, 1);
                stPropShape.Text = strPropName;
            }
        }

        private Shape DrawTask(string strTaskName, int xPos, int yPos, string strShapeName, string strColor = "RGB(192,255,206)")
        {
            Shape sp = m_stCurrentPage.Drop(m_stBasicMaster.Masters[strShapeName], xPos, yPos);
            sp.Text = strTaskName;
            //sp.get_CellsU("LineColor").ResultIU = (double)VisDefaultColors.visDarkGreen;
            sp.Cells["FillForegnd"].Formula = strColor;
            return sp;
        }

        private void ConnectTask(Shape sp1, Shape sp2, bool bIsSubmit = false)
        {
            if ((null == sp1) || (null == sp2))
            {
                return;
            }

            Shape conn = m_stCurrentPage.Drop(m_stClassMaster.Masters["继承"], 4.50, 4.50);
            if (bIsSubmit)
            {
                conn.Text = "submit";
            }
            ConnectShapes(sp1, sp2, conn);
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void SetGroupSelectMode(Shape targetShape, int selectMode)
        {
            targetShape.get_CellsSRC((short)VisSectionIndices.visSectionObject, 
                                     (short)VisRowIndices.visRowGroup,
                                     (short)VisCellIndices.visGroupSelectMode).FormulaU = selectMode.ToString();
        }

        private void newToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            m_stCurrentPage.Delete(1);
        }
    }
}
