using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Text.RegularExpressions;


using ClosedXML.Excel;

using DigitalPlatform;
using DigitalPlatform.IO;
using DigitalPlatform.Xml;
using DigitalPlatform.Marc;
using DigitalPlatform.Script;

namespace ExcelTobdf
{
    public partial class MainForm : Form
    {
        // 构造函数
        public MainForm()
        {
            InitializeComponent();
        }

        #region 一些成员变量

        // 列配置 xml对象
        XmlDocument _dom = null;

        // excel对象
        XLWorkbook workbook = null;


        const int COUNT = 10;
        Hashtable dataTable = new Hashtable();

        // 册条码重复列表
        //Hashtable dupBarcodes = new Hashtable();

        #endregion

        #region 一些界面输入的参数

        // 馆藏地
        string _location
        {
            get
            {
                return this.textBox_location.Text;
            }
        }

        // 机构代码
        string _libraryCode
        {
            get
            {
                return this.textBox_libraryCode.Text;
            }
        }

        // 册类型
        string _bookType
        {
            get
            {
                return this.textBox_bookType.Text;
            }
        }

        // 册前缀
        string _itemBarcodePrefix
        {
            get
            {
                return this.textBox_itemBarcodePrefix.Text;
            }
        }

        #endregion

        #region 窗体装载和控制可用

        // 窗体装载
        private void MainForm_Load(object sender, EventArgs e)
        {
            // 装载列配置文件
            string cfg = "cfg.xml";
            XmlDocument cfgDom = new XmlDocument();
            try
            {
                cfgDom.Load(cfg);
            }
            catch (Exception ex)
            {
                MessageBox.Show("载入配置文件发生错误:" + ex.Message);
                this.Close();
                return;
            }

            AddListViewColumn(cfgDom);

            this._dom = cfgDom;
        }

        // 设置控件是否可用
        void EnableControls(bool bEnable)
        {
            this.button_open.Enabled = bEnable;
            this.listView.Enabled = bEnable;
            this.button_output.Enabled = bEnable;
            this.button_start.Enabled = bEnable;
        }
        #endregion

        #region listview的一些事件

        // 把配置的列加到界面的listview里
        void AddListViewColumn(XmlDocument dom)
        {
            if (dom == null)
                return;

            // this.MainForm.OperHistory.AppendHtml("MarcSyntax = " + this._marcSyntax + "<br>");

            string xpath = "record[@marcSyntax='unimarc']/subfield";

            // this.MainForm.OperHistory.AppendHtml("XPath = " + xpath + "<br>");

            XmlNodeList subfields = dom.DocumentElement.SelectNodes(xpath);
            // this.MainForm.OperHistory.AppendHtml("NodeCount = " + labels.Count + "<br>");

            foreach (XmlNode node in subfields)
            {
                string label = DomUtil.GetAttr(node, "label");
                if (String.IsNullOrEmpty(label))
                    continue;

                if (node == null)
                    continue;

                string field = DomUtil.GetAttr(node, "field");
                if (field == "999")
                    continue;

                string name = field
                    + DomUtil.GetAttr(node, "ind")
                    + DomUtil.GetAttr(node, "code");

                this.listView.Columns.Add(new ColumnHeader()
                {
                    Name = name,
                    Text = label + "[" + name + "]",
                    Width = 150
                });
            }
            this.listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        // listview右键命令
        private void listView_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
                return;

            ContextMenu contextMenu = new ContextMenu();
            MenuItem menuItem = null;

            int nCount = this.listView.SelectedItems.Count;
            menuItem = new MenuItem(String.Format("移除选中的'{0}'行", nCount));
            menuItem.Click += menuItem_remove_Click;
            if (nCount == 0)
                menuItem.Enabled = false;
            contextMenu.MenuItems.Add(menuItem);
            contextMenu.Show(this.listView, e.Location);
        }

        // listview右键移除命令，可以给每个列头设置对应的字段
        private void menuItem_remove_Click(object sender, EventArgs e)
        {
            EnableControls(false);

            this.listView.BeginUpdate();
            for (int i = 0; i < this.listView.SelectedItems.Count; i++)
            {
                ListViewItem item = this.listView.SelectedItems[i];
                this.listView.Items.Remove(item);
                i--;
            }
            this.listView.EndUpdate();

            EnableControls(true);
        }

        // 确保列可用
        static void EnsureColumns(ListView listview,
            int nCount,
            int nInitialWidth = 200)
        {
            if (listview.Columns.Count >= nCount)
                return;

            for (int i = listview.Columns.Count; i < nCount; i++)
            {
                ColumnHeader col = new ColumnHeader();
                col.Text = i.ToString();
                col.Width = nInitialWidth;
                listview.Columns.Add(col);
            }
        }

        // 设置list列头的信息
        void listView_ColumnContextMenuClicked(object sender, ColumnHeader columnHeader)
        {
            if (this._dom == null)
                return;

            ContextMenuStrip contextMenu = new ContextMenuStrip();

            ToolStripMenuItem menuItem = null;
            Point mousePosition = System.Windows.Forms.Control.MousePosition;

            string xpath = "record[@marcSyntax='unimarc']/subfield";
            XmlNodeList subfields = this._dom.DocumentElement.SelectNodes(xpath);
            foreach (XmlNode node in subfields)
            {
                string label = DomUtil.GetAttr(node, "label");
                if (String.IsNullOrEmpty(label))
                    continue;

                string name = DomUtil.GetAttr(node, "field")
                    + DomUtil.GetAttr(node, "ind")
                    + DomUtil.GetAttr(node, "code");
                ColumnHeader column = this.listView.Columns[name];
                if (column != null)
                {
                    if (column == columnHeader)
                        continue;
                }

                menuItem = new ToolStripMenuItem();
                menuItem.Text = label + "[" + name + "]";
                menuItem.Click += (s, e) =>
                {
                    columnHeader.Name = name;
                    columnHeader.Text = label + "[" + name + "]";

                    this.listView.AutoResizeColumn(columnHeader.Index, ColumnHeaderAutoResizeStyle.HeaderSize);
                };
                contextMenu.Items.Add(menuItem);
            }

            contextMenu.Items.Add("-");

            menuItem = new ToolStripMenuItem();
            menuItem.Text = "不导入";
            menuItem.Click += (s, e) =>
            {
                columnHeader.Text = "[未设置]";
                columnHeader.Name = "";
                this.listView.AutoResizeColumn(columnHeader.Index, ColumnHeaderAutoResizeStyle.HeaderSize);
            };
            contextMenu.Items.Add(menuItem);

            if (this.listView.Columns[0] != columnHeader)
            {
                Point p = this.listView.PointToClient(System.Windows.Forms.Control.MousePosition);
                contextMenu.Show(this.listView, p);
            }
        }

        #endregion

        #region 选择来源文件和输出文件

        // 选择输入的excel文件
        private void button_open_Click(object sender, EventArgs e)
        {
            this.listView.Items.Clear();

            OpenFileDialog dlg = new OpenFileDialog()
            {
                Filter = "Excel文件(*.xlsx)|*.xlsx",
                RestoreDirectory = true,
                Title = "选择Excel文件...",
            };
            if (dlg.ShowDialog(this) == DialogResult.Cancel)
                return;

            this.textBox_filename.Text = dlg.FileName;

            try
            {
                workbook = new XLWorkbook(dlg.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            var sheet = workbook.Worksheets.FirstOrDefault();

            var rows = sheet.Rows();
            int nCount = rows.Count();

            EnableControls(false);

            foreach (var row in rows)
            {
                Application.DoEvents();
                this.Update();

                int rowNumber = row.RowNumber();

                ListViewItem listViewItem = new ListViewItem();
                listViewItem.Text = rowNumber.ToString();

                foreach (var cell in row.Cells())
                {
                    Application.DoEvents();
                    this.Update();

                    string strStyle = cell.Style.DateFormat.Format.Replace(";@", "");
                    string strValue = "";
                    if (cell.DataType == XLCellValues.DateTime)
                    {
                        cell.Style.DateFormat.Format = "yyyyMMdd";
                        strValue = cell.GetFormattedString();
                    }
                    else
                    {
                        strValue = cell.GetString();
                    }
                    int index = cell.Address.ColumnNumber;
                    if (index > listViewItem.SubItems.Count)
                    {
                        for (int i = listViewItem.SubItems.Count; i < index; i++)
                        {
                            listViewItem.SubItems.Add("");
                        }
                    }
                    listViewItem.SubItems.Insert(index, new ListViewItem.ListViewSubItem(listViewItem, strValue));
                }

                EnsureColumns(this.listView, listViewItem.SubItems.Count);
                this.listView.Items.Add(listViewItem);

                if (rowNumber == COUNT)
                    break;
            }

            EnableControls(true);
        }

        // 选择输出的bdf文件
        private void button_output_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog()
            {
                Filter = "bdf文件(*.bdf)|*.bdf|全部文件(*.*)|*.*",
                RestoreDirectory = true,
                Title = "指定输出文件名...",
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            this.textBox_output.Text = dlg.FileName;
        }

        #endregion

        #region 数据处理


        // 开始处理
        private void button_start_Click(object sender, EventArgs e)
        {
            // 定义第一个sheet
            var sheet = workbook.Worksheets.FirstOrDefault();

            // 行数
            var rows = sheet.Rows();
            int nCount = rows.Count();

            // 确保转换过程中控件不可用，防止用户乱点
            EnableControls(false);

            // 是否有marc
            int isHasMARC = this.listView.Columns.IndexOfKey("999##a");

            // 各个字段列号
            int nISBNIndex = this.listView.Columns.IndexOfKey("010##a");
            int nTitleIndex = this.listView.Columns.IndexOfKey("2001#a");
            int nBarcodeIndex = this.listView.Columns.IndexOfKey("905##b");
            int nLocationIndex = this.listView.Columns.IndexOfKey("905##c");
            int nBookTypeIndex = this.listView.Columns.IndexOfKey("905##a");
            int nAccessNoIndex = this.listView.Columns.IndexOfKey("905##d");
            int nNumSerialIndex = this.listView.Columns.IndexOfKey("905##e");
            int nClassIndex = this.listView.Columns.IndexOfKey("690##a");


            StringBuilder sb = new StringBuilder();
            foreach (var row in rows)
            {
                // 行号
                int rowNumber = row.RowNumber();

                // 忽略第一行
                if (rowNumber == 1)
                    continue;

                // 出让控制权
                Application.DoEvents();
                this.Update();
                // 在状态行显示进度
                this.toolStripStatusLabel_msg.Text = "正在转换...(" + rowNumber.ToString() + " / " + nCount.ToString() + ")";

                // 如果isbn 或者 题名为空，则不导入
                if (nISBNIndex == -1 || nTitleIndex == -1)
                    continue;

                // 取出isbn和题名，作为一个书目主键
                string strISBN = GetExcelCellValue(row.Cell(nISBNIndex));
                string strTitle = GetExcelCellValue(row.Cell(nTitleIndex));
                string strKey = strISBN.Replace("-", "") + "|" + strTitle;

#if ChongHua 
                // 如果是chonghua，只用题名作为主键

                if (nTitleIndex == -1)
                    continue;

                string strKey = GetExcelCellValue(row.Cell(nTitleIndex));
#endif
                // 册条码
                string strBarcode = nBarcodeIndex == -1 ? "" : GetExcelCellValue(row.Cell(nBarcodeIndex)).Trim();
                // 如果册条码为空，则不处理这行
                if (String.IsNullOrEmpty(strBarcode))
                    continue;

#if XIMA
                // ximajingrun册条码号二次处理
                // 999999
                // 888889
                // 888888
                // 0000009
                if (strBarcode.Length != 6)
                    strBarcode = "ERR" + strBarcode;
                else if (strBarcode == "999999" || strBarcode == "888889" || strBarcode == "888888")
                    strBarcode = "ERR" + strBarcode;
#endif
                // 如果配置了条码前缀，则在册条码前面加上前缀
                if (String.IsNullOrEmpty(_itemBarcodePrefix) == false)
                    strBarcode = _itemBarcodePrefix + strBarcode;

                // 2021/6/4 renyh 去掉这一段，册条码重复的记录也要导入，要不册数量不一致。最近chonghua就翻出这一个老问题了。
                //// 检查是否是重复的册条码
                //if (dupBarcodes.ContainsKey(strBarcode))
                //{
                //    sb.AppendLine(strBarcode + ":");
                //    sb.AppendLine(strKey);
                //    string dup = dupBarcodes[strBarcode] as string;
                //    sb.AppendLine(dup);
                //    sb.AppendLine("=====================");
                //    continue;
                //}
                //else
                //{
                //    dupBarcodes.Add(strBarcode, strKey);
                //}


                // 馆藏地
                string strLocation = nLocationIndex == -1 ? "" : GetExcelCellValue(row.Cell(nLocationIndex)).Trim();
                if (String.IsNullOrEmpty(strLocation) == true)
                {
                    if (String.IsNullOrEmpty(this._location) == false)
                        strLocation = this._location;
                }
                // 如果配置了分馆，馆藏地要前面加上分馆这一截
                if (String.IsNullOrEmpty(this._libraryCode) == false)
                    strLocation = this._libraryCode + "/" + strLocation;

                // 图书类型，如果没置图书类型，默认设为普通
                string strBookType = nBookTypeIndex == -1 ? "" : GetExcelCellValue(row.Cell(nBookTypeIndex)).Trim();
                if (String.IsNullOrEmpty(strBookType) == true)
                {
                    if (String.IsNullOrEmpty(this._bookType) == false)
                        strBookType = this._bookType;
                    else
                        strBookType = "普通";
                }

                // 索取号
                string strClassNum = nClassIndex == -1 ? "" : GetExcelCellValue(row.Cell(nClassIndex)).Trim();
                string strAccessNo = nAccessNoIndex == -1 ? "" : GetExcelCellValue(row.Cell(nAccessNoIndex)).Trim();
                if (string.IsNullOrEmpty(strAccessNo) || strAccessNo[0] == '/' || strAccessNo[strAccessNo.Length - 1] == '/')
                {
                    string strNumSerial = nNumSerialIndex == -1 ? "" : GetExcelCellValue(row.Cell(nNumSerialIndex)).Trim();
                    if (!string.IsNullOrEmpty(strClassNum) && !string.IsNullOrEmpty(strNumSerial))
                        strAccessNo = strClassNum + "/" + strNumSerial;
                }
                // 分类号和索取号
                if (string.IsNullOrEmpty(strClassNum) && !string.IsNullOrEmpty(strAccessNo))
                {
                    int nRet = strAccessNo.IndexOf('/');
                    if (nRet != -1)
                        strClassNum = strAccessNo.Substring(0, nRet);
                }

                // 检查书目是否已存在
                if (dataTable.ContainsKey(strKey))
                {
                    Record biblioRecord = dataTable[strKey] as Record;
                    if (biblioRecord == null)
                        continue;

                    // 册记录
                    Item item = new Item();
                    biblioRecord.Items.Add(item);

                    // 这条记录的marc
                    string strMARC = biblioRecord.Biblio;
                    if (String.IsNullOrEmpty(strMARC))
                        continue;

                    // 从书目中取出价格
                    MarcRecord record = new MarcRecord(strMARC);
                    string strPrice = record.select("field[@name=010]/subfield[@name='d']").FirstContent;
                    // 册记录的一些信息
                    item.Price = strPrice;
                    item.Barcode = strBarcode;
                    item.Location = strLocation;
                    item.BookType = strBookType;
                    item.AccessNo = strAccessNo;
                }
                else
                {
                    // 新创建marc记录
                    Record biblioRecord = new Record();
                    dataTable.Add(strKey, biblioRecord);

                    MarcRecord record = null;
                    if (isHasMARC == -1)
                    {
                        record = GetMarcRecord(row);
                    }
                    else
                    {
                        //todo
                        this.getMarc(row,
                            isHasMARC,
                            strAccessNo,
                            //ref strMARC,
                            ref record);
                    }

                    // 分类号
                    MarcNodeList field690 = record.select("field[@name=690]/subfield[@name='a']");
                    if (field690.count <= 0)
                    {
                        MarcField field = new MarcField("690  " + MarcQuery.SUBFLD + "a" + strClassNum + MarcQuery.SUBFLD + "v5");
                        record.ChildNodes.insertSequence(field, InsertSequenceStyle.PreferTail);
                    }
                    else
                    {
                        string strClass = field690.FirstContent;
                        if (string.IsNullOrEmpty(strClass))
                        {
                            field690[0].Content = strClassNum;
                        }
                    }

                    // 书目
                    biblioRecord.Biblio = record.Text;

                    // 册记录
                    Item item = new Item();
                    biblioRecord.Items.Add(item);

                    string strPrice = record.select("field[@name=010]/subfield[@name='d']").FirstContent;
                    item.Price = strPrice;
                    item.Barcode = strBarcode;
                    item.Location = strLocation;
                    item.BookType = strBookType;
                    item.AccessNo = strAccessNo;
                }
            }

            // 输出册条码重复的信息
            WriteDupBarcode(sb.ToString());

            // 2021/6/4 renyh注:这一段意义不大，每次转一个excel就可以了
            //DialogResult dlgResult = MessageBox.Show(this,
            //    "转换完成，共转换'" + nCount.ToString() + "'条，是否继续转换？"
            //    + "\r\n\r\n【是】选择下一个电子表格继续转换"
            //    + "\r\n【否】不继续转换，开始输出",
            //    "是否继续转换？",
            //    MessageBoxButtons.YesNo,
            //    MessageBoxIcon.Information,
            //    MessageBoxDefaultButton.Button1);
            //if (dlgResult == DialogResult.Yes)
            //{
            //    button_open_Click(sender, e);
            //    return;
            //}

            // 输出的bdf文件名
            string strOutputFilename = this.textBox_output.Text;
            if(string.IsNullOrEmpty(strOutputFilename))
            {
                MessageBox.Show("输出文件名不能为空");
                return;
            }

            // 输出到bdf
            using (XmlTextWriter writer = new XmlTextWriter(strOutputFilename, Encoding.UTF8))
            {
                writer.Formatting = Formatting.Indented;
                writer.Indentation = 4;

                writer.WriteStartDocument();
                writer.WriteStartElement("dprms", "collection", DpNs.dprms);

                writer.WriteAttributeString("xmlns", "dprms", null, DpNs.dprms);

                int i = 0;
                int total = dataTable.Count;
                foreach (DictionaryEntry de in dataTable)
                {
                    Application.DoEvents();
                    this.Update();

                    this.toolStripStatusLabel_msg.Text = "正在输出...(" + (i++).ToString() + " / " + total.ToString() + ")";

                    // 写入 dprms:record 元素
                    writer.WriteStartElement("dprms", "record", DpNs.dprms);

                    Import(de, writer);

                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
                this.toolStripStatusLabel_msg.Text = "完成";
                dataTable.Clear();
            }
            EnableControls(true);
        }


        // 2021/6/4 把这一段处理代码专门移到一个函数里，让主函数看上去简洁些
        public int getMarc(IXLRow row,
             int isHasMARC,
             string strAccessNo,
            ref  MarcRecord record)
        {
            
            string strMARC = GetExcelCellValue(row.Cell(isHasMARC.ToString())).Trim();
            if (String.IsNullOrEmpty(strMARC))
                return -1;

             record = new MarcRecord(strMARC);

            // 删除不需要的字段
            MarcNodeList nodes = record.select("field[@name='-01' or @name='-09' or @name='-99' or @name='-98' or @name='???']");
            nodes.detach();

            nodes = record.select("field[@name='010']/subfield[@name='d']");
            foreach (MarcNode node in nodes)
            {
                node.Content = node.Content.Replace("元", "").Replace("..", ".");
            }

            if (String.IsNullOrEmpty(strAccessNo))
                strAccessNo = GetAccessNo(record);

            //（英） 伯内特 （Burnett
            nodes = record.select("field[@name>=701 and @name<=702]/subfield[@name='a']");

            char[] trimChars = new char[] { '(', ')', '（', '）', '[', ']' };

            string pattern = @"\(.*?\)|\（.*?）|\[.*?\]|\(.*?）|\（.*?\)"; // |\(.*?）|\（.*?)
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            foreach (MarcNode node in nodes)
            {
                string strContent = node.Content;

                if (regex.IsMatch(strContent))
                {
                    MatchCollection matches = regex.Matches(strContent);

                    string strCountry = matches[0].Value;

                    strContent = strContent.Replace(strCountry, "");
                    string[] parts = strContent.Split(trimChars);

                    node.Content = parts[0].Replace("/", "").Trim();
                    node.after(MarcQuery.SUBFLD + "b" + strCountry.Trim(trimChars));
                }
                else
                {
                    pattern = @"\)|\）|\]";
                    regex = new Regex(pattern, RegexOptions.IgnoreCase);
                    if (regex.IsMatch(strContent) || strContent == "主")
                    {
                        if (node.Parent != null)
                            node.Parent.detach();
                    }
                }

                // 朱琦,杨辛
                string[] names = strContent.Split(',', '，');
                for (int i = 0; i < names.Length; i++)
                {
                    if (i == 0)
                    {
                        node.Content = names[0];
                    }
                    else
                    {
                        if (node.Parent != null)
                        {
                            MarcNode parent = node.Parent;
                            try
                            {
                                parent.after("701 0" + MarcQuery.SUBFLD + "a" + names[i]);
                            }
                            catch (Exception ex)
                            {
                                string title = record.select("field[@name='200']/subfield[@name='a']").FirstContent;
                                string s = ExceptionUtil.GetDebugText(ex);
                                MessageBox.Show(title + "\r\n" + s);
                                continue;
                            }
                        }
                        else
                            record.append("701 0" + MarcQuery.SUBFLD + "a" + names[i]);
                    }
                }


                pattern = @"(等|编|原)$";
                regex = new Regex(pattern, RegexOptions.IgnoreCase);
                Match match = regex.Match(strContent);
                if (match.Success)
                {
                    string strValue = match.Value;
                    if (strValue == "等")
                        node.Content = strContent.Replace("等", "");
                    else
                    {
                        if (node.Parent != null)
                        {
                            MarcNodeList childs = node.Parent.select("subfield[@name='4']");
                            if (childs.count > 0)
                                childs[0].Content = strValue.Trim() + childs[0].Content;

                            node.Content = strContent.Replace(strValue, "");
                        }
                    }
                    continue;
                }


                // 编写 文/图
                pattern = @"(编写|文/图|编文)$"; // |图
                regex = new Regex(pattern, RegexOptions.IgnoreCase);
                match = regex.Match(strContent);
                if (match.Success)
                {
                    node.Content = strContent.Replace(match.Value, "");
                    node.after(MarcQuery.SUBFLD + "4" + match.Value);
                    continue;
                }
            }


            pattern = @"\(.*?\)|\（.*?）|\[.*?\]|\(.*?）|\（.*?\)";
            regex = new Regex(pattern, RegexOptions.IgnoreCase);
            nodes = record.select("field[@name>=701 and @name<=702]/subfield[@name='A']");
            foreach (MarcNode node in nodes)
            {
                string strContent = node.Content;

                Match match = regex.Match(strContent);
                if (match.Success)
                    strContent = strContent.Replace(match.Value, "").Trim();

                if (String.IsNullOrEmpty(strContent))
                {
                    if (node.Parent != null)
                        node.Parent.detach();

                    continue;
                }

                string[] parts = strContent.Split(new string[] { "      " }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                    strContent = parts[1].Trim();

                node.Content = strContent;
            }

            return 0;
        }

        // 获取索取号
        string GetAccessNo(MarcRecord record)
        {
            string str905d = record.select("field[@name='905']/subfield[@name='d']").FirstContent;
            if (String.IsNullOrEmpty(str905d))
                return null;

            string str905e = record.select("field[@name='905']/subfield[@name='e']").FirstContent;
            if (String.IsNullOrEmpty(str905e))
                return null;

            return str905d + "/" + str905e;
        }

        // 导出每个书目及下级册
        void Import(DictionaryEntry de, XmlTextWriter writer)
        {
            string strError = "";
            int nRet = 0;

            writer.WriteStartElement("dprms", "biblio", DpNs.dprms);
            // MarcRecord record = GetMarcRecord(de);

            Record record = de.Value as Record;

            string strMARC = record.Biblio;

            XmlDocument dom = null;
            nRet = MarcUtil.Marc2Xml(strMARC,
                "unimarc",
                out dom,
                out strError);
            if (nRet == -1)
            {
                MessageBox.Show(this, strError);
                return;
            }

            dom.DocumentElement.WriteTo(writer);
            writer.WriteEndElement();

            writer.WriteStartElement("dprms", "itemCollection", DpNs.dprms);

            foreach (Item item in record.Items)
            {
                writer.WriteStartElement("dprms", "item", DpNs.dprms);

                if (String.IsNullOrEmpty(item.Barcode) == false)
                    writer.WriteElementString("barcode", item.Barcode);
                if (String.IsNullOrEmpty(item.Price) == false)
                    writer.WriteElementString("price", item.Price);
                if (String.IsNullOrEmpty(item.Location) == false)
                    writer.WriteElementString("location", item.Location);
                if (String.IsNullOrEmpty(item.BookType) == false)
                    writer.WriteElementString("bookType", item.BookType);
                if (String.IsNullOrEmpty(item.AccessNo) == false)
                    writer.WriteElementString("accessNo", item.AccessNo);

                DateTime time = DateTime.Now;
                writer.WriteElementString("batchNo", time.ToString("yyyyMMdd"));

                writer.WriteStartElement("operations");

                writer.WriteStartElement("operation");
                writer.WriteAttributeString("name", "create");
                writer.WriteAttributeString("time", DateTimeUtil.Rfc1123DateTimeStringEx(time));
                writer.WriteAttributeString("operator", "supervisor");
                writer.WriteEndElement();

                writer.WriteEndElement();

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }


        #region 从excel获取信息

        // 从excel中获取信息
        string GetExcelCellValue(IXLCell cell)
        {
            string strStyle = cell.Style.DateFormat.Format.Replace(";@", "");
            string strContent = "";
            if (cell.DataType == XLCellValues.DateTime)
            {
                cell.Style.DateFormat.Format = "yyyyMMdd";
                strContent = cell.GetFormattedString();
            }
            else
            {
                strContent = cell.GetString();
            }
            return strContent;
        }

        // 从excel中获取书目信息
        MarcRecord GetMarcRecord(IXLRow row)
        {
            string strError = "";
            MarcNodeList nodes = null;

            MarcRecord record = new MarcRecord("?????nam0 22?????   450 ");
            foreach (ColumnHeader column in this.listView.Columns)
            {
                string text = column.Text;
                if (String.IsNullOrEmpty(text))
                    continue;

                string colName = column.Name;
                if (String.IsNullOrEmpty(colName))
                    continue;

                string strFieldName = colName.Substring(0, 3);
                string strInd = colName.Substring(3, 2).Replace("#", " ");
                string strSubfieldName = colName.Substring(5);


                var cell = row.Cell(column.Index);
                string strContent = GetExcelCellValue(cell);
                strContent = strContent.Replace("&lt;&lt;", "《").Replace("&gt;&gt;", "》").Replace("\r\n", "\n").Replace("\n", "");


                if (String.IsNullOrEmpty(strContent))
                    continue;

                if (strFieldName == "010" && strSubfieldName == "d")
                {
                    if (!string.IsNullOrEmpty(strContent) && strContent.IndexOf("CNY") == -1)
                        strContent = "CNY" + strContent;
                    else if (strContent.IndexOf("元") != -1)
                        strContent = strContent.Replace("元", "");

                    /*
                    string p = @"[A-Z]{3}\d+(\.)?\d{0,2}";
                    if(Regex.IsMatch(strContent,p))
                    {

                    }
                    else
                    {

                    }
                    */
                }
                else if (strFieldName == "010" && strSubfieldName == "a")
                {
                    int nRet = IsbnSplitter.VerifyISBN(strContent, out strError);
                    if (nRet != 0)
                        continue;
                }
                else if (strFieldName == "101" && strSubfieldName == "a")
                {
                    strContent = strContent.Replace(",", MarcQuery.SUBFLD + "a");
                }
                else if (strFieldName == "905")
                    continue;


                MarcNodeList fields = record.select(String.Format("field[@name={0}]", strFieldName));
                if (fields.count == 0)
                    record.ChildNodes.insertSequence(new MarcField((String.Format("{0}{1}{2}{3}{4}",
                        strFieldName, strInd, MarcQuery.SUBFLD, strSubfieldName, strContent))));
                else
                    fields[0].ChildNodes.insertSequence(new MarcSubfield(String.Format("{0}{1}{2}", MarcQuery.SUBFLD, strSubfieldName, strContent)));
            }

            // 100
            string strNow = DateTime.Now.ToString("yyyyMMdd");
            string strPublishYear = record.select("field[@name=210]/subfield[@name='d']").FirstContent;
            if (String.IsNullOrEmpty(strPublishYear) == false && strPublishYear.Length >= 4)
                strPublishYear = strPublishYear.Substring(0, 4);
            else
                strPublishYear = strNow.Substring(0, 4);

            record.ChildNodes.insertSequence(
                new MarcField(
                    String.Format("100  {0}a{1}d{2}    ekmy0chiy50      ea", MarcQuery.SUBFLD, strNow, strPublishYear)
                    ),
                InsertSequenceStyle.PreferTail);

            // 101
            nodes = record.select("field[@name=101]/subfield[@name='a']");
            if (nodes.count == 0)
                record.ChildNodes.insertSequence(new MarcField("101 0" + MarcQuery.SUBFLD + "achi"), InsertSequenceStyle.PreferTail);

            // 102
            record.ChildNodes.insertSequence(new MarcField((String.Format("102  {0}aCN{0}b110000", MarcQuery.SUBFLD))), InsertSequenceStyle.PreferTail);

            // 105
            record.ChildNodes.insertSequence(new MarcField("105  " + MarcQuery.SUBFLD + "ay   z   000yy"), InsertSequenceStyle.PreferTail);

            // 106
            record.ChildNodes.insertSequence(new MarcField("106  " + MarcQuery.SUBFLD + "ar"), InsertSequenceStyle.PreferTail);

            // 210
            // record.append(String.Format("210  {0}a{0}c{0}d", MarcQuery.SUBFLD));

            // 215
            nodes = record.select("field[@name=215]");
            if (nodes.count == 0)
                record.ChildNodes.insertSequence(new MarcField(String.Format("215  {0}a页{0}c图{0}d20cm", MarcQuery.SUBFLD)), InsertSequenceStyle.PreferTail);

            // 606
            nodes = record.select("field[@name=606]");
            if (nodes.count == 0)
                record.ChildNodes.insertSequence(new MarcField(String.Format("606  {0}a", MarcQuery.SUBFLD)), InsertSequenceStyle.PreferTail);

            // 690
            nodes = record.select("field[@name=690]");
            if (nodes.count == 0)
                record.ChildNodes.insertSequence(new MarcField(String.Format("690  {0}a{0}v5", MarcQuery.SUBFLD)), InsertSequenceStyle.PreferTail);

            // 701
            nodes = record.select("field[@name=200]/subfield[@name='f']");
            if (nodes.count > 0)
            {
                string strContent = nodes.FirstContent.Trim();

                strContent = strContent.Replace("…", "").Replace("[等]", "").Replace("等", "").Replace("[著]", "著").Replace("[绘]", "绘")
                    .Replace("[著编]", "编著").Replace("[编著]", "编著").Replace("[著、图]", "著图").Replace("[译]", "译")
                    .Replace("（", "(").Replace("）", ")").Replace("【", "(").Replace("】", ")").Replace("[", "(").Replace("]", ")");

                string str7XX4 = "";

                string nationPattern = @"(\(.*?\))";
                string pattern7XX4 = "(编译)|(美术)|(著绘)|(绘著)|(绘图)|(摄影)|(改编)|(改写)|(编写(?!组))|(译创)|(绘画)|(原著)|(编著)|(译注)|(选注)|(文字)|(图画)"
                    + "|(编纂)|(编辑(?!部))|(主编)|(编文)|(编选)|(选编)|(插画)|(校点、注释、整理)|(校点)|(撰文)|(撰写)|(编绘)|(注释)|(漫画)|(制作)|(编剧)|(美术插图)|(翻译设计)|(著)|(译)|(评)|(绘)|(撰)|(编(?!委会|写组|辑部))"; // |((?<!柏拉)图) // |(编(?!委会))
                // |(文)|(图)

                string[] items = Regex.Split(strContent, pattern7XX4);
                List<string> list = new List<string>(items);
                list.Remove(string.Empty);

                StringBuilder sb = new StringBuilder(256);
                for (int i = 0; i < list.Count; i++)
                {
                    if ((i % 2) == 1)
                        continue;

                    if ((i + 1) < list.Count)
                        str7XX4 = list[i + 1];
                    else
                        sb.AppendLine("越界 " + strContent);

                    string strItem = list[i].Trim();
                    strItem = Regex.Replace(strItem, @"(?<=\(.*),(?=[^(]*\))", "#");
                    string[] parts = strItem.Split(new char[] { ',', '，' });
                    foreach (string part in parts)
                    {
                        strItem = part.Replace("#", ",").Trim();

                        Add7XXField(strItem, nationPattern, record, str7XX4);
                    }
                }

                WriteHtml(sb.ToString());
            }

            // 801
            record.ChildNodes.insertSequence(new MarcField(String.Format("801 0{0}aCN{0}c{1}", MarcQuery.SUBFLD, strNow)), InsertSequenceStyle.PreferTail);

            record.append(String.Format("998  {0}a{1}{0}u{2}", MarcQuery.SUBFLD, "电子表格转入", DateTime.Now.ToString("u")));

            // record.ChildNodes.sort();

            return record;
        }

        // 处理7字段
        void Add7XXField(string strItem, string nationPattern, MarcRecord record, string str7XX4)
        {
            string[] groups = Regex.Split(strItem, nationPattern);
            if (groups.Length == 1)
            {
                record.ChildNodes.insertSequence(new MarcField(String.Format("701 0{0}a{1}{0}4{2}",
                    MarcQuery.SUBFLD,
                    strItem,
                    str7XX4)), InsertSequenceStyle.PreferTail);
            }
            else
            {
                List<string> autors = new List<string>(groups);
                autors.Remove(string.Empty);


                string strNation = "";
                string strName = "";
                string str7XXg = "";
                foreach (string item in autors)
                {
                    strItem = item.Trim();
                    if (string.IsNullOrEmpty(strItem))
                        continue;

                    if (string.IsNullOrEmpty(strNation))
                    {
                        strNation = strItem;
                        continue;
                    }
                    if (string.IsNullOrEmpty(strName))
                    {
                        strName = strItem;
                        continue;
                    }
                    if (string.IsNullOrEmpty(str7XXg))
                    {
                        str7XXg = strItem;
                        continue;
                    }
                }
                MarcField field = new MarcField("701", " 0");
                field.ChildNodes.insertSequence(new MarcSubfield("a", strName));

                if (!string.IsNullOrEmpty(strNation))
                    field.append(new MarcSubfield("c", strNation));

                if (!string.IsNullOrEmpty(str7XXg))
                    field.append(new MarcSubfield("g", str7XXg));

                field.append(new MarcSubfield("4", str7XX4));
                record.ChildNodes.insertSequence(field, InsertSequenceStyle.PreferTail);
            }
        }

        #endregion

        #endregion

        #region 输出信息

        // 输出信息
        public void WriteDupBarcode(string strHtml)
        {
            WriteHtml(this.webBrowser2, strHtml);
        }
        // 输出信息
        public void WriteHtml(string strHtml)
        {
            WriteHtml(this.webBrowser1,
                strHtml);
        }

        // 不支持异步调用
        public static void WriteHtml(WebBrowser webBrowser,
            string strHtml)
        {

            HtmlDocument doc = webBrowser.Document;

            if (doc == null)
            {
                webBrowser.Navigate("about:blank");
                doc = webBrowser.Document;
            }

            // doc = doc.OpenNew(true);
            doc.Write("<pre>");
            doc.Write(strHtml);

            // 保持末行可见
            // ScrollToEnd(webBrowser);
        }

        #endregion
    }

    // 一条完整的书目及下级册对象
    public class Record
    {
        public string Biblio = "";
        public List<Item> Items = new List<Item>();
    }

    // 册
    public class Item
    {
        public string Barcode = "";
        public string Location = "";
        public string Price = "";
        public string AccessNo = "";
        public string BookType = "";
    }
}
