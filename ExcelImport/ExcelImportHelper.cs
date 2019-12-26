using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelImport
{
    public class ExcelImportHelper<T> where T : ExcelData, new()
    {
        //--------------------Members--------------------

        #region 成员属性

        /// <summary>
        /// 导入文件路径
        /// </summary>
        protected string _importFilePath;

        /// <summary>
        /// 导出文件路径
        /// </summary>
        protected string _exportFilePath;

        /// <summary>
        /// 导出文件路径
        /// </summary>
        public string ExportFilePath
        {
            get
            {
                return _exportFilePath;
            }
            set
            {
                _exportFilePath = value;
            }
        }

        /// <summary>
        /// (模板)列头文本信息
        /// </summary>
        protected List<string> _headerTexts
        {
            get => TemplateObj.GetHeaderProperty().Keys.ToList();

        }

        /// <summary>
        /// 斜杠
        /// </summary>
        private readonly char _slash = '/';

        /// <summary>
        /// 反斜杠
        /// </summary>
        private readonly char _backSlash = '\\';

        /// <summary>
        /// 分隔符
        /// </summary>
        private char _separator
        {
            get
            {
                return _backSlash;
            }
        }


        private T TemplateObj = new T();

        /// <summary>
        /// Excel 数据属性字典
        /// </summary>
        protected Dictionary<string, string> _dataFields => TemplateObj.GetHeaderProperty();

        private bool _IsxlsxFile
        {
            get
            {
                if (_importFilePath.EndsWith("xlsx", false, null))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        #endregion


        //--------------------Methods--------------------

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelImportHelper()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath">导入文件路径</param>
        public ExcelImportHelper(string filePath)
        : this()
        {
            _importFilePath = filePath;
        }

        #endregion

        /// <summary>
        /// 工作薄
        /// </summary>
        protected IWorkbook _workbook;

        /// <summary>
        /// 工作表
        /// </summary>
        protected ISheet _sheet;

        /// <summary>
        /// 行
        /// </summary>
        protected IRow _row;

        /// <summary>
        /// 初始化工作薄
        /// </summary>
        private void InitWorkbook(bool isTemplate)
        {
            _workbook = new HSSFWorkbook();
            //创建表
            _sheet = _workbook.CreateSheet();

            //创建第一行并填写数据
            _row = _sheet.CreateRow(0);
            int headerCount = 0;
            if (isTemplate)
            {
                SetRowTexts(_row, 0, _headerTexts.ToArray());
                headerCount = _headerTexts.Count;
            }
            else
            {
                SetRowTexts(_row, 0, _dataFields.Keys.ToArray());
                headerCount = _dataFields.Count;
            }

            //设置列头样式
            ICellStyle cellStyle = GetCellStyle(_workbook, "Header");
            for (int i = 0; i < headerCount; i++)
            {
                ICell cell = _row.GetCell(i);
                _row.GetCell(i).CellStyle = cellStyle;
            }

            SetRowComment(_sheet, _row);
        }

        /// <summary>
        /// 设置列头备注
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        private void SetRowComment(ISheet sheet, IRow row)
        {
            List<string> comments = new List<string>();

            foreach (var prop in TemplateObj.GetType().GetProperties())
            {
                var attr = prop.GetCustomAttributes(typeof(ExcelHeaderAttribute), false).FirstOrDefault() as ExcelHeaderAttribute;
                comments.Add(attr.Comment);
            }

            for (int i = 0; i < comments.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(comments[i]))
                {
                    row.GetCell(i).CellComment = GetCellComment(sheet, 5, 8, comments[i], "someone");
                }
            }
        }

        /// <summary>
        /// 填写数据
        /// </summary>
        /// <param name="datas"></param>
        protected virtual void FillWorkbook(IEnumerable<T> datas)
        {
            int index = 0;
            foreach (var data in datas)
            {
                List<string> fields = new List<string>();
                foreach (var field in typeof(T).GetFields())
                {
                    object value = field.GetValue(data);
                    if (value == null)
                    {
                        value = string.Empty;
                    }
                    fields.Add(value.ToString());
                }
                _row = _sheet.CreateRow(++index);
                SetRowTexts(_row, 0, fields.ToArray());
            }
        }

        /// <summary>
        /// 导出模板
        /// </summary>
        public void ExportTemplate()
        {
            if (!CheckExportFilePath(_exportFilePath))
            {
                return;
            }
            using (FileStream fs = new FileStream(_exportFilePath, FileMode.OpenOrCreate, FileAccess.Write))
            {
                InitWorkbook(true);

                //将数据写入至 Excel 文件
                _workbook.Write(fs);
            }

        }

        public void ExportTemplate(string path)
        {
            if (!CheckExportFilePath(path))
            {
                return;
            }
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
                InitWorkbook(true);

                //将数据写入至 Excel 文件
                _workbook.Write(fs);
            }

        }

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="datas">IP 摄像机 Excel 数据</param>
        public void ExportData(IEnumerable<T> datas)
        {
            if (!CheckExportFilePath(_exportFilePath))
            {
                return;
            }
            using (FileStream fs = new FileStream(_exportFilePath, FileMode.OpenOrCreate, FileAccess.Write))
            {
                InitWorkbook(false);
                FillWorkbook(datas);

                //将数据写入至 Excel 文件
                _workbook.Write(fs);
            }

        }

        /// <summary>
        /// 填写行文本信息
        /// </summary>
        /// <param name="row">行对象</param>
        /// <param name="col">开始列</param>
        /// <param name="texts">文本信息集合</param>
        private void SetRowTexts(IRow row, int col, string[] texts)
        {
            foreach (string text in texts)
            {
                row.CreateCell(col).SetCellValue(text);
                col++;
            }
        }

        /// <summary>
        /// 获取文本所在列索引
        /// </summary>
        /// <param name="row">行对象</param>
        /// <param name="text">文本信息</param>
        /// <param name="index">列索引</param>
        /// <returns></returns>
        private bool TryGetColumnIndexInRowByText(IRow row, string text, ref int index)
        {
            bool result = false;
            foreach (ICell cell in row.Cells)
            {
                if (cell.StringCellValue == text)
                {
                    index = cell.ColumnIndex;
                    result = true;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// 获取单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private ICellStyle GetCellStyle(IWorkbook workbook, string type)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            switch (type)
            {
                case "Header":
                    font.Boldweight = (short)FontBoldWeight.Bold;
                    cellStyle.SetFont(font);
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;

                    break;
                default:
                    break;
            }
            return cellStyle;
        }

        /// <summary>
        /// 获取批注
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="col2"></param>
        /// <param name="row2"></param>
        /// <param name="content"></param>
        /// <param name="author"></param>
        /// <returns></returns>
        private IComment GetCellComment(ISheet sheet, int col2, int row2, string content, string author)
        {
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, col2, row2);
            anchor.AnchorType = AnchorType.DontMoveAndResize;
            IComment comment = drawing.CreateCellComment(anchor);
            comment.String = new HSSFRichTextString(content);
            comment.Author = author;
            return comment;
        }

        /// <summary>
        /// 设置数据有效性
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="explictList"></param>
        private void SetDataValidation(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, string[] explictList)
        {
            //设置限制的区域
            CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
            //设置限制文字
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(explictList);
            //区域与文字设置数据有效性对象
            IDataValidation validation = new HSSFDataValidation(regions, constraint);
            //在指定列表中加入该限制
            sheet.AddValidationData(validation);
        }

        /// <summary>
        /// 检查导出文件路径是否可用
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private bool CheckExportFilePath(string filePath)
        {
            bool result = true;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                result = false;
            }
            return result;
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <param name="sources"></param>
        /// <returns></returns>
        public List<T> Import()
        {
            List<T> result = new List<T>();
            using (FileStream fs = new FileStream(_importFilePath, FileMode.Open, FileAccess.Read))
            {
                _workbook = GetWorkbookBySuffix(fs);
                for (int sheetIndex = 0; sheetIndex < _workbook.NumberOfSheets; sheetIndex++)
                {
                    _sheet = _workbook.GetSheetAt(sheetIndex);
                    var datas = GetDatasFromSheet(_sheet);
                    if (datas != null)
                    {
                        result.AddRange(GetDatasFromSheet(_sheet));
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public List<T> Import(string path)
        {
            _importFilePath = path;
            List<T> result = new List<T>();
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                _workbook = GetWorkbookBySuffix(fs);
                for (int sheetIndex = 0; sheetIndex < _workbook.NumberOfSheets; sheetIndex++)
                {
                    _sheet = _workbook.GetSheetAt(sheetIndex);
                    var datas = GetDatasFromSheet(_sheet);
                    if (datas != null)
                    {
                        result.AddRange(GetDatasFromSheet(_sheet));
                    }
                }
            }
            return result;
        }

        private IEnumerable<T> GetDatasFromSheet(ISheet sheet)
        {
            List<T> result = new List<T>();
            //获取列头文本与列索引关系
            IRow row = sheet.GetRow(0);

            if (row == null)
                return null;

            Dictionary<string, int> propertyDic = GeneratePropertyDic(row);

            //依次填充结构体
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row == null)
                    continue;

                T data = new T();
                if (TryFillData(row, propertyDic, ref data))
                {
                    result.Add(data);
                }
            }
            return result;
        }

        /// <summary>
        /// 生成属性字典
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        protected Dictionary<string, int> GeneratePropertyDic(IRow row)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();
            foreach (var cell in row.Cells)
            {
                if (!result.ContainsKey(cell.StringCellValue))
                {
                    result.Add(cell.StringCellValue, cell.ColumnIndex);
                }
            }
            return result;
        }

        /// <summary>
        /// 单元格值转换成字符串
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private string CellValueConvert2String(ICell cell)
        {
            string result;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    result = cell.NumericCellValue.ToString();
                    break;
                default:
                    result = cell.StringCellValue;
                    break;
            }
            return result;
        }

        /// <summary>
        /// 填充IPCameraData结构体
        /// </summary>
        /// <param name="row">行对象</param>
        /// <param name="headerDic">列头字典</param>
        /// <param name="data">IPCameraData结构体</param>
        /// <returns>是否填充</returns>
        protected bool TryFillData(IRow row, Dictionary<string, int> headerDic, ref T data)
        {
            //导入逻辑右边，设备名称空当作子码流，故不做判断
            //判断设备名称是否为空
            //if (string.IsNullOrWhiteSpace(CellValueConvert2String(row.GetCell(headerDic["设备名称"], MissingCellPolicy.CREATE_NULL_AS_BLANK))))
            //    return false;

            //进行分类填充
            foreach (var item in headerDic)
            {
                string value = CellValueConvert2String(row.GetCell(item.Value, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                if (_dataFields.ContainsKey(item.Key))
                {
                    var prop = typeof(T).GetProperty(_dataFields[item.Key]);
                    if (prop != null)
                    {
                        prop.SetValue(data, value, null);
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// 通过后缀名获取 Workbook
        /// </summary>
        /// <returns></returns>
        private IWorkbook GetWorkbookBySuffix(FileStream fs)
        {
            if (_IsxlsxFile)
            {
                return new XSSFWorkbook(fs);
            }
            else
            {
                return new HSSFWorkbook(fs);
            }
        }

    }
}
