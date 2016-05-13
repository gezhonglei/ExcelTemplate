using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportTemplate.Export.Writer.CellRender
{
    public class CommentRender : ICellRender
    {
        public string Author;
        public string Comment;

        public void Render(ICell cell)
        {
            NPOIExcelUtil.AddComment(cell, Author, Comment);
        }
    }

    public class HyperLinkRender : ICellRender
    {
        public string LinkType;
        public string Address;

        protected HyperlinkType GetLinkType(string type)
        {
            HyperlinkType linkType = HyperlinkType.Unknown;
            if ("auto".Equals(type.ToLower()))
            {
                linkType = NPOIExcelUtil.GetLinkTypeByData(Address);
                linkType = linkType == HyperlinkType.Unknown ? HyperlinkType.Url : linkType;
            }
            else if (!Enum.TryParse(type, true, out linkType))
            {
                //默认URL类型（Excel2003用Unknown时会报引用为空错误；Excel2007中如Address与LinkType不匹配会报错，即使调整为Unknown导出Excel打开异常）
                linkType = HyperlinkType.Url; //: HyperlinkType.Unknown;
            }
            return linkType;
        }

        public void Render(ICell cell)
        {
            //NPOIExcelUtil.AddHyperLink(cell, GetLinkType(LinkType), Address);
            HyperlinkType linkType = GetLinkType(LinkType);
            IWorkbook book = cell.Sheet.Workbook;
            IHyperlink link = book.GetCreationHelper().CreateHyperlink(linkType);
            link.Address = Address;
            cell.Hyperlink = link;

            IFont linkFont = book.CreateFont();
            linkFont.Underline = FontUnderlineType.Single;
            linkFont.Color = IndexedColors.Blue.Index;
            ICellStyle linkStyle = cell.CellStyle ?? book.CreateCellStyle();
            linkStyle.SetFont(linkFont);
            cell.CellStyle = linkStyle;
        }
    }

    public class DropDownListRender : ICellRender
    {
        public string[] ValueList;
        public Location FillArea;

        public void Render(ICell cell)
        {
            if (ValueList.Length > 0)
            {
                ISheet sheet = cell.Sheet;
                IDataValidationHelper helper = sheet.GetDataValidationHelper();
                IDataValidationConstraint dvconstraint = helper.CreateExplicitListConstraint(ValueList);
                CellRangeAddressList rangeList = new CellRangeAddressList(FillArea.RowIndex, FillArea.RowIndex + FillArea.RowCount - 1,
                    FillArea.ColIndex, FillArea.ColIndex + FillArea.ColCount - 1);
                //DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
                //exSheet.AddValidationData(new HSSFDataValidation(rangeList, constraint)); 
                sheet.AddValidationData(helper.CreateValidation(dvconstraint, rangeList));
            }
        }
    }

    public class PictureRender : ICellRender
    {
        public byte[] PicSource;
        public PictureType PicType;
        public Location PicArea;

        public void Render(ICell cell)
        {
            NPOI.SS.Util.CellRangeAddress region = NPOIExcelUtil.GetRange(cell);
            ISheet sheet = cell.Sheet;
            IDrawing draw = sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = region != null ?
                draw.CreateAnchor(20, 20, 0, 0, region.FirstColumn, region.FirstRow, region.LastColumn + 1, region.LastRow + 1) :
                PicArea != null ?
                draw.CreateAnchor(20, 20, 0, 0, PicArea.ColIndex, PicArea.RowIndex, PicArea.ColIndex + PicArea.ColCount, PicArea.RowIndex + PicArea.RowCount):
                draw.CreateAnchor(20, 20, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 1, cell.RowIndex + 1);
            draw.CreatePicture(anchor, sheet.Workbook.AddPicture(PicSource, PictureType.JPEG));
        }
    }
}
