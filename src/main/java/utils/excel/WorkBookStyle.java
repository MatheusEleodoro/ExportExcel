package utils.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.excel.TExcelGenerate.ExcelCompatibility;

import java.util.Arrays;
import java.util.List;

class WorkBookStyle {
    protected static final byte[] DEFAULT_COLOR = {(byte) 241, 86, 32};
    public static final String VALUE_PATTERN = "#,##0.00";
    public static final String FONT_NAME = "Arial";
    protected ExcelCompatibility compatibility;

    public WorkBookStyle(ExcelCompatibility compatibility) {
        this.compatibility = compatibility;
    }

    public CellStyle getHeaderStyle(Workbook workbook){

        if(compatibility == ExcelCompatibility.COMPATIBILITY_2007){
            XSSFWorkbook wb = ((XSSFWorkbook) workbook);
            XSSFCellStyle style = wb.createCellStyle();
            XSSFFont font = wb.createFont();
            font.setBold(true);
            font.setFamily(FontFamily.MODERN);
            font.setFontHeight(14);
            font.setColor(IndexedColors.WHITE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(new XSSFColor(DEFAULT_COLOR,null));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setAlignment(HorizontalAlignment.CENTER);
            return style;
        }else {
            HSSFWorkbook wb = ((HSSFWorkbook) workbook);
            HSSFCellStyle style = wb.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setBold(true);
            font.setFontName(FONT_NAME);
            font.setFontHeightInPoints((short)14);
            font.setColor(IndexedColors.WHITE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setAlignment(HorizontalAlignment.CENTER);
            return style;
        }
    }

    public List<CellStyle> getRowStyle(Workbook workbook, HorizontalAlignment alignment, TExcelGenerate.ColumnType type, String pattern){

        if(compatibility == ExcelCompatibility.COMPATIBILITY_2007){
            XSSFWorkbook wb = ((XSSFWorkbook) workbook);
            DataFormat wbFormat = wb.createDataFormat();
            List<CellStyle> rowStyles = Arrays.asList(wb.createCellStyle(), wb.createCellStyle());
            XSSFFont font = wb.createFont();
            font.setFontName(FONT_NAME);
            font.setFontHeight(12);

            byte[] rgb = {(byte) 242, (byte) 242, (byte) 242};

            setDataFormatter(type, pattern, wbFormat, rowStyles);

            rowStyles.get(0).setAlignment(alignment);
            rowStyles.get(1).setAlignment(alignment);

            rowStyles.get(0).setFont(font);
            rowStyles.get(1).setFont(font);

            ((XSSFCellStyle) rowStyles.get(0)).setFillForegroundColor(new XSSFColor(rgb,null));
            rowStyles.get(0).setFillPattern(FillPatternType.SOLID_FOREGROUND);

            return rowStyles;
        }else {
            HSSFWorkbook wb = ((HSSFWorkbook) workbook);
            DataFormat wbFormat = wb.createDataFormat();
            List<CellStyle> rowStyles = Arrays.asList(wb.createCellStyle(), wb.createCellStyle());
            HSSFFont font = wb.createFont();
            font.setFontName(FONT_NAME);
            font.setFontHeightInPoints(((short) 12));

            setDataFormatter(type, pattern, wbFormat, rowStyles);

            rowStyles.get(0).setAlignment(alignment);
            rowStyles.get(1).setAlignment(alignment);

            rowStyles.get(0).setFont(font);
            rowStyles.get(1).setFont(font);

            rowStyles.get(0).setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            rowStyles.get(0).setFillPattern(FillPatternType.SOLID_FOREGROUND);

            return rowStyles;
        }

    }

    private void setDataFormatter(TExcelGenerate.ColumnType type, String pattern, DataFormat wbFormat, List<CellStyle> rowStyles) {
        if(type == TExcelGenerate.ColumnType.VALUE){
            rowStyles.get(0).setDataFormat(wbFormat.getFormat(VALUE_PATTERN));
            rowStyles.get(1).setDataFormat(wbFormat.getFormat(VALUE_PATTERN));
        } else if (type == TExcelGenerate.ColumnType.DATE) {
            rowStyles.get(0).setDataFormat(wbFormat.getFormat(pattern));
            rowStyles.get(1).setDataFormat(wbFormat.getFormat(pattern));
        }
    }

    public CellStyle getFooterStyle(Workbook workbook){
        if(compatibility == ExcelCompatibility.COMPATIBILITY_2007){
            XSSFWorkbook wb = ((XSSFWorkbook) workbook);
            XSSFCellStyle style = wb.createCellStyle();
            XSSFFont font = wb.createFont();
            font.setBold(true);
            font.setFamily(FontFamily.MODERN);
            font.setFontHeight(9);
            font.setColor(IndexedColors.WHITE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(new XSSFColor(DEFAULT_COLOR,null));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setAlignment(HorizontalAlignment.CENTER);
            return style;
        }else{
            HSSFWorkbook wb = ((HSSFWorkbook) workbook);
            HSSFCellStyle style = wb.createCellStyle();
            HSSFFont font = wb.createFont();
            font.setBold(true);
            font.setFontName(FONT_NAME);
            font.setFontHeightInPoints(((short) 9));
            font.setColor(IndexedColors.WHITE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setAlignment(HorizontalAlignment.CENTER);
            return style;
        }

    }
}
