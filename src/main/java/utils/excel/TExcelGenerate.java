package utils.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class TExcelGenerate<T> {

    private final Collection<T> entities;
    private final Workbook workbook;

    private final WorkBookStyle wbStyle;
    private Sheet sheet;
    private final String sheetTitle;
    private final LinkedHashMap<Integer, ExcelColumn> sheetHeaders;
    private final LinkedHashMap<String, List<Object>> sheetContents;
    private boolean stylizedRow;


    public TExcelGenerate(Collection<T> entities, Class<T> clazz,ExcelCompatibility compatibility) {
        this.entities = entities;
        this.sheetHeaders = new LinkedHashMap<>();
        this.sheetContents = new LinkedHashMap<>();
        this.wbStyle = new WorkBookStyle(compatibility);

        if(compatibility == ExcelCompatibility.COMPATIBILITY_2003) this.workbook = new HSSFWorkbook();
        else this.workbook = new XSSFWorkbook();

        this.stylizedRow = true;
        ExcelSheet annotation = clazz.getAnnotation(ExcelSheet.class);
        sheetTitle = annotation.title();
        int index = 0;
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn fieldAnnotation = field.getAnnotation(ExcelColumn.class);
            if(fieldAnnotation!=null){
                String description = fieldAnnotation.description();
                sheetHeaders.put(index++,fieldAnnotation);
                sheetContents.put(description, new ArrayList<>());
            }
        }
    }

    public Workbook create(String fileName) throws IOException {
        writeTitleAndHeader();
        writeRowData();
        writeFooter();
        String name = "C:/Users/mathe/OneDrive/Área de Trabalho/"+fileName;
        FileOutputStream outputStream = new FileOutputStream(name);
        workbook.write(outputStream);
        workbook.close();
        return workbook;
    }

    private void writeTitleAndHeader() {
        sheet = workbook.createSheet(sheetTitle.toLowerCase(Locale.ROOT));
        Row row = sheet.createRow(0);
        CellStyle style = wbStyle.getHeaderStyle(workbook);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, sheetHeaders.size()-1));
        for(int i = 0; i < sheetHeaders.size()-1; i++) {
            if(i>0) createCell(row,i,0,null,style);
            else createCell(row,i,0,sheetTitle,style);
        }
        Row headerRow = sheet.createRow(1);
        for(int i = 0; i < sheetHeaders.size(); i++)
            createCell(headerRow,i,0,sheetHeaders.get(i).description(),style);

    }


    private void writeRowData(){
        setSheetContents();
        int maxRow = sheetContents.values().stream().map(List::size).max(Integer::compareTo).orElse(0);

        for (int rowNum = 2 ; rowNum < maxRow + 2 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            int styleIndex = !stylizedRow ?1:rowNum%2;
            for (int colNum = 0; colNum < sheetHeaders.size(); colNum++) {
                int colWidth = sheetHeaders.get(colNum).width();
                HorizontalAlignment alignmentCell = sheetHeaders.get(colNum).alignment().alignment;
                ColumnType columnType = sheetHeaders.get(colNum).type();
                String pattern = sheetHeaders.get(colNum).pattern();
                String outputPattern = sheetHeaders.get(colNum).outputPattern();
                CellStyle style = wbStyle.getRowStyle(workbook,alignmentCell,columnType,outputPattern.toLowerCase()).get(styleIndex);
                Object value = sheetContents.get(sheetHeaders.get(colNum).description()).get(0);

                if(columnType == ColumnType.DATE){
                    SimpleDateFormat dateFormat = new SimpleDateFormat(pattern);
                    try{
                        value = dateFormat.parse(value.toString());
                    }catch (ParseException e){
                        continue;
                    }
                }


                sheetContents.get(sheetHeaders.get(colNum).description()).remove(0);
                createCell(row,colNum,colWidth,value,style);
            }
        }

    }

    @SuppressWarnings("java:S3011")
    private void setSheetContents(){
        entities.forEach(entity->{
            for (Field field : entity.getClass().getDeclaredFields()) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                field.setAccessible(true);
                try {
                    sheetContents.get(annotation.description()).add(field.get(entity));
                } catch (IllegalAccessException ignored) {
                    continue;
                }
                field.setAccessible(false);
            }
        });
    }

    private void writeFooter(){
        String dateNow = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/y"));
        String footer = String.format("Gerado automaticamente pelo sistema %s - %s (%s) ","ExcelGenerator","Eleodoro Dev",dateNow);
        int lastRow = sheet.getLastRowNum() + 1;
        Row footerRow = sheet.createRow(lastRow);
        sheet.addMergedRegion(new CellRangeAddress(lastRow, lastRow, 0, sheetHeaders.size()-1));

        for(int colNum = 0; colNum < sheetHeaders.size();colNum++){
            Cell footerCell = footerRow.createCell(colNum);
            if (colNum == 0) footerCell.setCellValue(footer);
            footerCell.setCellStyle(wbStyle.getFooterStyle(workbook));
        }
    }

    private void createCell(Row row, int columnCount,int columnWidth, Object value, CellStyle style) {
        if(columnWidth == 0) sheet.autoSizeColumn(columnCount);
        else sheet.setColumnWidth(columnCount,columnWidth * 256);
        Cell cell = row.createCell(columnCount);

        if (value instanceof Integer)
            cell.setCellValue((Integer) value);
        else if (value instanceof Boolean)
            cell.setCellValue((Boolean) value);
        else if(value instanceof Date)
            cell.setCellValue((Date) value);
        else if(value instanceof BigDecimal)
            cell.setCellValue(new BigDecimal(value.toString()).doubleValue());
        else
            cell.setCellValue(((String) value));

        cell.setCellStyle(style);
    }

    @Target(ElementType.TYPE)
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelSheet {
        String title();
        String[] columnOrder() default {};

    }
    @Target({ElementType.METHOD,ElementType.FIELD})
    @Retention(RetentionPolicy.RUNTIME)
    public @interface ExcelColumn{
        String description();
        AlignmentCell alignment() default AlignmentCell.LEFT;
        int width() default 0;
        ColumnType type() default ColumnType.DEFAULT;
        String pattern() default "yyyyMMdd";
        String outputPattern() default "dd/MM/yyyy";
    }
    public enum AlignmentCell {
        CENTER(HorizontalAlignment.CENTER),
        LEFT(HorizontalAlignment.LEFT),
        RIGHT(HorizontalAlignment.RIGHT);
        public final HorizontalAlignment alignment;

        AlignmentCell(HorizontalAlignment alignment) {
            this.alignment = alignment;
        }
    }
    public enum ColumnType{
        DEFAULT,
        DATE,
        VALUE
    }

    public enum ExcelCompatibility {
        COMPATIBILITY_2003,
        COMPATIBILITY_2007,
    }
}
