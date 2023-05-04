package com.excel.generator.exporter;

import com.excel.generator.util.ExcelGeneratorConstant;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.lang.reflect.Field;
import java.util.*;

/**
 * This Class is responsible to write the data in Excel Sheet
 *
 * @param <T>
 * @author Manish
 */
@Slf4j
@Data
@Builder
public class ExcelWriter<T> {
    private List<T> reportRecords;
    private List<String> headerNames;
    private List<String> fieldNames;
    @Builder.Default
    private SXSSFWorkbook workbook = new SXSSFWorkbook();
    private SXSSFSheet currentSheet;
    private int rowIndex = 0;
    private short maxColumnIndex = 0;
    @Builder.Default
    private Map<String, HSSFColor.HSSFColorPredefined> headerNameVsColorCode = new HashMap<>();
    @Builder.Default
    private Map<String, HSSFColor.HSSFColorPredefined> headerNameVsFontColorCode = new HashMap<>();
    @Builder.Default
    private Map<String, Integer> headerVsColumnWidth = new HashMap<>();
    private CellStyle headerCellStyle;
    private CellStyle dataCellStyle;
    private boolean verticalColumnHeader;
    private ExcelMetaData<T> excelMetaData;

    /**
     * Method to write Excel data with sheet info
     *
     * @param sheetLabel
     * @return
     */
    public SXSSFWorkbook exportToExcel(String sheetLabel) {
        populateRequiredDataFromMetaData();
        long startTime = System.currentTimeMillis();
        writeDataPerSheet(reportRecords, sheetLabel);
        log.debug(":::::: Time taken to Generate Excel File  :::: {} ms", (System.currentTimeMillis() - startTime));
        return workbook;
    }

    /**
     * method to populate the require metadata for Excel writing
     */
    private void populateRequiredDataFromMetaData() {
        if (Objects.nonNull(excelMetaData)) {
            this.fieldNames = excelMetaData.getFieldNames();
            this.headerNameVsColorCode = excelMetaData.getHeaderVsColor();
            this.headerNameVsFontColorCode = excelMetaData.getHeaderVsFontColor();
            this.headerVsColumnWidth = excelMetaData.getHeaderVsColumnWidth();
            this.headerNames = excelMetaData.getHeaderNames();
        }
    }

    /**
     * Method to chop list into partial list for writing
     * upto records in one excel sheet
     *
     * @param choppedList
     * @param index
     * @return
     */
    private int recordsUpto(List<List<T>> choppedList, int index) {
        int recordsUpto = 0;
        for (int i = index; i >= 0; i--) {
            recordsUpto += choppedList.get(i).size();
        }
        return recordsUpto;
    }

    /**
     * get records from main list
     *
     * @param choppedList
     * @param index
     * @return
     */
    private int recordsFrom(List<List<T>> choppedList, int index) {
        return recordsUpto(choppedList, index - 1) + 1;
    }

    /**
     * method to write records per excel sheet
     *
     * @param records
     * @param sheetLabel
     */
    private void writeDataPerSheet(List<T> records, String sheetLabel) {
        if (null != sheetLabel) {
            currentSheet = workbook.createSheet(sheetLabel);
        } else {
            currentSheet = workbook.createSheet();
        }
        writeHeaderDataAndColumnVal(records);
    }

    /**
     * method to write data and column information in excel
     *
     * @param records
     */
    private void writeHeaderDataAndColumnVal(List<T> records) {
        currentSheet.createFreezePane(0, 1);
        writeHeaderData();
        CellStyle dataCellStyle1 = createDataCellStyle();
        if (CollectionUtils.isEmpty(records)) {
            log.debug("Empty records:: nothing to write");
            return;
        }

        records.forEach(record -> {
            Map<String, Field> allFields = getAllFields(record);
            Row row = currentSheet.createRow(rowIndex++);
            short columnIndex = 0;
            for (String fieldName : fieldNames) {
                try {
                    Field field = allFields.get(fieldName);
                    if (Objects.nonNull(field)) {
                        field.setAccessible(true);
                        Object columnValue = field.get(record);
                        createDataCell(columnIndex++, row, columnValue, dataCellStyle1);
                    }
                } catch (Exception e) {
                    log.error("Exception while Writing reportFile", e);
                }
            }
        });
    }

    /**
     * method to get all fields including super class fields names
     * using reflection
     *
     * @param record
     * @return
     */
    private Map<String, Field> getAllFields(T record) {
        Map<String, Field> allFields = new HashMap<String, Field>();
        Class<?> classObject = record.getClass();
        Field[] fields = classObject.getDeclaredFields();
        if (ArrayUtils.isNotEmpty(fields)) {
            Arrays.stream(fields).forEach(field -> {
                allFields.put(field.getName(), field);
            });
        }
        Class<?> superClassObject = classObject.getSuperclass();
        while (null != superClassObject && !superClassObject.equals(java.lang.Object.class)) {
            Field[] superClassFields = superClassObject.getDeclaredFields();
            if (ArrayUtils.isNotEmpty(superClassFields)) {
                Arrays.stream(superClassFields).forEach(field -> {
                    allFields.put(field.getName(), field);
                });
                superClassObject = superClassObject.getSuperclass();
            }
        }
        return allFields;
    }

    /**
     * method to create Data cell
     *
     * @return
     */
    private CellStyle createDataCellStyle() {
        if (Objects.nonNull(this.dataCellStyle)) {
            return dataCellStyle;
        }
        Font dataFont = workbook.createFont();
        dataFont.setFontHeightInPoints(ExcelGeneratorConstant.FONT_HEIGHT_IN_POINTS);
        dataFont.setFontName(ExcelGeneratorConstant.GE_INSPIRA_SANS);
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.GENERAL);
        style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        style.setFont(dataFont);
        style.setWrapText(true);
        allBordersWithBlackColor(style);
        return style;
    }

    /**
     * Method to write header data
     */
    private void writeHeaderData() {
        final Row headerRow = currentSheet.createRow(rowIndex++);
        short headerColumnIndex = 0;
        if (CollectionUtils.isEmpty(headerNames)) {
            log.debug("Empty Header");
            return;
        }
        for (String headerName : headerNames) {
            currentSheet.setColumnWidth(headerColumnIndex, getColumnWidth(headerName));
            createHeaderCell(getHeaderCellStyle(headerName), headerRow, headerName, headerColumnIndex++);
        }
        if (0 == maxColumnIndex) {
            maxColumnIndex = headerColumnIndex;
        }
    }

    /**
     * method to get the column width
     *
     * @param headerName
     * @return
     */
    private int getColumnWidth(String headerName) {
        int columnWidth = 15;
        if (MapUtils.isNotEmpty(headerVsColumnWidth) && Objects.nonNull(headerVsColumnWidth.get(headerName))) {
            columnWidth = headerVsColumnWidth.getOrDefault(headerName, columnWidth);
        }
        return columnWidth * ExcelGeneratorConstant.ONE_CHAR_WIDTH;
    }

    /**
     * method to get the default Header Cell style
     *
     * @param headerName
     * @return
     */
    private CellStyle getHeaderCellStyle(String headerName) {
        if (Objects.nonNull(this.headerCellStyle)) {
            return headerCellStyle;
        }
        Font font = workbook.createFont();
        font.setFontHeightInPoints(ExcelGeneratorConstant.FONT_HEIGHT_IN_POINTS);
        font.setFontName(ExcelGeneratorConstant.GE_INSPIRA_SANS);
        font.setBold(true);
        short headerFontColor = HSSFColor.HSSFColorPredefined.BLACK.getIndex();
        if (MapUtils.isNotEmpty(headerNameVsFontColorCode)) {
            HSSFColor.HSSFColorPredefined hssfColorPredefined = headerNameVsFontColorCode.get(headerName);
            if (Objects.nonNull(hssfColorPredefined)) {
                headerFontColor = hssfColorPredefined.getIndex();
            }
        }
        font.setColor(headerFontColor);
        short headerForGroundColor = HSSFColor.HSSFColorPredefined.YELLOW.getIndex();
        if (MapUtils.isNotEmpty(headerNameVsColorCode)) {
            HSSFColor.HSSFColorPredefined hssfColorPredefined = headerNameVsColorCode.get(headerName);
            if (Objects.nonNull(hssfColorPredefined)) {
                headerForGroundColor = hssfColorPredefined.getIndex();
            }
        }
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setFillForegroundColor(headerForGroundColor);
        cellStyle.setFont(font);
        setCellAlignment(cellStyle);
        cellStyle.setWrapText(true);
        allBordersWithBlackColor(cellStyle);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    /**
     * method to set the Cell Alignment
     *
     * @param cellStyle
     */
    private void setCellAlignment(CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    /**
     * method to set the Border as Black in cell
     *
     * @param cellStyle
     */
    private void allBordersWithBlackColor(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        cellStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
    }

    /**
     * method to create a Data cell
     *
     * @param columnIndex
     * @param row
     * @param columnValue
     * @param dataCellStyle
     * @return
     */
    private Cell createDataCell(short columnIndex, final Row row, final Object columnValue, CellStyle dataCellStyle) {
        Cell columnDataValue = row.createCell(columnIndex);
        columnDataValue.setCellStyle(dataCellStyle);
        if (null != columnValue) {
            columnDataValue.setCellValue(columnValue.toString());
        }
        return columnDataValue;
    }

    /**
     * method to create a Header cell
     *
     * @param headerCellStyle
     * @param row
     * @param headerText
     * @param headerColumnIndex
     */
    private void createHeaderCell(CellStyle headerCellStyle, Row row, String headerText, Short headerColumnIndex) {
        Cell cell = row.createCell(headerColumnIndex);
        cell.setCellStyle(headerCellStyle);
        cell.setCellValue(new HSSFRichTextString(headerText));
    }



}

