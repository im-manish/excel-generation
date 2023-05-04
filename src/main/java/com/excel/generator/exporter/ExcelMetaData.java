package com.excel.generator.exporter;

import com.excel.generator.annotation.ExcelInfo;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Excel Metadata class that handle all the metainfo information
 * require for excel generation
 * @param <T>
 * @author Manish
 */
@Slf4j
@Data
public class ExcelMetaData<T> {
    private  Class<T> classObject;
    private  List<MetaInfo> metaInfoList = new ArrayList<MetaInfo>();
    private  List<String> headerNames = new ArrayList<String>();
    private  List<String> fieldNames = new ArrayList<String>();
    private boolean includeSuperClassFields;
    private final Map<String, HSSFColor.HSSFColorPredefined> headerVsColor = new HashMap<>();
    private final Map<String, Integer> headerVsColumnWidth = new HashMap<>();
    private final Map<String, HSSFColor.HSSFColorPredefined> headerVsFontColor = new HashMap<>();

    /**
     * constructor with classObject
     * @param classObject
     */
    public ExcelMetaData(Class<T> classObject) {
        this.classObject = classObject;
        populateExcelMetaData();
    }

    /**
     * constructor with class object and boolean value that say do we need to include
     * superclass of the VOs or not
     * @param classObject
     * @param includeSuperClassFields
     */
    public ExcelMetaData(Class<T> classObject, boolean includeSuperClassFields) {
        this.classObject = classObject;
        this.includeSuperClassFields = includeSuperClassFields;
        populateExcelMetaData();
    }

    /**
     * constructor with List of {@link MetaInfo}
     * @param metaInfoList
     */
    public ExcelMetaData(List<MetaInfo> metaInfoList) {
        this.metaInfoList = metaInfoList;
        prepareMetaInfoDetails();
    }

    /**
     * getter of header
     * @return List of Header names
     */

    public List<String> getHeaderNames() {
        return headerNames;
    }

    /**
     * getter of Header vs Color
     * @return
     */
    public Map<String, HSSFColor.HSSFColorPredefined> getHeaderVsColor() {
        return this.headerVsColor;
    }

    /**
     * getter of HeaderVs ColumnWidth
     * @return
     */
    public Map<String, Integer> getHeaderVsColumnWidth() { return headerVsColumnWidth;   }

    /**
     * getter of fieldName
     * @return
     */
    public List<String> getFieldNames() {
        return fieldNames;
    }

    /**
     * method to populate MetaData
     */
    public  void populateExcelMetaData() {
        Field[] fields = classObject.getDeclaredFields();
        if (includeSuperClassFields) {
            fields = getAllFields(classObject.getSuperclass().getDeclaredFields(), classObject.getDeclaredFields());
        }
        Arrays.stream(fields).filter(field -> field.isAnnotationPresent(ExcelInfo.class)).forEach(field -> {
            try {
                ExcelInfo excelInfo = field.getAnnotation(ExcelInfo.class);
                MetaInfo metaInfo = MetaInfo.builder().fieldName(field.getName())
                        .headerColor(excelInfo.headerColor())
                        .sortIndex(excelInfo.sortIndex())
                        .headerName(excelInfo.headerName())
                        .headerFontColor(excelInfo.headerFontColor())
                        .columnWidth(excelInfo.columnWidth())
                        .build();
                metaInfoList.add(metaInfo);
            } catch (Exception ex) {
                log.error("Error while populating excelMetaData", ex);
            }
        });
        prepareMetaInfoDetails();
    }

    /**
     * method to prepare the MetaInfo Details
     */
    private void prepareMetaInfoDetails() {
        Collections.sort(metaInfoList, MetaInfo.METADATA_COMPARATOR);
        metaInfoList.stream().forEach(metaInfo -> {
            String header = metaInfo.headerName;
            if (StringUtils.isNotBlank(header)) {
                headerNames.add(header);
            }
            fieldNames.add(metaInfo.fieldName);
            headerVsColor.put(metaInfo.headerName, metaInfo.headerColor);
            headerVsFontColor.put(metaInfo.headerName, metaInfo.headerFontColor);
            headerVsColumnWidth.put(metaInfo.headerName, metaInfo.columnWidth);
        });
    }

    /**
     * Static Inner class that handle Excel MetaInfo
     */
    @AllArgsConstructor
    @NoArgsConstructor
    @Data
    @Builder
    public static class MetaInfo {
        private String fieldName;
        private String headerName;
        private double sortIndex;
        private HSSFColor.HSSFColorPredefined headerColor;
        private Integer columnWidth;
        public static Comparator<MetaInfo> METADATA_COMPARATOR = Comparator.comparing(MetaInfo::getSortIndex);
        private HSSFColor.HSSFColorPredefined headerFontColor;
    }

    /**
     * method to get all field via reflection
     * @param baseClassFields
     * @param childClassFields
     * @return
     */
    private Field[] getAllFields(Field[] baseClassFields, Field[] childClassFields) {
        int size = baseClassFields.length + childClassFields.length;
        Field[] fields = new Field[size];
        AtomicInteger atomicInteger = new AtomicInteger(0);
        Arrays.stream(baseClassFields).forEach(baseClassField -> {
            fields[atomicInteger.getAndIncrement()] = baseClassField;
        });
        Arrays.stream(childClassFields).forEach(childClassField -> {
            fields[atomicInteger.getAndIncrement()] = childClassField;
        });
        return fields;
    }

}
