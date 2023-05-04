package com.excel.generator.vo;

import com.excel.generator.exporter.ExcelMetaData;
import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * The VO class holds the information required for generation of Excel
 *
 * @author Manish
 */

@Data
@Builder
public class ExcelGenerationVO<T> {

    /** the sheet name */
    private String sheetName;

    /** the metaData */
    private ExcelMetaData metaData;

    /** the list of records */
    private List records;
}
