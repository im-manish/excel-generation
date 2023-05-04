package com.excel.generator.impl;

import com.excel.generator.api.IExcelGenerator;
import com.excel.generator.exporter.ExcelWriter;
import com.excel.generator.vo.ExcelGenerationVO;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Objects;

/**
 * Class that expose API to write the excel workbook
 *
 * @author Manish
 */

@Component
@Slf4j
public class ExcelGenerator implements IExcelGenerator {

    /**
     * Method to write the Workbook with multiple sheets info passed in Excel
     * generation vos
     *
     * @param excelGenerationVOS
     * @return
     */
    @Override
    public SXSSFWorkbook generateExcel(List<ExcelGenerationVO> excelGenerationVOS) {
        log.debug("Entered :: ExcelGenerator :: generateExcel :: excelGenerationVOS size: {}",
                CollectionUtils.size(excelGenerationVOS));
        SXSSFWorkbook workbook = null;
        if (CollectionUtils.isEmpty(excelGenerationVOS)) {
            return null;
        }
        for (ExcelGenerationVO excelGenerationVO : excelGenerationVOS) {
            if (Objects.isNull(workbook)) {
                workbook = ExcelWriter.builder().excelMetaData(excelGenerationVO.getMetaData())
                        .reportRecords(excelGenerationVO.getRecords()).build()
                        .exportToExcel(excelGenerationVO.getSheetName());
            } else {
                workbook = ExcelWriter.builder().excelMetaData(excelGenerationVO.getMetaData())
                        .reportRecords(excelGenerationVO.getRecords())
                        .workbook(workbook)
                        .build()
                        .exportToExcel(excelGenerationVO.getSheetName());
            }
        }
        log.debug("Exiting :: ExcelGenerator :: generateExcel :: number of sheets in workbook: {}",
                null != workbook ? workbook.getNumberOfSheets() : 0);
        return workbook;
    }
}
