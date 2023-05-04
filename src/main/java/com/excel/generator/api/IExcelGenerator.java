package com.excel.generator.api;

import com.excel.generator.vo.ExcelGenerationVO;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

/**
 * Excel generator interface
 * @author Manish
 */
public interface IExcelGenerator {

    SXSSFWorkbook generateExcel(List<ExcelGenerationVO> excelGenerationVOS);
}
