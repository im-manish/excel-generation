package com.excel.generator.excel_test;

import com.excel.generator.exporter.ExcelMetaData;
import com.excel.generator.exporter.ExcelWriter;
import com.excel.generator.impl.ExcelGenerator;
import com.excel.generator.util.TestUtils;
import com.excel.generator.vo.ExcelGenerationVO;
import com.fasterxml.jackson.core.type.TypeReference;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * @author 6118454 [Manish Kumar]
 */

public class ExcelGeneratorWithoutExcelInfoTest {


    @Test
    public void testWithoutMemberInfoSheet() {

        Map<String,List<ExcelMetaData.MetaInfo>> excelInfos = TestUtils.getResource("test.json", new TypeReference<HashMap<String,List<ExcelMetaData.MetaInfo>>>() {});

        List<ExcelMetaData.MetaInfo> data = excelInfos.get(Department.class.getSimpleName());


        ExcelMetaData<Department> departmentExcelMetaData = new ExcelMetaData<Department>(data);

        ExcelWriter<Department> departmentInfoExcelWriter = ExcelWriter.<Department>builder()
                .reportRecords(getDepartmentRecord())
                .headerNames(departmentExcelMetaData.getHeaderNames())
                .fieldNames(departmentExcelMetaData.getFieldNames())
                .headerNameVsColorCode(departmentExcelMetaData.getHeaderVsColor())
                .headerVsColumnWidth(departmentExcelMetaData.getHeaderVsColumnWidth())
                .build();
        SXSSFWorkbook sxssfWorkbook = departmentInfoExcelWriter.exportToExcel("Without_MemberInfo");

        ExcelWriter<Department> departmentInfoExcelWriter1 = ExcelWriter.<Department>builder()
                .reportRecords(getDepartmentRecord())
                .headerNames(departmentExcelMetaData.getHeaderNames())
                .fieldNames(departmentExcelMetaData.getFieldNames())
                .headerNameVsColorCode(departmentExcelMetaData.getHeaderVsColor())
                .headerVsColumnWidth(departmentExcelMetaData.getHeaderVsColumnWidth())
                .workbook(sxssfWorkbook)
                .build();
        sxssfWorkbook = departmentInfoExcelWriter.exportToExcel("Without_MemberInfo_1");



        TestUtils.writeWorkBook(sxssfWorkbook, "C:\\Users\\6118454\\OneDrive - Thomson Reuters Incorporated\\Desktop\\Without_MemberInfo.xlsx");
    }


    @Test
    public void testMultiSheetInSingleCall() {
        Map<String,List<ExcelMetaData.MetaInfo>> excelInfos = TestUtils.getResource("json_file.json", new TypeReference<HashMap<String,List<ExcelMetaData.MetaInfo>>>() {});
        List<ExcelMetaData.MetaInfo> data = excelInfos.get(Department.class.getSimpleName());

        List<ExcelGenerationVO> excelGenerationVOS = getExcelGenerationVOS(excelInfos);

        ExcelGenerator excelGenerator = new ExcelGenerator();


        SXSSFWorkbook workbook = excelGenerator.generateExcel(excelGenerationVOS);


        TestUtils.writeWorkBook(workbook, "C:\\Users\\212743143\\Desktop\\Without_MemberInfo.xlsx");
    }

    private List<ExcelGenerationVO> getExcelGenerationVOS(Map<String, List<ExcelMetaData.MetaInfo>> excelInfos) {
        return Stream.of(ExcelGenerationVO
                        .builder()
                        .sheetName("Department1")
                        .metaData(new ExcelMetaData(excelInfos.get(Department.class.getSimpleName())))
                        .records(getDepartmentRecord())
                        .build(),
                ExcelGenerationVO
                        .builder()
                        .sheetName("Department2")
                        .metaData(new ExcelMetaData(excelInfos.get(Department.class.getSimpleName())))
                        .records(getDepartmentRecord())
                        .build()).collect(Collectors.toList());
    }

    private List<Department> getDepartmentRecord() {
        return IntStream.range(0, 1000)
                .mapToObj(x -> Department.builder().id(x).code("Code_"+x).name("Name_"+x).build()).collect(Collectors.toList());
    }


}

@AllArgsConstructor
@NoArgsConstructor
@lombok.Data
@Builder
class Department {

    private Integer id;

    private String name;

    private String code;

    private Person person;

}
