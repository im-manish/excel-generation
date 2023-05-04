package com.excel.generator.excel_test;

import com.excel.generator.annotation.ExcelInfo;
import com.excel.generator.exporter.ExcelMetaData;
import com.excel.generator.exporter.ExcelWriter;
import com.excel.generator.util.TestUtils;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.SuperBuilder;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author 212743143
 */

public class ExcelGenerationWithInheritance {

    private static final ExcelMetaData<EmployeeInfo> EMPLOYEE_INFO_META_DATA = new ExcelMetaData<EmployeeInfo>(EmployeeInfo.class, true);

    @Test
    public void testInheritanceExcelExport() {

        ExcelWriter<EmployeeInfo> employeeInfoExcelWriter = ExcelWriter.<EmployeeInfo>builder()
                .excelMetaData(EMPLOYEE_INFO_META_DATA)
                //.fieldNames(EMPLOYEE_INFO_META_DATA.getFieldNames())
                //.headerNames(EMPLOYEE_INFO_META_DATA.getHeaderNames())
                //.headerNameVsColorCode(EMPLOYEE_INFO_META_DATA.getHeaderVsColor())
                .reportRecords(getEmployeeInfoRecords())
                //.headerVsColumnWidth(EMPLOYEE_INFO_META_DATA.getHeaderVsColumnWidth())
                .build();

        SXSSFWorkbook employeeDetails = employeeInfoExcelWriter.exportToExcel("Inherited Employee Details");


        TestUtils.writeWorkBook(employeeDetails, "C:\\Users\\6118454\\Desktop\\inherited_test.xlsx");
    }

    private List<EmployeeInfo> getEmployeeInfoRecords() {
        return IntStream.range(0,10).mapToObj(index ->
                EmployeeInfo.builder()
                        .employeeId(index)
                        .salary(100l+(float)index)
                        .joiningDate(new Date())
                        .firstName("Manish_"+index)
                        .lastName("Kumar_"+index)
                        .build()).collect(Collectors.toList());
    }

}


@Data
@SuperBuilder
@AllArgsConstructor
@NoArgsConstructor
class Person {
    @ExcelInfo(headerName = "First Name", sortIndex = 10.1, headerColor = HSSFColor.HSSFColorPredefined.BLUE_GREY,columnWidth = 20)
    private String firstName;
    @ExcelInfo(headerName = "Last Name", sortIndex = 10.2, headerColor = HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE, columnWidth = 20)
    private String lastName;
}

@SuperBuilder
@Data
@NoArgsConstructor
@AllArgsConstructor
class EmployeeInfo extends Person {
    @ExcelInfo(headerName = "Salary", sortIndex = 10.3, headerColor = HSSFColor.HSSFColorPredefined.RED, columnWidth = 20)
    private Float salary;

    @JsonFormat(shape = JsonFormat.Shape.STRING, pattern = "yyyy-MM-dd HH:mm:ss")
    @ExcelInfo(headerName = "Joining Date", sortIndex = 10.4, headerColor = HSSFColor.HSSFColorPredefined.SKY_BLUE,columnWidth = 40)
    private Date joiningDate;

    @ExcelInfo(headerName = "Employee Id", sortIndex = 9.5, headerColor = HSSFColor.HSSFColorPredefined.PALE_BLUE,columnWidth = 20)
    private Integer employeeId;
}