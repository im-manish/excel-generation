package com.excel.generator.excel_test;

import com.excel.generator.annotation.ExcelInfo;
import com.excel.generator.exporter.ExcelMetaData;
import com.excel.generator.exporter.ExcelWriter;
import com.excel.generator.util.TestUtils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * @author 212743143
 */
public class ExcelGeneratorTest {

    private static final ExcelMetaData<Employee> EMPLOYEE_INFO_JQ_GRID_META_DATA = new ExcelMetaData<>(Employee.class);

    private static final ExcelMetaData<DepartmentInfo> DEPARTMENT_INFO_JQ_GRID_META_DATA = new ExcelMetaData<>(DepartmentInfo.class);


    @Test
    public void testSingleSheet() {
        ExcelWriter<Employee> excelDemoExcelWriter = ExcelWriter.<Employee>builder()
                .reportRecords(getEmployeeRecord())
                .headerNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderNames())
                .fieldNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getFieldNames())
                .headerNameVsColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColor())
                .headerNameVsFontColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsFontColor())
                .headerVsColumnWidth(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColumnWidth())
                .build();
        SXSSFWorkbook employeeDetails = excelDemoExcelWriter.exportToExcel("Employee Details");
        TestUtils.writeWorkBook(employeeDetails, "C:\\Users\\6118454\\OneDrive - Thomson Reuters Incorporated\\Documents\\Single_sheet.xlsx");
    }

    @Test
    public void testMultiSheet() {

        ExcelWriter<Employee> excelDemoExcelWriter =
                ExcelWriter.<Employee>builder()
                        .reportRecords(getEmployeeRecord())
                        .headerNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderNames())
                        .fieldNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getFieldNames())
                        .headerNameVsColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColor())
                        .headerNameVsFontColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsFontColor())
                        .headerVsColumnWidth(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColumnWidth())
                        .build();
        SXSSFWorkbook existingWorkbook = excelDemoExcelWriter.exportToExcel("Employee Details");
        ExcelWriter<DepartmentInfo> departmentInfoExcelWriter = ExcelWriter.<DepartmentInfo>builder()
                .reportRecords(getDepartmentRecord())
                .headerNames(DEPARTMENT_INFO_JQ_GRID_META_DATA.getHeaderNames())
                .fieldNames(DEPARTMENT_INFO_JQ_GRID_META_DATA.getFieldNames())
                .workbook(existingWorkbook)
                .headerNameVsColorCode(DEPARTMENT_INFO_JQ_GRID_META_DATA.getHeaderVsColor())
                .build();
        existingWorkbook = departmentInfoExcelWriter.exportToExcel("Department Info");
        TestUtils.writeWorkBook(existingWorkbook, "C:\\Users\\6118454\\OneDrive - Thomson Reuters Incorporated\\Documents\\Multi_sheet.xlsx");
    }

    private List<DepartmentInfo> getDepartmentRecord() {
        return Stream.of(DepartmentInfo.builder().id(1).code("DEPARTMENT_1").name("Department 1").build(),
                DepartmentInfo.builder().id(2).code("DEPARTMENT_2").name("Department 2").build()).collect(Collectors.toList());
    }


    private static List<Employee> getEmployeeRecord() {
        return IntStream.range(0, 1000)
                .mapToObj(x -> Employee.builder().id(x).city("Patna").mobile("9019024962").name("Manish").state("Bihar").pinCode("800001").build()).collect(Collectors.toList());

    }
}


@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
class Employee {

    @ExcelInfo(headerName = "Employee Id", sortIndex = 25,headerColor = HSSFColor.HSSFColorPredefined.AQUA,headerFontColor = HSSFColor.HSSFColorPredefined.BLUE_GREY)
    private Integer id;

    @ExcelInfo(headerName = "Employee Name", sortIndex = 35,columnWidth = 40, headerColor = HSSFColor.HSSFColorPredefined.RED,headerFontColor = HSSFColor.HSSFColorPredefined.GREEN)
    private String name;

    @ExcelInfo(headerName = "City", sortIndex = 30)
    private String city;

    @Override
    public String toString() {
        return "Employee{" +
                "pinCode='" + pinCode + '\'' +
                ", mobile='" + mobile + '\'' +
                ", address='" + address + '\'' +
                '}';
    }

    @ExcelInfo(headerName = "State", sortIndex = 40)
    private String state;

    @ExcelInfo(headerName = "Pin Code", sortIndex = 50)
    private String pinCode;

    @ExcelInfo(headerName = "Mobile", sortIndex = 60)
    private String mobile;

    @ExcelInfo(headerName = "Address", sortIndex = 45)
    private String address;


}


@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
class DepartmentInfo {

    @ExcelInfo(headerName = "Department Id", sortIndex = 10, headerColor = HSSFColor.HSSFColorPredefined.AQUA)
    private Integer id;

    @ExcelInfo(headerName = "Department Name", sortIndex = 11, headerColor = HSSFColor.HSSFColorPredefined.AQUA)
    private String name;

    @ExcelInfo(headerName = "Department Code", sortIndex = 14, headerColor = HSSFColor.HSSFColorPredefined.AQUA)
    private String code;

}
