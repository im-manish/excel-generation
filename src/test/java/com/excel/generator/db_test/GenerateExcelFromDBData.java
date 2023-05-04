package com.excel.generator.db_test;

import com.excel.generator.annotation.ColumnMapper;
import com.excel.generator.annotation.ExcelInfo;
import com.excel.generator.db.DatabaseConnectionPool;
import com.excel.generator.exporter.ExcelMetaData;
import com.excel.generator.exporter.ExcelWriter;
import com.excel.generator.util.ExcelGeneratorConstant;
import com.excel.generator.util.RowMapperUtils;
import com.excel.generator.util.TestUtils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.List;
import java.util.Properties;

/**
 * @author 6118454 - [Manish Kumar]
 */
class GenerateExcelFromDBData {

    private Properties properties;
    private static final ExcelMetaData<Order> EMPLOYEE_INFO_JQ_GRID_META_DATA = new ExcelMetaData<>(Order.class);

    @BeforeEach
    public void init() {
        properties = new Properties();
        properties.setProperty(ExcelGeneratorConstant.DRIVER, "");
        properties.setProperty(ExcelGeneratorConstant.URL, "");
        properties.setProperty(ExcelGeneratorConstant.USER_ID, "");
        properties.setProperty(ExcelGeneratorConstant.PASSWORD, "");
        properties.setProperty(ExcelGeneratorConstant.DRIVER, "");

    }

    @Test
    void testDBData() {

        try (Connection connection = DatabaseConnectionPool.getConnection(properties);
             PreparedStatement preparedStatement = connection.prepareStatement("Select * from ic_order")) {
           // preparedStatement.setLong(IPConstants.PARAMETER_INDEX_1, pubId);
            ResultSet resultSet = preparedStatement.executeQuery();
            RowMapperUtils<Order> rowMapperUtils = new RowMapperUtils<>(Order.class);
            List<Order> orders = rowMapperUtils.mapDataObjectToDTO(resultSet);

            ExcelWriter<Order> excelDemoExcelWriter = ExcelWriter.<Order>builder()
                    .reportRecords(orders)
                    .headerNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderNames())
                    .fieldNames(EMPLOYEE_INFO_JQ_GRID_META_DATA.getFieldNames())
                    .headerNameVsColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColor())
                    .headerNameVsFontColorCode(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsFontColor())
                    .headerVsColumnWidth(EMPLOYEE_INFO_JQ_GRID_META_DATA.getHeaderVsColumnWidth())
                    .build();
            SXSSFWorkbook employeeDetails = excelDemoExcelWriter.exportToExcel("Employee Details");
            TestUtils.writeWorkBook(employeeDetails, "C:\\Users\\6118454\\Desktop\\Single_sheet.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}


@AllArgsConstructor
@NoArgsConstructor
@Data
@Builder
class Order {
    @ColumnMapper(name = "ID")
    @ExcelInfo(headerName = "Employee Id", sortIndex = 10,headerColor = HSSFColor.HSSFColorPredefined.YELLOW)
    private Integer id;

    @ColumnMapper(name = "NAME")
    @ExcelInfo(headerName = "Employee Name", sortIndex = 20, headerColor = HSSFColor.HSSFColorPredefined.YELLOW)
    private String name;

    @ColumnMapper(name = "CITY")
    @ExcelInfo(headerName = "City", sortIndex = 30)
    private String city;
}