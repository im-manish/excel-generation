package com.excel.generator.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Collectors;

@Slf4j
public class TestUtils {

    private final static ObjectMapper mapper = new ObjectMapper()
            .configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false)
            .configure(DeserializationFeature.READ_ENUMS_USING_TO_STRING, true)
            .configure(DeserializationFeature.ACCEPT_SINGLE_VALUE_AS_ARRAY,true)
            .configure(DeserializationFeature.ACCEPT_EMPTY_ARRAY_AS_NULL_OBJECT, true)
            ;


    public static <T> T getResource(String resourcePath, TypeReference<T> valueType) {
        try {
            Path path = Paths.get(TestUtils.class.getClassLoader().getResource(resourcePath).toURI());
            String txt = Files.lines(path).collect(Collectors.joining("\n"));
            return mapper.readValue( txt, valueType);
        } catch (IOException | URISyntaxException e) {
            e.printStackTrace();
        }

        return null;
    }

    public static <T> T getResource(String resourcePath, Class<T> valueType) {
        try {
            Path path = Paths.get(TestUtils.class.getClassLoader().getResource(resourcePath).toURI());
            String txt = Files.lines(path).collect(Collectors.joining("\n"));
            return mapper.readValue( txt, valueType);
        } catch (IOException | URISyntaxException e) {
            e.printStackTrace();
        }

        return null;
    }

    @SneakyThrows
    public static void writeWorkBook(SXSSFWorkbook hssfWorkbook, String path) {
        try (FileOutputStream out = new FileOutputStream(new File(path))) {
            hssfWorkbook.write(out);
        } catch (Exception e) {
            log.error("Error while generating the excel file",e);
        } finally {
            hssfWorkbook.close();
        }
    }
}
