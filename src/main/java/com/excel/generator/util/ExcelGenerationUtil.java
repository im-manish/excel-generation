package com.excel.generator.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Utility class for Excel generation
 *
 * @author Manish
 */

@Slf4j
public class ExcelGenerationUtil {

    /** the object mapper */
    private final static ObjectMapper mapper = new ObjectMapper()
            .configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false)
            .configure(DeserializationFeature.READ_ENUMS_USING_TO_STRING, true)
            .configure(DeserializationFeature.ACCEPT_SINGLE_VALUE_AS_ARRAY, true)
            .configure(DeserializationFeature.ACCEPT_EMPTY_ARRAY_AS_NULL_OBJECT, true);

    /**
     * method to divide the list in given size
     *
     * @param list
     * @param L
     * @param <T>
     * @return
     */
    public static <T> List<List<T>> getChoppedList(List<T> list, final int L) {
        log.debug("Entering ExcelGenerationUtil :: getChoppedList :: list size: {} :: list division factor: {}",
                CollectionUtils.size(list), L);
        List<List<T>> parts = new ArrayList<List<T>>();
        final int N = list.size();
        for (int i = 0; i < N; i += L) {
            parts.add(new ArrayList<T>(list.subList(i, Math.min(N, i + L))));
        }
        log.debug("Exiting ExcelGenerationUtil :: getChoppedList :: chopped list size: {}",
                CollectionUtils.size(parts));
        return parts;
    }

    /**
     * method to read the resource json file and convert it to passed class object
     *
     * @param resourcePath
     * @param valueType
     * @param <T>
     * @return
     */
    public static <T> T getResourceDataJson(String resourcePath, TypeReference<T> valueType) {
        log.debug("Entering ExcelGenerationUtil :: getResourceDataJson :: resourcePath: {}", resourcePath);
        try {
            Path path = Paths.get(resourcePath);
            String txt = Files.lines(path).collect(Collectors.joining("\n"));
            return mapper.readValue(txt, valueType);
        } catch (Exception e) {
            log.error("Error while reading the resource file:: resourcePath : {}, exception: {}", resourcePath, e);
            throw new RuntimeException("Error while reading the resource file in ExcelGenerationUtil");
        }
    }

    /**
     * method to write a excel file to given path
     *
     * @param hssfWorkbook
     * @param path
     */
    @SneakyThrows
    public static void writeWorkBook(SXSSFWorkbook hssfWorkbook, String path) {
        log.debug("Entering ExcelGenerationUtil :: writeWorkBook :: path: {}", path);
        try (FileOutputStream out = new FileOutputStream(new File(path))) {
            hssfWorkbook.write(out);
        } catch (Exception e) {
            log.error("Error while generating the excel file", e);
        } finally {
            hssfWorkbook.close();
        }
        log.debug("Exiting ExcelGenerationUtil :: writeWorkBook :: path: {}", path);
    }

}