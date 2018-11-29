package com.github.syl20lego.excel2json;

import com.fasterxml.jackson.core.util.DefaultPrettyPrinter;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.ObjectWriter;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Executor {

    public static void main(String[] args) throws InvalidFormatException {
        try {
            FileDialog dialog = new FileDialog((Frame) null, "Select File to Open");
            dialog.setMode(FileDialog.LOAD);
            dialog.setFilenameFilter((dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls"));
            dialog.setVisible(true);
            String file = dialog.getFile();
            if (file != null) {
                String directory = dialog.getDirectory();
                String jsonFile = file.substring(0, file.lastIndexOf(".")) + ".json";

                Workbook workbook = WorkbookFactory.create(new File(directory + file), null, true);

                Map<String, List> sheets = new LinkedHashMap<>();
                workbook.forEach(sheet -> {
                    List<Map<String, Object>> list = new ArrayList<>();
                    sheets.put(sheet.getSheetName(), list);
                    List<String> header = new ArrayList<>();
                    sheet.forEach(row -> {
                        Map<String, Object> outputRow = new LinkedHashMap<>();
                        boolean headerRow = header.isEmpty();
                        row.forEach(cell -> {
                            if (headerRow) {
                                header.add(cell.toString());
                            } else {
                                if (cell.toString().length() > 0) {
                                    outputRow.put(header.get(cell.getColumnIndex() - 1), cellValue(cell));
                                }
                            }
                        });
                        if (!outputRow.isEmpty()) {
                            list.add(outputRow);
                        }
                    });

                });


                ObjectMapper mapper = new ObjectMapper();
                mapper.enable(SerializationFeature.INDENT_OUTPUT);
                ObjectWriter writer = mapper.writer(new DefaultPrettyPrinter());
                writer.writeValue(new File(directory + jsonFile), sheets);

                // Closing the workbook
                workbook.close();
            }

            System.out.println("Done");
        } catch (IOException e) {
            System.out.println("Expecting file xlsx");
        }
    }

    private static Object cellValue(Cell cell) {
        DataFormatter dataFormatter = new DataFormatter();

        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case STRING:
                return cell.getRichStringCellValue().getString();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return dataFormatter.formatCellValue(cell);
                } else {
                    if ((cell.getNumericCellValue() % 1) == 0) {
                        return Double.valueOf(cell.getNumericCellValue()).intValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                }
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
