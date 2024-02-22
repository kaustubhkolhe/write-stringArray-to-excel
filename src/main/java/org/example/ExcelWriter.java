package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelWriter {

    public static void writeDataToExcel(String filePath, Map<String, List<String[]>> sheetData) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(filePath)) {
            int sheetIndex = 0;
            for (Map.Entry<String, List<String[]>> entry : sheetData.entrySet()) {
                String sheetName = entry.getKey();
                List<String[]> data = entry.getValue();
                Sheet sheet = workbook.createSheet(sheetName);

                int rowNum = 0;
                for (String[] rowData : data) {
                    Row row = sheet.createRow(rowNum++);
                    int colNum = 0;
                    for (String cellData : rowData) {
                        Cell cell = row.createCell(colNum++);
                        cell.setCellValue(cellData);
                    }
                }
                workbook.setSheetOrder(sheetName, sheetIndex++);
            }
            workbook.write(outputStream);
            System.out.println("Data has been written to " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error occurred while writing data to Excel file: " + e.getMessage());
        }
    }

    public static void writeDataToExcel(String filePath, List<String[]> sheetData) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(filePath)) {
            Sheet sheet = workbook.createSheet("Sheet1");

            int rowNum = 0;
            for (String[] rowData : sheetData) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(cellData);
                }
            }

            workbook.write(outputStream);
            System.out.println("Data has been written to " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error occurred while writing data to Excel file: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        Map<String, List<String[]>> sheetData = new LinkedHashMap<>();
        sheetData.put("English", List.of(
                new String[]{"client_id", "file_name", "sentence"},
                new String[]{"voice01", "englishVoice01", "Good morning! How did you sleep last night?"},
                new String[]{"voice02", "englishVoice02", "It's been such a long time since we last caught up. What have you been up to lately?"},
                new String[]{"voice03", "englishVoice03", "I hope your day is going well so far. Is there anything exciting happening for you today?"},
                new String[]{"voice04", "englishVoice04", "Nice to see you again! I've missed our chats. How have things been in your world?"},
                new String[]{"voice05", "englishVoice05", "What's the weather like today?"}
        ));
        sheetData.put("Hindi", List.of(
                new String[]{"client_id", "file_name", "sentence"},
                new String[]{"voice01", "hindiVoice01", "शुभ प्रभात! कल रात आपकी नींद केसी थी?"},
                new String[]{"voice02", "hindiVoice02", "हमारी आखिरी मुलाकात से बहुत समय बीत गया है। आप हाल ही में क्या कर रहे हैं?"},
                new String[]{"voice03", "hindiVoice03", "में आशा करता हूँ कि आपका दिन अब तक अच्छा बीत रहा होगा। क्या आज आपके लिए कुछ दिलचस्प हो रहा है?"},
                new String[]{"voice04", "hindiVoice04", "आपको फिर से देखकर खुशी हुई! मुझे हमारी बातचीतों की कमी महसूस हो रही थी। आपकी दुनिया में सब कुछ कैसा चल रहा है?"},
                new String[]{"voice05", "hindiVoice05", "आज का मौसम केसा है?"}
        ));
        sheetData.put("Gujarati", List.of(
                new String[]{"client_id", "file_name", "sentence"},
                new String[]{"voice01", "gujaratiVoice01", "શુભ સવાર! કેમ છો?"},
                new String[]{"voice02", "gujaratiVoice02", "આખરી મુલાકાત થી કેટલો સમય ગુજર્યો છે. તમે આજે શું કરી રહ્યાં છો?"},
                new String[]{"voice03", "gujaratiVoice03", "હું આશા કરું છું કે તમારો દિવસ અત્યંત સરસ ગયો છે. આજે તમારે કોઈ રસીક વિચારો છે?"},
                new String[]{"voice04", "gujaratiVoice04", "તમારી મુલાકાત પછી ફરીથી આનંદ થયો! મને તમારી સંવાદોની અભાવ અહેસાસ થયો. તમારી દુનિયામાં સર્વ કામ કેમ ચાલી રહ્યું છે?"},
                new String[]{"voice05", "gujaratiVoice05", "આજનું હવામાન કેમ છે?"}
        ));


        List<String[]> inputData = new ArrayList<>();


        inputData.add(new String[]{"client_id", "file_name", "sentence"});
        inputData.add(new String[]{"voice01", "englishVoice01", "Good morning! How did you sleep last night?"});
        inputData.add(new String[]{"voice02", "englishVoice02", "It's been such a long time since we last caught up. What have you been up to lately?"});
        inputData.add(new String[]{"voice03", "englishVoice03", "I hope your day is going well so far. Is there anything exciting happening for you today?"});
        inputData.add(new String[]{"voice04", "englishVoice04", "Nice to see you again! I've missed our chats. How have things been in your world?"});
        inputData.add(new String[]{"voice05", "englishVoice05", "What's the weather like today?"});

        String filePath = System.getProperty("user.home") + "\\JIO Internship\\output.xlsx";
        writeDataToExcel(filePath, inputData);
//        writeDataToExcel(filePath,sheetData);
    }
}
