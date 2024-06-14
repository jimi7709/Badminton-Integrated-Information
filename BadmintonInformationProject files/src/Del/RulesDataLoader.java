package Del;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class RulesDataLoader {
    private static final String filePath = "./src";
    private static final String fileNm = "Rules.xlsx";

    private JTextArea textArea;

    public RulesDataLoader(JTextArea textArea) {
        this.textArea = textArea;
    }

    public void loadExcelData(int sheetIndex) {
        textArea.setText(""); // 이전 텍스트 지우기
        try (FileInputStream file = new FileInputStream(new File(filePath, fileNm))) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

            // 모든 행과 열에 대해서 데이터를 읽기
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell == null) {
                        textArea.append("NULL\t");
                        continue;
                    }

                    // cell의 타입을 확인하고, 값을 가져온다.
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            textArea.append((int) cell.getNumericCellValue() + "\t");
                            break;

                        case STRING:
                            textArea.append(cell.getStringCellValue() + "\t");
                            break;

                        default:
                            textArea.append("UNKNOWN\t");
                            break;
                    }
                }
                textArea.append("\n");
            }

        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Error reading file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}
