package Del;



import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import javax.imageio.ImageIO;

public class BagDataLoader {
    private static final String filePath = "./src";
    private static final String fileNm = "Bag.xlsx";

    private JTextArea textArea;
    private ImagePanel imagePanel;

    public BagDataLoader(JTextArea textArea, ImagePanel imagePanel) {
        this.textArea = textArea;
        this.imagePanel = imagePanel;
    }

    // 시트 인덱스를 추가로 받는 메서드
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

            // 이미지 데이터를 가져와서 ImagePanel에 표시
            loadImages(sheet);

        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Error reading file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void loadImages(XSSFSheet sheet) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();

        if (drawing != null) {
            List<XSSFShape> shapes = drawing.getShapes();
            for (XSSFShape shape : shapes) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    PictureData pictureData = picture.getPictureData();
                    byte[] data = pictureData.getData();

                    try {
                        // 이미지 데이터를 BufferedImage로 변환
                        BufferedImage img = ImageIO.read(new ByteArrayInputStream(data));
                        // BufferedImage를 ImagePanel에 설정
                        imagePanel.setImage(img);
                        break; // 첫 번째 이미지만 표시
                    } catch (IOException ex) {
                        ex.printStackTrace();
                    }
                }
            }
        }
    }
}
