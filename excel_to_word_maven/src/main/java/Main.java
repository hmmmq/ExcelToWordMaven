import jakarta.xml.bind.JAXBElement;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

class FilePath {
    // Path to the file
    public static final String EXCEL_FILE_PATH = "src/main/resources/excel2.xlsx";
    public static final String DOC_FILE_PATH = "src/main/resources/application2.docx";

    public static final String ICON_FILE_PATH_PEEFIX = "src/main/resources/icon/";
    public static final String ICON_FILE_PATH_SUFFIX_PNG = ".png";
    public static final String ICON_FILE_PATH_SUFFIX_JPG = ".jpg";

    public static final String FILLED_DOC_FILE_PATH_PREFIX = "src/main/resources/";
    public static final String FILLED_DOC_FILE_PATH_SUFFIX = ".docx";
}

class FieldName {
    // Field names in the Excel file
    public static final String NO = "NO";
    public static final String NAME = "NAME";
    public static final String GENDER = "GENDER";
    public static final String HIGHEST_EDUCATION = "HIGHEST_EDUCATION";
    public static final String NATIONALITY = "NATIONALITY";
    public static final String CARD_TYPE = "CARD_TYPE";
    public static final String ID_CARD_NUMBER = "ID_CARD_NUMBER";
    public static final String DATE_OF_BIRTH = "DATE_OF_BIRTH";
    public static final String PERSONAL_HEALTH_COMMITMENT_LETTER = "PERSONAL_HEALTH_COMMITMENT_LETTER";
    public static final String PHYSICAL_CONDITION = "PHYSICAL_CONDITION";
    public static final String JOB_POSITION = "JOB_POSITION";
    public static final String APPLICATION_PROJECT = "APPLICATION_PROJECT";
    public static final String EMPLOYMENT_CATEGORY = "EMPLOYMENT_CATEGORY";
    public static final String TRAINING_TYPE = "TRAINING_TYPE";
    public static final String TELEPHONE = "TELEPHONE";
    public static final String WORK_UNIT = "WORK_UNIT";
    public static final String CORRESPONDENCE_ADDRESS = "CORRESPONDENCE_ADDRESS";
    public static final String ICON_A = "ICON_A";
    public static final String ICON_B = "ICON_B";
    public static final String ICON_C = "ICON_C";
    public static final String CERTIFICATE = "CERTIFICATE";

}

class ReadIO {
    HashMap<Integer, String> dictionary = new HashMap<>();


    {
        dictionary.put(0, FieldName.NO);
        dictionary.put(1, FieldName.NAME);
        dictionary.put(2, FieldName.GENDER);
        dictionary.put(3, FieldName.HIGHEST_EDUCATION);
        dictionary.put(4, FieldName.NATIONALITY);
        dictionary.put(5, FieldName.CARD_TYPE);
        dictionary.put(6, FieldName.ID_CARD_NUMBER);
        dictionary.put(7, FieldName.DATE_OF_BIRTH);
        dictionary.put(8, FieldName.PERSONAL_HEALTH_COMMITMENT_LETTER);
        dictionary.put(9, FieldName.PHYSICAL_CONDITION);
        dictionary.put(10, FieldName.JOB_POSITION);
        dictionary.put(11, FieldName.APPLICATION_PROJECT);
        dictionary.put(12, FieldName.EMPLOYMENT_CATEGORY);
        dictionary.put(13, FieldName.TRAINING_TYPE);
        dictionary.put(14, FieldName.TELEPHONE);
        dictionary.put(15, FieldName.WORK_UNIT);
//        dictionary.put(16, FieldName.CORRESPONDENCE_ADDRESS);
    }
    public HashMap<Integer, String> getDictionary() {
        return dictionary;
    }

    public void setDictionary(HashMap<Integer, String> dictionary) {
        this.dictionary = dictionary;
    }

    public HashMap<String, String> readFromExcel(int row_id) {
        HashMap<String, String> mappings = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(FilePath.EXCEL_FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 假设数据在第一个sheet
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(row_id); // 读取第一行数据

            // 获取每一列的数据
            for (int i = 0; i < dictionary.size(); i++) {
                Cell cell = row.getCell(i);
                String cellValue = "";
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            cellValue = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            cellValue = String.format("%.0f", cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            cellValue = cell.getCellFormula();
                            break;
                        default:
                            cellValue = "";
                    }
                }
                mappings.put(dictionary.get(i), cellValue);
            }


        } catch (IOException e) {
            e.printStackTrace();
        }
        return mappings;
    }
}

class WriteIO {
    HashMap<Integer, String> imageDictionary = new HashMap<>();
    int[] imageWidth = {400, 400, 400, 150};
    int[] imageHeight = {300, 300, 300, 200};

    {
        imageDictionary.put(1, FieldName.ICON_A);
        imageDictionary.put(2, FieldName.ICON_B);
        imageDictionary.put(3, FieldName.CERTIFICATE);
        imageDictionary.put(4, FieldName.ICON_C);
    }

    public void writeToDoc(HashMap<String, String> mappings) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(FilePath.DOC_FILE_PATH));
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            // 遍历文档的所有文本节点，进行替换
            List<Object> texts = documentPart.getJAXBNodesViaXPath("//w:t", true);
            for (Object obj : texts) {
                JAXBElement jaxbElement = (JAXBElement) obj;
                Text textElement = (Text) jaxbElement.getValue();
                String text = textElement.getValue();
                if (mappings.containsKey(text)) {
                    textElement.setValue(mappings.get(text));
                }
            }
            // 保存修改后的文档
            wordMLPackage.save(new File(FilePath.FILLED_DOC_FILE_PATH_PREFIX + mappings.get(FieldName.ID_CARD_NUMBER) + "_" + mappings.get(FieldName.NAME) + FilePath.FILLED_DOC_FILE_PATH_SUFFIX));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void insertImageToDoc(String name) {

        try {
            String docPath = FilePath.FILLED_DOC_FILE_PATH_PREFIX + name + FilePath.FILLED_DOC_FILE_PATH_SUFFIX;
            for (int i = 1; i < 5; i++) {
                String imagePath = FilePath.ICON_FILE_PATH_PEEFIX + name + i + FilePath.ICON_FILE_PATH_SUFFIX_PNG;
                Path path = Paths.get(imagePath);
                if (!Files.exists(path)) {
                    imagePath = FilePath.ICON_FILE_PATH_PEEFIX + name + i + FilePath.ICON_FILE_PATH_SUFFIX_JPG;
                }
                insertImageToDoc(docPath, imagePath, imageDictionary.get(i), imageWidth[i - 1], imageHeight[i - 1]);
            }
        } catch (Exception e) {
            System.out.println("没有足够的图片");
            e.printStackTrace();
        }
    }

    public void insertImageToDoc(String IdCardNumber, String name) {
        try {
            String docPath = FilePath.FILLED_DOC_FILE_PATH_PREFIX + IdCardNumber + "_" + name + FilePath.FILLED_DOC_FILE_PATH_SUFFIX;
            for (int i = 1; i < 5; i++) {
                String imagePath = FilePath.ICON_FILE_PATH_PEEFIX + name + i + FilePath.ICON_FILE_PATH_SUFFIX_PNG;
                Path path = Paths.get(imagePath);
                if (Files.exists(path)) {
                    insertImageToDoc(docPath, imagePath, imageDictionary.get(i), imageWidth[i - 1], imageHeight[i - 1]);
                } else {
                    imagePath = FilePath.ICON_FILE_PATH_PEEFIX + name + i + FilePath.ICON_FILE_PATH_SUFFIX_JPG;
                    path = Paths.get(imagePath);
                    if (Files.exists(path)) {
                        insertImageToDoc(docPath, imagePath, imageDictionary.get(i), imageWidth[i - 1], imageHeight[i - 1]);
                    } else {
                        System.out.println("没有足够的图片(.jpg 或 .png):" + name + i);
                    }
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void insertImageToDoc(String docPath, String imagePath, String tablePlaceholder, int width, int height) {
        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // 寻找目标表格和占位符
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        if (cell.getText().contains(tablePlaceholder)) {
                            // 移除占位符文本
                            cell.removeParagraph(0);
                            // 插入图片
                            try (InputStream is = new FileInputStream(imagePath)) {
                                XWPFParagraph paragraph = cell.addParagraph();
                                XWPFRun run = paragraph.createRun();
                                run.addPicture(is, XWPFDocument.PICTURE_TYPE_PNG, imagePath,
                                        width * 9525, height * 9525); // 图片尺寸 100x100 pt
                            }
                            // 替换一次后退出
                            break;
                        }
                    }
                }
            }

            // 保存修改后的文档
            try (FileOutputStream fos = new FileOutputStream(docPath)) {
                doc.write(fos);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

public class Main {
    public static void main(String[] args) {
        //让用户输入行数
         Scanner scanner = new Scanner(System.in);
         System.out.println("请输入excel最后一行的序号：");
         int last_row = scanner.nextInt();
         //询问打印已知字段跟用户核对
            System.out.println("请核对以下字段是否正确：");
            ReadIO excel_dictionary_readio = new ReadIO();
            HashMap<Integer, String> excel_dictionary = excel_dictionary_readio.getDictionary();
            for (Map.Entry<Integer, String> entry : excel_dictionary.entrySet()) {
                System.out.println("第" + entry.getKey() + "列数据对应word文档里面的占位符：" + entry.getValue());
            }
            //用户核对完之后要用户打Y或者N,如果打Y则继续，如果打N则退出
            System.out.println("请核对完毕后输入Y或y继续，输入N或n退出：");
            String user_input = scanner.next();
            if (user_input.equals("N")||user_input.equals("n")) {
                System.out.println("如果字段不匹配,请检查ReadIO类里面的HashMap<Integer, String> dictionary的键值对是否正确");
                System.out.println("若要增加字段,需要在FieldName类里面增加字段名,并且在ReadIO类里面增加对应的键值对,并且在document里面增加对应的占位符");
                System.out.println("若要减少字段,需要在ReadIO类里面删除对应的键值对,并且在document里面删除对应的占位符");
                System.out.println("若要在document里面不显示某个字段,在document里面删除对应的占位符即可");
                System.out.println("若要修改excel字段的位置,比如姓名改成第5列数据了,需要在ReadIO类里面修改对应的键值对,将键值对的键改成4,(从0开始数列值)");
                System.out.println("若要修改doc文档里面占位符的位置,直接doc文档里面修改占位符的位置即可");
                System.out.println("程序退出");
                System.exit(0);
            }

        for (int i = 0; i < last_row; i++) {
            System.out.println("Reading row " + i);
            ReadIO readIO = new ReadIO();
            HashMap<String, String> mappings = readIO.readFromExcel(i);
            if (mappings.get(FieldName.NAME)==""||mappings.get(FieldName.ID_CARD_NUMBER)=="") {
                continue;
            }
            WriteIO writeIO = new WriteIO();
            writeIO.writeToDoc(mappings);
            writeIO.insertImageToDoc(mappings.get(FieldName.ID_CARD_NUMBER), mappings.get(FieldName.NAME));
        }
    }
}
