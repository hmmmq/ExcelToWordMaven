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
import org.docx4j.wml.Text;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class FilePath {
    // Path to the file
    public static final String EXCEL_FILE_PATH = "src/main/resources/excel.xlsx";
    public static final String DOC_FILE_PATH = "src/main/resources/application.docx";

    public static final String ICON_FILE_PATH_PEEFIX = "src/main/resources/icon/";
    public static final String ICON_FILE_PATH_SUFFIX = ".png";

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
    public HashMap<String, String> readFromExcel(int row_id) {
        HashMap<String, String> mappings = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(FilePath.EXCEL_FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 假设数据在第一个sheet
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(row_id); // 读取第一行数据

            // 获取每一列的数据
            int no = (int) row.getCell(0).getNumericCellValue();
            mappings.put(FieldName.NO, String.valueOf(no));
            String name = row.getCell(1).getStringCellValue();
            mappings.put(FieldName.NAME, name);
            String gender = row.getCell(2).getStringCellValue();
            mappings.put(FieldName.GENDER,gender);
            String highestEducation = row.getCell(3).getStringCellValue();
            mappings.put(FieldName.HIGHEST_EDUCATION,highestEducation);
            String nationality = row.getCell(4).getStringCellValue();
            mappings.put(FieldName.NATIONALITY,nationality);
            String cardType = row.getCell(5).getStringCellValue();
            mappings.put(FieldName.CARD_TYPE,cardType);
            String idCardNumber = row.getCell(6).getStringCellValue();
            mappings.put(FieldName.ID_CARD_NUMBER,idCardNumber);
            String dateOfBirth = row.getCell(7).getStringCellValue();
            mappings.put(FieldName.DATE_OF_BIRTH,dateOfBirth);
            String personalHealthCommitmentLetter = row.getCell(8).getStringCellValue();
            mappings.put(FieldName.PERSONAL_HEALTH_COMMITMENT_LETTER,personalHealthCommitmentLetter);
            String physicalCondition = row.getCell(9).getStringCellValue();
            mappings.put(FieldName.PHYSICAL_CONDITION,physicalCondition);
            String jobPosition = row.getCell(10).getStringCellValue();
            mappings.put(FieldName.JOB_POSITION,jobPosition);
            String applicationProject = row.getCell(11).getStringCellValue();
            mappings.put(FieldName.APPLICATION_PROJECT,applicationProject);
            String employmentCategory = row.getCell(12).getStringCellValue();
            mappings.put(FieldName.EMPLOYMENT_CATEGORY,employmentCategory);
            String trainingType = row.getCell(13).getStringCellValue();
            mappings.put(FieldName.TRAINING_TYPE,trainingType);
            mappings.put(FieldName.TELEPHONE,String.format("%.0f", row.getCell(14).getNumericCellValue()));
            String workUnit = row.getCell(15).getStringCellValue();
            mappings.put(FieldName.WORK_UNIT,workUnit);
            String correspondenceAddress = row.getCell(16).getStringCellValue();
            mappings.put(FieldName.CORRESPONDENCE_ADDRESS,correspondenceAddress);

        } catch (IOException e) {
            e.printStackTrace();
        }
        return mappings;
    }
}
class WriteIO {
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
                //插入文本
                for (Map.Entry<String, String> entry : mappings.entrySet()) {
                    if (text.contains(entry.getKey())) {
                        textElement.setValue(text.replace(entry.getKey(), entry.getValue()));
                    }
                }
            }
            // 保存修改后的文档
            wordMLPackage.save(new File(FilePath.FILLED_DOC_FILE_PATH_PREFIX + mappings.get(FieldName.NAME) + FilePath.FILLED_DOC_FILE_PATH_SUFFIX));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public void insertImageToDoc(String name){
        try{
        String docPath = FilePath.FILLED_DOC_FILE_PATH_PREFIX + name + FilePath.FILLED_DOC_FILE_PATH_SUFFIX;
        String imagePath1 = FilePath.ICON_FILE_PATH_PEEFIX + name + "1" + FilePath.ICON_FILE_PATH_SUFFIX;
        String imagePath2 = FilePath.ICON_FILE_PATH_PEEFIX + name + "2" + FilePath.ICON_FILE_PATH_SUFFIX;
        String imagePath3 = FilePath.ICON_FILE_PATH_PEEFIX + name + "3" + FilePath.ICON_FILE_PATH_SUFFIX;
        String imagePath4 = FilePath.ICON_FILE_PATH_PEEFIX + name + "4" + FilePath.ICON_FILE_PATH_SUFFIX;
        insertImageToDoc(docPath,imagePath1,FieldName.ICON_A,400,300);
        insertImageToDoc(docPath,imagePath2,FieldName.ICON_B,400,300);
        insertImageToDoc(docPath,imagePath3,FieldName.ICON_C,150,200);
        insertImageToDoc(docPath,imagePath4,FieldName.CERTIFICATE,400,300);}
        catch (Exception e){
            System.out.println("没有足够的图片");
            e.printStackTrace();
        }
    }
    private void insertImageToDoc(String docPath,String imagePath,String tablePlaceholder,int width, int height){
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
    public static void main(String[] args)  {
        for (int i = 2; i < 12; i++){
            System.out.println("Reading row " + i);
            ReadIO readIO = new ReadIO();
            HashMap<String, String> mappings = readIO.readFromExcel(i);
            WriteIO writeIO = new WriteIO();
            writeIO.writeToDoc(mappings);
            writeIO.insertImageToDoc(mappings.get(FieldName.NAME));
        }
    }
}
