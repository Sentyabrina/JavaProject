package readcsv;
import javafx.stage.FileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import java.io.*;
import java.util.ArrayList;

/**
 * @author Katerina
 * @version 1.0
 *
 */
public class WriteDOCX {
    /** Класс для записи из шаблона в новый файл
     * @param csvFile  - файл с таблицей, содержащей информацию
     * @exception IOException - исключение, если файл не был создан
     * @exception InvalidFormatException - исключение, если не верный формат файла
     * */
    public void writeDoc(ArrayList csvFile) throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\template.docx");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));

        /** создание новой таблицы */
        XWPFTable table = doc.getTables().get(0);
        CTTbl ctTbl = CTTbl.Factory.newInstance();
        FileChooser fileChooser = new FileChooser();
        /** указание пути к файлу*/
        File userDirectory = new File("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\result.doc");
        fileChooser.setInitialDirectory(userDirectory);
        int i = 0;
        try {
            /** запись в таблицу*/
            for (i = 1; i < csvFile.size(); i++) {
                ArrayList<?> ardata = (ArrayList<?>) csvFile.get(i);
                XWPFTableRow row1 = table.createRow();
                for (int k = 0; k < ardata.size(); k++) {
                    XWPFTableCell cell = row1.createCell();
                    cell.setText(ardata.get(k).toString());
                }
                System.out.println();
            }
            System.out.println("Done");
            /** запись doc файла */
            OutputStream outputStream = new FileOutputStream(new File(userDirectory.getAbsolutePath()));
            doc.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}