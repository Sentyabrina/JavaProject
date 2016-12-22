package readcsv;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.io.*;
/**
 *@author Katerina Perepechenko 535st group
 *@version 1.0
 *основной класс программы
 */

public class MainClass {
    @SuppressWarnings("deprecation")
    /** метод, содержащий вызов екземпляров всех клаассов
     * нужных для работы программы
     */
    public static void main(String args[]) throws IOException, InvalidFormatException {

    ReadCSV read = new ReadCSV();
    WriteDOCX write = new WriteDOCX();
    write.writeDoc(read.csv());
    System.out.println("Plese, open file result.doc");
    }
}