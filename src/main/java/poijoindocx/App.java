package poijoindocx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class App {

  public static void main(String[] args) {
    try (FileInputStream fisMainDoc = new FileInputStream(new File("./docx/Main.docx"));
        FileInputStream fisSubDoc = new FileInputStream(new File("./docx/Insert.docx"));
        FileOutputStream fileOutputStream = new FileOutputStream(new File("./docx/Result.docx"));) {
          XWPFDocument wordDoc = new XWPFDocument(fisMainDoc);
          
          wordDoc.write(fileOutputStream);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}