package poijoindocx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

public class App {

  private static final String BOOKMARK_NAME = "Angebotspositionen";

  public static void main(final String[] args) {
    try (FileInputStream fisMainDoc = new FileInputStream(new File("./docx/Main.docx"));
        FileInputStream fisSubDoc = new FileInputStream(new File("./docx/Insert.docx"));
        FileOutputStream fosResult = new FileOutputStream(new File("./docx/Result.docx"));) {

      final XWPFDocument wordDoc = new XWPFDocument(fisMainDoc);
      final List<IBodyElement> bodyElements = wordDoc.getBodyElements();

      final XWPFParagraph bookmarkParagraph = findParagraphWithBookmark(bodyElements, BOOKMARK_NAME);
      if (bookmarkParagraph != null) {

        for (final XWPFRun run : bookmarkParagraph.getRuns()) {
          run.setText("", 0);
        }
        final XmlCursor cursor = bookmarkParagraph.getCTP().newCursor();
        final XWPFDocument subDoc = new XWPFDocument(fisSubDoc);
        final List<IBodyElement> bodyElementsSub = subDoc.getBodyElements();

        for (final IBodyElement bodyElementSub : bodyElementsSub) {
          if (bodyElementSub.getElementType() == BodyElementType.TABLE) {
            final XWPFTable x = (XWPFTable) bodyElementSub;
            XWPFTable newTable;
            if (cursor.toNextSibling()) {
              newTable = wordDoc.insertNewTbl(cursor);
            } else {
              newTable = wordDoc.createTable();
            }
            cloneTable(newTable, x);
          }
        }
      }

      wordDoc.write(fosResult);
    } catch (final Exception e) {
      e.printStackTrace();
    }
  }

  private static void cloneTable(final XWPFTable clone, final XWPFTable source) {
    for (final XWPFTableRow row : source.getRows()) {
      clone.addRow(row);
    }
  }

  private static XWPFParagraph findParagraphWithBookmark(final List<IBodyElement> elements, final String bookmarkName) {
    for (final IBodyElement element : elements) {
      final XWPFParagraph p = findParagraphWithBookmark(element, bookmarkName);
      if (p != null) {
        return p;
      }
    }
    return null;
  }

  private static XWPFParagraph findParagraphWithBookmark(final IBodyElement element, final String bookmarkName) {
    if (element.getElementType() == BodyElementType.PARAGRAPH) {
      final XWPFParagraph x = (XWPFParagraph) element;
      final List<CTBookmark> bookmarkStartList = x.getCTP().getBookmarkStartList();
      for (final CTBookmark b : bookmarkStartList) {
        if (b.getName().equalsIgnoreCase(bookmarkName)) {
          return x;
        }
      }
    } else if (element.getElementType() == BodyElementType.TABLE) {
      final XWPFTable x = (XWPFTable) element;
      for (final XWPFTableRow row : x.getRows()) {
        for (final XWPFTableCell cell : row.getTableCells()) {
          final XWPFParagraph p = findParagraphWithBookmark(cell.getBodyElements(), bookmarkName);
          if (p != null) {
            return p;
          }
        }
      }
    }
    return null;
  }
}