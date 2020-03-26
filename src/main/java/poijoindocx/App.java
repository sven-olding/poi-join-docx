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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

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
        final XWPFDocument subDoc = new XWPFDocument(fisSubDoc);
        final List<IBodyElement> bodyElementsSub = subDoc.getBodyElements();

        for (int i = 0; i < 3; i++) {
          XmlCursor cursor = bookmarkParagraph.getCTP().newCursor();
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
              cursor = newTable.getCTTbl().newCursor();

            } else if (bodyElementSub.getElementType() == BodyElementType.PARAGRAPH) {
              XWPFParagraph p = (XWPFParagraph) bodyElementSub;

              XWPFParagraph newP;
              if (cursor.toNextSibling()) {
                newP = wordDoc.insertNewParagraph(cursor);
              } else {
                newP = wordDoc.createParagraph();
              }
              cloneParagraph(newP, p);
              cursor = newP.getCTP().newCursor();
            }
          }
        }
      }

      wordDoc.write(fosResult);
    } catch (final Exception e) {
      e.printStackTrace();
    }
  }

  private static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
    CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
    pPr.set(source.getCTP().getPPr());
    for (XWPFRun r : source.getRuns()) {
      XWPFRun nr = clone.createRun();
      cloneRun(nr, r);
    }
  }

  private static void cloneRun(XWPFRun clone, XWPFRun source) {
    CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
    rPr.set(source.getCTR().getRPr());
    clone.setText(source.getText(0));
    clone.setBold(source.isBold());
    clone.setItalic(source.isItalic());
    clone.setStrike(source.isStrike());
    clone.setFontFamily(source.getFontFamily() != null ? source.getFontFamily() : "Arial");
    clone.setFontSize(source.getFontSize() > -1 ? source.getFontSize() : 10);
    clone.setUnderline(source.getUnderline());
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