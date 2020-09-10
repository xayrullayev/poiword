import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ParagraphDecorator {
  private XWPFRun run;

  ParagraphDecorator(XWPFDocument document) {
    XWPFParagraph paragraph = document.createParagraph();
    this.run = paragraph.createRun();
  }

  ParagraphDecorator(XWPFRun run) {
    this.run = run;
  }

  public ParagraphDecorator setText(String text) {
    this.run.setText(text);
    return this;
  }

  public ParagraphDecorator setText(String text, Integer position) {
    this.run.setText(text, position);
    return this;
  }

  public ParagraphDecorator setBold(Boolean isBold) {
    this.run.setBold(isBold);
    return this;
  }

  public ParagraphDecorator setItalic(Boolean isBold) {
    this.run.setItalic(isBold);
    return this;
  }

  public ParagraphDecorator setFontFamily(String fontFamily) {
    this.run.setFontFamily(fontFamily);
    return this;
  }

  public ParagraphDecorator setFontSize(Integer size) {
    this.run.setFontSize(size);
    return this;
  }

  public ParagraphDecorator setVerticalAlignment(String alignment) {
    this.run.setVerticalAlignment(alignment);
    return this;
  }

  public ParagraphDecorator addBreak() {
    this.run.addBreak();
    return this;
  }

  public void setRun(XWPFRun run) {
    this.run = run;
  }

  public XWPFRun getRun() {
    return run;
  }

  public XWPFRun build() {
    return this.run;
  }

  @Override
  public String toString() {
    return "RunnerDecorator{" +
        "run=" + run +
        '}';
  }
}
