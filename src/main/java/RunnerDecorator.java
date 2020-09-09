import org.apache.poi.xwpf.usermodel.XWPFRun;

public class RunnerDecorator {
  XWPFRun run;

  public RunnerDecorator() {
  }

  public RunnerDecorator(XWPFRun run) {
    this.run = run;
  }

  public RunnerDecorator setText(String text) {
    this.run.setText(text);
    return this;
  }

  public RunnerDecorator setText(String text, Integer position) {
    this.run.setText(text, position);
    return this;
  }

  public RunnerDecorator setBold(Boolean isBold) {
    this.run.setBold(isBold);
    return this;
  }

  public RunnerDecorator setItalic(Boolean isBold) {
    this.run.setItalic(isBold);
    return this;
  }

  public RunnerDecorator setFontFamily(String fontFamily) {
    this.run.setFontFamily(fontFamily);
    return this;
  }

  public RunnerDecorator setFontSize(Integer size) {
    this.run.setFontSize(size);
    return this;
  }

  public RunnerDecorator setVerticalAlignment(String alignment) {
    this.run.setVerticalAlignment(alignment);
    return this;
  }

  public RunnerDecorator addBreak() {
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
