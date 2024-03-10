package example.office.word;

import example.office.Dict;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Word {

    private Dict dict;

    private List<Heading> headingList;

    public Word(Dict dict) {
        this.dict = dict;
    }


    /**
     * 替换doc、docx文件
     */
    public void processDoc(File doc, File targetFile) throws IOException {
        headingList = new ArrayList<>();
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(doc))) {
            // 遍历所有段落
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                String text = paragraph.getText();

                // Heading3
                addHeading(paragraph);

                // 替换文本
                String newParagraphText = dict.replaceAll(text);
                if (!newParagraphText.equals(text)) {
                    removeRuns(paragraph);

                    XWPFRun run = paragraph.createRun();
                    run.setText(newParagraphText);
                    paragraph.addRun(run);
                }

                // 替换段落
                dict.replaceParagraph(paragraph);

                // 遍历所有文本块
                /*for (XWPFRun r : paragraph.getRuns()) {
                    String runText = r.getText(0);
                    if(runText == null) {
                        continue;
                    }

                    // System.out.println(text);
                    // 检查并替换文本
                    String newText = dict.replaceAll(runText);
                    if(!newText.equals(runText)) {
                        r.setText(newText, 0);  // 设置新文本
                    }
                }*/
            }
            // 遍历所有表格
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables){
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows){
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells){
                        String text = cell.getText();
                        if(text == null) {
                            continue;
                        }

                        String newText = dict.replaceAll(text);
                        if(!newText.equals(text)) {
                            cell.setText(newText);  // 设置新文本
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                document.write(fos);
            }
        }
        System.out.println(headingList);
    }

    private void addHeading(XWPFParagraph paragraph) {
        // Title
        // Heading3
        // Caption
        // fs-4-first-line-indent-2
        String pStyle = paragraph.getCTPPr() != null && paragraph.getCTPPr().getPStyle() != null ? paragraph.getCTPPr().getPStyle().getVal() : null;
        if(pStyle != null) {
            System.out.println(pStyle);
            Heading heading = new Heading(pStyle, paragraph.getText());
            if(headingList.isEmpty()) {
                headingList.add(heading);
                return;
            }

            Heading prevHeading = headingList.get(headingList.size() - 1);
            if(prevHeading.getHeading().compareTo(heading.getHeading()) <= 0) {
                // 同级或更高级
                headingList.add(heading);
                return;
            }
            // 子级
            addChild(prevHeading, heading);
        }
    }

    private void addChild(Heading prevHeading, Heading heading) {
        if(prevHeading.getChildren() == null || prevHeading.getChildren().isEmpty()) {
            prevHeading.addChild(heading);
            return;
        }

        Heading lastChildHeading = prevHeading.lastChild();
        if(lastChildHeading.getHeading().compareTo(heading.getHeading()) <= 0) {
            // 同级或更高级
            prevHeading.addChild(heading);
            return;
        }

        // 子级
        addChild(lastChildHeading, heading);
    }

    private void removeRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = runs.size() - 1; i >= 0; i --) {
            paragraph.removeRun(i);
        }
    }
}
