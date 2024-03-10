package example.office;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Dict implements Closeable {
    Workbook workbook;

    private Map<String, TokenConf> cellMap = new HashMap<>();

    public Dict(File dictFile) throws IOException {
        // 读取 Excel 文件
        FileInputStream excelIn = new FileInputStream(dictFile);
        Workbook workbook = WorkbookFactory.create(excelIn);
        Sheet datatypeSheet = workbook.getSheetAt(0);

        // 读取 Excel 数据
        boolean isFirstRow = true;
        for (Row currentRow : datatypeSheet) {
            if(isFirstRow) {
                isFirstRow = false;
            }
            Cell tokenCell = currentRow.getCell(1);
            Cell valueCell = currentRow.getCell(2);

            TokenConf tokenConf = new TokenConf(tokenCell, valueCell);
            cellMap.put(tokenConf.getToken(), tokenConf);
        }
    }


    public TokenConf getTokenConfByToken(String token) {
        return cellMap.get(token);
    }

    @Override
    public void close() throws IOException {
        if(workbook != null) {
            workbook.close();
        }
    }

    public String contains(String text) {
        Set<String> tokenSet = cellMap.keySet();
        for(String token: tokenSet) {
            if(text.contains(token)) {
                return token;
            }
        }
        return null;
    }

    public String replaceAll(String text) {
        Collection<TokenConf> tokenConfSet = cellMap.values();
        for(TokenConf tokenConf: tokenConfSet) {
            if(text.contains(tokenConf.getToken())) {
                System.out.println("Replace: " + tokenConf.getToken());
                text = text.replaceAll(tokenConf.getToken().replace("{", "\\{"), tokenConf.getValue());
            }
        }
        return text;
    }

    /**
     * 替换段落
     * @param paragraph 段落, 不含序号
     */
    public void replaceParagraph(XWPFParagraph paragraph) {
        String text = paragraph.getText();
        String paragraphText = paragraph.getParagraphText();
        Pattern pattern = Pattern.compile("\\{(.*/?.*):(.*)}");
        Matcher matcher = pattern.matcher(text);
        // "{素材/素材1.docx:2.2.1}"
        while(matcher.find()) {
            // Heading3
            String pStyle = paragraph.getCTPPr() != null && paragraph.getCTPPr().getPStyle() != null ? paragraph.getCTPPr().getPStyle().getVal() : null;
            System.out.println("pStyle=" + pStyle);

            String token = matcher.group(0);
            String materialPath = matcher.group(1);
            String chapter = matcher.group(2);
            System.out.println("引用素材路径: " + materialPath + ", 章节: " + chapter);

            int level = paragraph.getCTP() != null && paragraph.getCTP().getPPr() != null
                    && paragraph.getCTP().getPPr().getOutlineLvl() != null ? paragraph.getCTP().getPPr().getOutlineLvl().getVal().intValue() : -1;
            if (level >= 0) {
                System.out.println("Heading Level :: " + (level+1) + " Text::" + paragraphText);
            }
        }
    }
}
