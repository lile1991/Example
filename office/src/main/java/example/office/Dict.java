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
        if(paragraphText.contains("{素材/素材1.docx:2.2.1}")) {
            System.out.println(paragraph.getText());
        }
    }
}
