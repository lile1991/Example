package example.office;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

@Setter
@Getter
@ToString
public class TokenConf {
    private Cell tokenCell;
    private Cell valueCell;
    private String token;
    private String value;

    public TokenConf(Cell tokenCell, Cell valueCell) {
        this.tokenCell = tokenCell;
        this.valueCell = valueCell;

        this.token = cellValue(tokenCell);
        this.value = cellValue(valueCell);
    }

    public static String cellValue(Cell cell) {
        // 根据单元格类型获取数据
        return cell.getCellType() == CellType.STRING ?
                cell.getStringCellValue() :
                String.valueOf(cell.getNumericCellValue());
    }

}
