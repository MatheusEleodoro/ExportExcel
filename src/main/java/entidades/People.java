package entidades;

import utils.excel.TExcelGenerate.*;

@ExcelSheet(title = "Relatorio de Pessoas")
public class People {

    public People(String name, Integer age, Integer birth) {
        this.name = name;
        this.age = age;
        this.birth = birth;
    }

    @ExcelColumn(description = "Nome",width = 30,alignment = AlignmentCell.LEFT)
    private String name;

    @ExcelColumn(description = "Idade",width = 15,alignment = AlignmentCell.LEFT)
    private Integer age;

    @ExcelColumn(description = "Data de Nascimento", width = 30,alignment = AlignmentCell.CENTER, type = ColumnType.DATE,outputPattern = "MM/yyyy")
    private Integer birth;
}
