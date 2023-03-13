package entidades;

import utils.excel.TExcelGenerate.*;

import java.math.BigDecimal;

@ExcelSheet(title = "Relatorio dos Carros", columnOrder = {"Nome","Valor"})
public class Car {
    @ExcelColumn(description = "Nome",width = 30)
    private String name;
    @ExcelColumn(description = "Cor", width = 50, alignment = AlignmentCell.CENTER)
    private String color;
    @ExcelColumn(description = "Marca", width = 50, alignment = AlignmentCell.LEFT)
    private String brand;
    @ExcelColumn(description = "Valor", width = 30, alignment = AlignmentCell.RIGHT,type = ColumnType.VALUE)
    private BigDecimal value;



    public Car(String name, String color, String brand, BigDecimal value) {
        this.name = name;
        this.color = color;
        this.brand = brand;
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getBrand() {
        return brand;
    }

    public void setBrand(String brand) {
        this.brand = brand;
    }

    public BigDecimal getValue() {
        return value;
    }

    public void setValue(BigDecimal value) {
        this.value = value;
    }
}
