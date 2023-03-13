import entidades.Car;
import entidades.People;
import utils.excel.TExcelGenerate;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, IllegalAccessException {
        List<Car> cars = new ArrayList<>();
        cars.add(new Car("Chevrolet Onix", "Branco", "Chevrolet", new BigDecimal(60000)));
        cars.add(new Car("Fiat Strada", "Vermelho", "Fiat", new BigDecimal(75000)));


        List<People> people = new ArrayList<>();
        people.add(new People("Matheus Santos Eleodoro",26,19970102));
        people.add(new People("Joao Da Silva ",28,19970102));

        TExcelGenerate<Car> generator = new TExcelGenerate<>(cars,Car.class, TExcelGenerate.ExcelCompatibility.COMPATIBILITY_2007);
        generator.create("car.xlsx");

        TExcelGenerate<People> generator2 = new TExcelGenerate<>(people,People.class,TExcelGenerate.ExcelCompatibility.COMPATIBILITY_2003);
        generator2.create("people.xls");

    }
}
