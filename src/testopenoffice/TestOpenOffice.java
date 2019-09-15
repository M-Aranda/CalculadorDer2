/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package testopenoffice;

import com.github.miachm.sods.Range;
import com.github.miachm.sods.Sheet;
import com.github.miachm.sods.SpreadSheet;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Chelo
 */
public class TestOpenOffice {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        ArrayList<Double> listaDeValores = new ArrayList<>();

        try {
            SpreadSheet spread = new SpreadSheet(new File("src/ps.ods"));
            //System.out.println("Number of sheets: " + spread.getNumSheets());

            List<Sheet> sheets = spread.getSheets();

            for (Sheet sheet : sheets) {
                //System.out.println("In sheet " + sheet.getName());

                Range range = sheet.getDataRange();

                int numerodeFilas = range.getNumRows();
                int numerodeColumnas = range.getNumColumns();

                ArrayList<Double> listaDeValoresxy = new ArrayList<>();
                ArrayList<Double> listaDeValoresxCuadrado = new ArrayList<>();
                ArrayList<Double> listaDeValoresyCuadrado = new ArrayList<>();

                Double sumaDeLasXs = 0.0;
                Double sumaDeLasYs = 0.0;
                Double sumaDeLasXYs = 0.0;
                Double sumaDeLasXsAlCuadrado = 0.0;
                Double sumaDeLasYAlCuadrado = 0.0;

                for (int i = 0; i < numerodeFilas; i++) {

                    Double valorxy = ((Double) range.getCell(i, 0).getValue()) * ((Double) range.getCell(i, 1).getValue());
                    Double valorxCuadrado = Math.pow((Double) range.getCell(i, 0).getValue(), 2);
                    Double valoryCuadrado = Math.pow((Double) range.getCell(i, 1).getValue(), 2);

                    listaDeValoresxy.add(valorxy);
                    listaDeValoresxCuadrado.add(valorxCuadrado);
                    listaDeValoresyCuadrado.add(valoryCuadrado);

                    sumaDeLasXs += (Double) range.getCell(i, 0).getValue();
                    sumaDeLasYs += (Double) range.getCell(i, 1).getValue();
                    sumaDeLasXYs += valorxy;
                    sumaDeLasXsAlCuadrado += valorxCuadrado;
                    sumaDeLasYAlCuadrado += valoryCuadrado;

                    for (int j = 0; j < numerodeColumnas; j++) {
                        listaDeValores.add((Double) range.getCell(i, j).getValue());

                    }

                }

                sheet.appendColumn();
                sheet.appendColumn();
                sheet.appendColumn();

                int numeroDeColumnasAhora = sheet.getMaxColumns();

                for (int i = 0; i < listaDeValoresxy.size(); i++) {
                    sheet.getDataRange().getCell(i, (numeroDeColumnasAhora - 3)).setValue(listaDeValoresxy.get(i));
                }

                for (int i = 0; i < listaDeValoresxCuadrado.size(); i++) {
                    sheet.getDataRange().getCell(i, (numeroDeColumnasAhora - 2)).setValue(listaDeValoresxCuadrado.get(i));
                }

                for (int i = 0; i < listaDeValoresyCuadrado.size(); i++) {
                    sheet.getDataRange().getCell(i, (numeroDeColumnasAhora - 1)).setValue(listaDeValoresyCuadrado.get(i));
                }

                Double promedioDeXs = sumaDeLasXs / numerodeFilas;
                Double promedioDeYs = sumaDeLasYs / numerodeFilas;

                Double b = ((numerodeFilas * sumaDeLasXYs) - (sumaDeLasXs * sumaDeLasYs)) / ((numerodeFilas * sumaDeLasXsAlCuadrado) - Math.pow(sumaDeLasXs, 2));
                Double a= promedioDeYs-(b*promedioDeXs);
                
                Double r2 = Math.pow((numerodeFilas * sumaDeLasXYs - (sumaDeLasXs * sumaDeLasYs)), 2)
                        / ((numerodeFilas * sumaDeLasXsAlCuadrado - Math.pow(sumaDeLasXs, 2)) * (numerodeFilas * sumaDeLasYAlCuadrado - Math.pow(sumaDeLasYs, 2)));

                sheet.insertRowBefore(0);
                sheet.getDataRange().getCell(0, 0).setValue("x");
                sheet.getDataRange().getCell(0, 1).setValue("y");
                sheet.getDataRange().getCell(0, 2).setValue("xy");
                sheet.getDataRange().getCell(0, 3).setValue("x al cuadrado");
                sheet.getDataRange().getCell(0, 4).setValue("y al cuadrado");

                System.out.println("Suma de las xs: " + sumaDeLasXs);
                System.out.println("Suma de las ys: " + sumaDeLasYs);
                System.out.println("Suma de las xys: " + sumaDeLasXYs);
                System.out.println("Suma de las xs al cuadrado: " + sumaDeLasXsAlCuadrado);
                System.out.println("Suma de las ys al cuadrado: " + sumaDeLasYAlCuadrado);
                System.out.println("Promedio de xs es: " + promedioDeXs);
                System.out.println("Promedio de ys es: " + promedioDeYs);
                System.out.println("a es: " + a);
                System.out.println("b es: " + b);
                System.out.println("r2 es: " + r2);

                sheet.setColumnWidth(0, 25.0);
                sheet.setColumnWidth(1, 25.0);

                spread.save(new File("src/ps.ods"));

            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
