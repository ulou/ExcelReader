package com.main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.*;

/**
 * Created by marsel on 08/10/2016.
 */
public class ReadExcel {
    public static void main(String args[]) {
        try {
            InputStream inp = new FileInputStream("assets/tabele-wartości-odżywczych.xls");
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheet("produkty rynkowe");

            PrintWriter writer = new PrintWriter("assets/produkty_rynkowe.sql", "UTF-8");

            for (Row row : sheet) {
                String productQuery = "INSERT INTO produkty_rynkowe VALUES(";
                String mineralQuery = "INSERT INTO skladniki_mineralne VALUES(";
                String proteinQuery = "INSERT INTO bialka VALUES(";
                String vitaminQuery = "INSERT INTO witaminy VALUES(";
                String saturatedFattyAcidsQuery = "INSERT INTO kwasy_tluszczowe_nasycone VALUES(";
                String monounsaturatedFattyAcidsQuery = "INSERT INTO kwasy_tluszczowe_jednonienasycone VALUES(";
                String polyunsaturatedFattyAcidsQuery = "INSERT INTO kwasy_tluszczowe_wielonienasycone VALUES(";
                String aminoAcids = "INSERT INTO aminokwasy VALUES(";
                String carbohydrates = "INSERT INTO weglodowany VALUES(";
                String energy = "INSERT INTO energia VALUES(";
                if (row.getRowNum() > 2 && row.getCell(0) != null) {
                    for (Cell cell : row) {
                        switch (CellReference.convertNumToColString(cell.getColumnIndex())){
                            case "C":
                                productQuery += cell.getStringCellValue() + ", ";
                                break;
                            case "A":
                                mineralQuery += (int) cell.getNumericCellValue() + ", ";
                                proteinQuery += (int) cell.getNumericCellValue() + ", ";
                                vitaminQuery += (int) cell.getNumericCellValue() + ", ";
                                saturatedFattyAcidsQuery += (int) cell.getNumericCellValue() + ", ";
                                monounsaturatedFattyAcidsQuery += (int) cell.getNumericCellValue() + ", ";
                                polyunsaturatedFattyAcidsQuery += (int) cell.getNumericCellValue() + ", ";
                                aminoAcids += (int) cell.getNumericCellValue() + ", ";
                                carbohydrates += (int) cell.getNumericCellValue() + ", ";
                                energy += (int) cell.getNumericCellValue() + ", ";
                                productQuery += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "D":
                            case "E":
                            case "F":
                                productQuery += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "I":
                                proteinQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "J":
                                proteinQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "G":
                            case "H":
                            case "K":
                            case "L":
                            case "M":
                            case "AT":
                            case "BB":
                            case "BK":
                                productQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "BL":
                                productQuery += (int) cell.getNumericCellValue() + ");";
                                break;
                            case "N":
                            case "O":
                            case "P":
                            case "Q":
                            case "R":
                                mineralQuery += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "S":
                            case "T":
                            case "U":
                                mineralQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "V":
                                mineralQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "W":
                            case "X":
                            case "Y":
                                vitaminQuery += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "Z":
                            case "AA":
                            case "AB":
                            case "AC":
                            case "AD":
                            case "AE":
                            case "AF":
                            case "AG":
                                vitaminQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "AH":
                                vitaminQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "AI":
                            case "AJ":
                            case "AK":
                            case "AL":
                            case "AM":
                            case "AN":
                            case "AO":
                            case "AP":
                            case "AQ":
                            case "AR":
                                saturatedFattyAcidsQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "AS":
                                saturatedFattyAcidsQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "AU":
                            case "AV":
                            case "AW":
                            case "AX":
                            case "AY":
                            case "AZ":
                                monounsaturatedFattyAcidsQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "BA":
                                monounsaturatedFattyAcidsQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "BC":
                            case "BD":
                            case "BE":
                            case "BF":
                            case "BG":
                            case "BH":
                            case "BI":
                                polyunsaturatedFattyAcidsQuery += cell.getNumericCellValue() + ", ";
                                break;
                            case "BJ":
                                polyunsaturatedFattyAcidsQuery += cell.getNumericCellValue() + ");";
                                break;
                            case "BM":
                            case "BN":
                            case "BO":
                            case "BP":
                            case "BQ":
                            case "BR":
                            case "BS":
                            case "BT":
                            case "BU":
                            case "BV":
                            case "BW":
                            case "BX":
                            case "BY":
                            case "BZ":
                            case "CA":
                            case "CB":
                            case "CC":
                                aminoAcids += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "CD":
                                aminoAcids += (int) cell.getNumericCellValue() + ");";
                                break;
                            case "CE":
                            case "CF":
                            case "CG":
                                carbohydrates += cell.getNumericCellValue() + ", ";
                                break;
                            case "CH":
                                carbohydrates += cell.getNumericCellValue() + ");";
                                break;
                            case "CI":
                            case "CJ":
                                energy += (int) cell.getNumericCellValue() + ", ";
                                break;
                            case "CK":
                                energy += (int) cell.getNumericCellValue() + ");";
                                break;
                        }
                    }
                    writer.println(productQuery);
                    writer.println(mineralQuery);
                    writer.println(vitaminQuery);
                    writer.println(proteinQuery);
                    writer.println(saturatedFattyAcidsQuery);
                    writer.println(monounsaturatedFattyAcidsQuery);
                    writer.println(polyunsaturatedFattyAcidsQuery);
                    writer.println(aminoAcids);
                    writer.println(carbohydrates);
                    writer.println(energy);
                }
                writer.println();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
