package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static String orig_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/final.xlsx";
    public static String dest_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/ND.xlsx";

    public static void main(String[] args) {
        fillData();

    }

    public static void fillData() {
        try (FileInputStream fis = new FileInputStream(orig_xlsx);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(1);

            Iterator<Row> iterator = sheet.iterator();
            //dest
            FileInputStream fis1 = new FileInputStream(dest_xlsx);
            Workbook workbook1 = new XSSFWorkbook(fis1);
            Sheet sheet1 = workbook1.getSheetAt(0);
            Iterator<Row> iterator1 = sheet1.iterator();

            // Skip the header row
            if (iterator.hasNext()) {
                iterator.next();
            }
            if (iterator1.hasNext()) {
                iterator1.next();
            }
            int previous_score = 0;
            int previous2_score = 0;
            int previous3_score = 0;

            int previous_game_win = 0;
            int previous2_game_win = 0;
            int previous3_game_win = 0;

            int previous_set_win = 0;
            int previous2_set_win = 0;

            int current_score = 0;
            int current_game_win = 0;
            int current_set_win = 0;

            int point_sum = 0;
            int point_sum_previous = 0;
            int point_sum_previous2 = 0;
            int point_sum_previous3 = 0;

            int point_no = 1;
            int set_no = 1;
            int game_no = 1;
            int point_no_previous = 1;
            int set_no_previous = 1;
            int game_no_previous = 1;

            int server_count = 0;
            int ace_sum = 0;
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Row nextRow1 = iterator1.next();
                System.out.println("------------------\n Index " + nextRow.getRowNum());

                previous3_score = previous2_score;
                previous2_score = previous_score;
                previous_score = current_score;


                point_no_previous = point_no;
                set_no_previous = set_no;
                game_no_previous = game_no;
                point_sum_previous3 = point_sum_previous2;
                point_sum_previous2 = point_sum_previous;
                point_sum_previous = point_sum;


                set_no = (int) nextRow.getCell(4).getNumericCellValue();
                game_no = (int) nextRow.getCell(5).getNumericCellValue();
                point_no = (int) nextRow.getCell(6).getNumericCellValue();
                if ((int) nextRow.getCell(13).getNumericCellValue() == 2) {
                    server_count += 1;
                }
                if ((int) nextRow.getCell(21).getNumericCellValue() == 1) {
                    ace_sum += 1;
                }
                if (game_no > game_no_previous) {
                    previous3_game_win = previous2_game_win;
                    previous2_game_win = previous_game_win;
                    previous_game_win = current_game_win;
                }
                if (set_no > set_no_previous) {
                    previous2_set_win = previous_set_win;
                    previous_set_win = current_set_win;
                }

                // parse score sum
                Cell cell = nextRow.getCell(12);
                if (cell.getCellType() == CellType.NUMERIC) {
                    current_score = (int) nextRow.getCell(12).getNumericCellValue();
                    System.out.println(nextRow.getRowNum() + " current_score " + current_score);
                    if (current_score > previous_score) {
                        point_sum += 1;
                        nextRow1.getCell(13).setCellValue(point_sum);
                        System.out.println(nextRow1.getRowNum() + " New  " + point_sum);

                    } else {
                        nextRow1.getCell(13).setCellValue(point_sum);
                        System.out.println(nextRow1.getRowNum() + " New  " + point_sum);
                    }
                } else {
                    nextRow1.getCell(13).setCellValue(point_sum);
                    System.out.println(nextRow1.getRowNum() + " New  " + point_sum);
                }
                System.out.println(nextRow.getRowNum() + " point_sum " + point_sum + "\n"
                +" point_sum_previous " + point_sum_previous + "\n"
                +" point_sum_previous2 " + point_sum_previous2 + "\n"
                +" point_sum_previous3 " + point_sum_previous3 + "\n");

                //parse point_con_win3
                if (point_sum > point_sum_previous &&
                        point_sum_previous > point_sum_previous2 &&
                        point_sum_previous2 > point_sum_previous3) {
                    System.out.println(nextRow1.getRowNum() + " Point Con Win3  " );
                    nextRow1.getCell(18).setCellValue(1);
                }else{
                    nextRow1.getCell(18).setCellValue(0);
                }
                //parse point_con_fall3
                if (point_sum <= point_sum_previous &&
                        point_sum_previous <= point_sum_previous2 &&
                        point_sum_previous2 <= point_sum_previous3) {
                    if (point_sum > 0) {
                        System.out.println(nextRow1.getRowNum() + " Point Con Fall3  " );
                        nextRow1.getCell(19).setCellValue(1);
                    }else{
                        nextRow1.getCell(19).setCellValue(0);
                    }
                }else{
                    nextRow1.getCell(19).setCellValue(0);
                }
                System.out.println(nextRow.getRowNum() + " game_win " + current_game_win + "\n"
                        + " previous_game_win " + previous_game_win + "\n"
                        + " previous2_game_win " + previous2_game_win + "\n"
                        + " previous3_game_win " + previous3_game_win + "\n");
                //parse game_con win3
                current_game_win = (int)nextRow.getCell(10).getNumericCellValue();
                if (current_game_win > previous_game_win &&
                        previous_game_win > previous2_game_win &&
                        previous2_game_win > previous3_game_win) {
                    System.out.println(nextRow1.getRowNum() + " Game Con Win3  ");
                    nextRow1.getCell(20).setCellValue(1);
                }else {
                    nextRow1.getCell(20).setCellValue(0);
                }
                //parse game_con_fall3
                if (current_game_win <= previous_game_win &&
                        previous_game_win <= previous2_game_win &&
                        previous2_game_win <= previous3_game_win) {
                    if (current_game_win > 0) {
                        System.out.println(nextRow1.getRowNum() + " Game Con Fall3  ");
                        nextRow1.getCell(21).setCellValue(1);
                    }else{
                        nextRow1.getCell(21).setCellValue(0);
                    }

                }else {

                    nextRow1.getCell(21).setCellValue(0);
                }
                System.out.println(nextRow.getRowNum() + " set_win " + current_set_win + "\n"
                        + " previous_set_win " + previous_set_win + "\n"
                        + " previous2_set_win " + previous2_set_win + "\n");
                //parse set_con_win2
                current_set_win = (int)nextRow.getCell(8).getNumericCellValue();
                if (current_set_win > previous_set_win &&
                        previous_set_win > previous2_set_win) {
                    System.out.println(nextRow1.getRowNum() + " Set Con Win2  ");
                    nextRow1.getCell(22).setCellValue(1);
                }else {
                    nextRow1.getCell(22).setCellValue(0);
                }
                //parse set_con_fall2
                if (current_set_win <= previous_set_win &&
                        previous_set_win <= previous2_set_win) {
                    if (current_set_win > 0) {
                        System.out.println(nextRow1.getRowNum() + " Set Con Fall2  ");
                        nextRow1.getCell(23).setCellValue(1);
                    }else{
                        nextRow1.getCell(23).setCellValue(0);
                    }
                }else {
                    nextRow1.getCell(23).setCellValue(0);
                }

                double rate = (double)ace_sum / server_count;
                System.out.println(nextRow.getRowNum() + " ace_sum " + ace_sum + "\n"
                        + " server_count " + server_count + "\n"
                +" ace_sum/server_count " + rate + "\n");
                //parse ace/server_count
                nextRow1.getCell(24).setCellValue(rate);

            }
            try (FileOutputStream outputStream = new FileOutputStream(dest_xlsx)) {
                workbook1.write(outputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        // Now, the 'head' variable points to the head of the linked list of DayMarketState objects
    }
}