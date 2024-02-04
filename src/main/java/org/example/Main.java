package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static String orig_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/full_data.xlsx";
    public static String dest_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/P1.xlsx";
    public static int orig_xlsx_sheet = 0;
    public static int dest_xlsx_sheet = 0;
    public static int set_no_column = 4;
    public static int game_no_column = 5;
    public static int point_no_column = 6;
    public static int server_column = 13;
    public static int ace_column = 20;
    public static int win_column = 22;
    public static int point_score_column = 11;
    public static int point_sum_column = 20;
    public static int point_con_win3_column = 21;
    public static int point_con_fall3_column = 22;
    public static int current_game_win_column = 9;
    public static int game_con_win3_column = 23;
    public static int game_con_fall3_column = 24;
    public static int current_set_win_column = 7;
    public static int set_con_win2_column = 25;
    public static int set_con_fall2_column = 26;
    public static int ace_sum_d_server_count_column = 27;
    public static int player_name_column = 1;
    public static int speed_column = 42;

    public static int game_direct_serve_success_count_column = 29;
    public static int game_hit_success_count_column = 30;
    public static int game_con_win3_count_column = 31;
    public static int game_speed_standarded_column = 32;
    public static int game_speed_avarage_column = 33;
    public static int set_direct_serve_success_count_column = 34;
    public static int set_hit_success_count_column = 35;
    public static int set_con_win3_count_column = 36;
    public static int set_speed_standarded_column = 37;
    public static int set_speed_avarage_column = 38;

    public static final String player_unknown = "unknown";

    public static int is_server_value = 1;

    public static void main(String[] args) {
        loadP2Statistics();
        fillData();
    }

    public static void loadP1Statistics() {
        orig_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/full_data.xlsx";
        dest_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/P1.xlsx";
        orig_xlsx_sheet = 0;
        dest_xlsx_sheet = 0;
        set_no_column = 4;
        game_no_column = 5;
        point_no_column = 6;
        server_column = 13;
        ace_column = 20;
        win_column = 22;
        point_score_column = 11;
        point_sum_column = 20;
        point_con_win3_column = 21;
        point_con_fall3_column = 22;
        current_game_win_column = 9;
        game_con_win3_column = 23;
        game_con_fall3_column = 24;
        current_set_win_column = 7;
        set_con_win2_column = 25;
        set_con_fall2_column = 26;
        ace_sum_d_server_count_column = 27;
        player_name_column = 1;

        is_server_value = 1;
    }

    public static void loadP2Statistics() {
        orig_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/full_data.xlsx";
        dest_xlsx = "/home/jackmaxpale/Desktop/Projects/ICM/2024_MCM-ICM_Problems/parsedData/P2.xlsx";
        orig_xlsx_sheet = 0;
        dest_xlsx_sheet = 0;
        set_no_column = 4;
        game_no_column = 5;
        point_no_column = 6;
        server_column = 13;
        ace_column = 21;
        win_column = 23;
        point_score_column = 12;
        point_sum_column = 20;
        point_con_win3_column = 21;
        point_con_fall3_column = 22;
        current_game_win_column = 10;
        game_con_win3_column = 23;
        game_con_fall3_column = 24;
        current_set_win_column = 8;
        set_con_win2_column = 25;
        set_con_fall2_column = 26;
        ace_sum_d_server_count_column = 27;
        player_name_column = 1;

        is_server_value = 2;
    }

    public static void fillData() {
        try (FileInputStream fis = new FileInputStream(orig_xlsx);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(orig_xlsx_sheet);

            Iterator<Row> iterator = sheet.iterator();
            //dest
            FileInputStream fis1 = new FileInputStream(dest_xlsx);
            Workbook workbook1 = new XSSFWorkbook(fis1);
            Sheet sheet1 = workbook1.getSheetAt(dest_xlsx_sheet);
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
            int win_sum = 0;

            //Summary data
            int game_direct_serve_success_count = 0;
            int game_hit_success_count = 0;
            int game_con_win3_count = 0;
            int set_direct_serve_success_count = 0;
            int set_hit_success_count = 0;
            int set_con_win3_count = 0;
            int set_speed_standarded = 0;
            int set_speed_avarage = 0;
            List<Double> game_speed_list = new ArrayList<>();
            List<Double> set_speed_list = new ArrayList<>();

            String player_name = player_unknown;
            String player_name_previous = player_unknown;
            Row previousRow = null;
            Row previousRow1 = null;
            List<Row> pendingRemoveRows=new ArrayList<>();
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Row nextRow1 = iterator1.next();
                System.out.println("------------------\n Index " + nextRow.getRowNum());
//                if (nextRow.getRowNum() >= 100) {
//                    break;
//                }

                previous3_score = previous2_score;
                previous2_score = previous_score;
                previous_score = current_score;


                point_no_previous = point_no;
                set_no_previous = set_no;
                game_no_previous = game_no;
                point_sum_previous3 = point_sum_previous2;
                point_sum_previous2 = point_sum_previous;
                point_sum_previous = point_sum;

                player_name_previous = player_name;


                set_no = (int) nextRow.getCell(set_no_column).getNumericCellValue();
                game_no = (int) nextRow.getCell(game_no_column).getNumericCellValue();
                point_no = (int) nextRow.getCell(point_no_column).getNumericCellValue();
                player_name = nextRow.getCell(player_name_column).getStringCellValue();
                System.out.println("Current player " + player_name + " Previous Player " + player_name_previous);
                if (!player_name.equals(player_name_previous)) {
                    if (!player_name.equals(player_unknown)) {
                        if (!player_name_previous.equals(player_unknown)) {
                            System.out.println("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^6\nFresh Player " + player_name);
                            previous_score = 0;
                            previous2_score = 0;
                            previous3_score = 0;

                            previous_game_win = 0;
                            previous2_game_win = 0;
                            previous3_game_win = 0;

                            previous_set_win = 0;
                            previous2_set_win = 0;

                            current_score = 0;
                            current_game_win = 0;
                            current_set_win = 0;

                            point_sum = 0;
                            point_sum_previous = 0;
                            point_sum_previous2 = 0;
                            point_sum_previous3 = 0;

                            point_no = 1;
                            set_no = 1;
                            game_no = 1;
                            point_no_previous = 1;
                            set_no_previous = 1;
                            game_no_previous = 1;

                            server_count = 0;
                            ace_sum = 0;
                            win_sum = 0;

                            game_direct_serve_success_count = 0;
                            game_hit_success_count = 0;
                            game_con_win3_count = 0;
                            set_direct_serve_success_count = 0;
                            set_hit_success_count = 0;
                            set_con_win3_count = 0;
                            set_speed_standarded = 0;
                            set_speed_avarage = 0;
                            game_speed_list.clear();
                            set_speed_list.clear();

                            player_name = player_unknown;
                            player_name_previous = player_unknown;
                            System.out.println("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^6");
                        }
                    }
                }

                if ((int) nextRow.getCell(server_column).getNumericCellValue() == is_server_value) {
                    server_count += 1;
                }

                if (game_no != game_no_previous) {
                    System.out.println("Clear game summary data");
                    game_direct_serve_success_count = 0;
                    game_hit_success_count = 0;
                    game_con_win3_count = 0;
                    //calculate speed details
                    double deviation = calculateStandardDeviation(game_speed_list);
                    double mean = calculateMean(game_speed_list);
                    System.out.println("deviation " + deviation + " mean " + mean);
                    previousRow1.getCell(game_speed_standarded_column).setCellValue(deviation);
                    previousRow1.getCell(game_speed_avarage_column).setCellValue(mean);
                    game_speed_list.clear();
                }
                if (set_no != set_no_previous) {
                    System.out.println("Clear set summary data");
                    set_hit_success_count = 0;
                    set_direct_serve_success_count = 0;
                    set_con_win3_count = 0;
                    //calculate speed details
                    double deviation = calculateStandardDeviation(set_speed_list);
                    double mean = calculateMean(set_speed_list);
                    System.out.println("deviation " + deviation + " mean " + mean);
                    previousRow1.getCell(set_speed_standarded_column).setCellValue(deviation);
                    previousRow1.getCell(set_speed_avarage_column).setCellValue(mean);
                    set_speed_list.clear();
                }
                //get speed
                //System.out.println("Speed Type " + nextRow.getCell(speed_column).getCellType());
                if (nextRow.getCell(speed_column).getCellType() == CellType.NUMERIC) {
                    System.out.println("Speed " + nextRow.getCell(speed_column).getNumericCellValue());
                    game_speed_list.add(nextRow.getCell(speed_column).getNumericCellValue());
                    set_speed_list.add(nextRow.getCell(speed_column).getNumericCellValue());

                }
                if ((int) nextRow.getCell(ace_column).getNumericCellValue() == 1) {
                    ace_sum += 1;
                    game_direct_serve_success_count += 1;
                    set_direct_serve_success_count += 1;
                }
                if ((int) nextRow.getCell(win_column).getNumericCellValue() == 1) {
                    win_sum += 1;
                    game_hit_success_count += 1;
                    set_hit_success_count += 1;
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
                Cell cell = nextRow.getCell(point_score_column);
                if (cell.getCellType() == CellType.NUMERIC) {
                    current_score = (int) nextRow.getCell(point_score_column).getNumericCellValue();
                    System.out.println(nextRow.getRowNum() + " current_score " + current_score);
                    if (current_score > previous_score) {
                        point_sum += 1;
                        nextRow1.getCell(point_sum_column).setCellValue(point_sum);
                        System.out.println(nextRow1.getRowNum() + " New  " + point_sum);

                    } else {
                        nextRow1.getCell(point_sum_column).setCellValue(point_sum);
                        System.out.println(nextRow1.getRowNum() + " New  " + point_sum);
                    }
                } else {
                    nextRow1.getCell(point_sum_column).setCellValue(point_sum);
                    System.out.println(nextRow1.getRowNum() + " New  " + point_sum);
                }
                System.out.println(nextRow.getRowNum() + " point_sum " + point_sum + "\n"
                        + " point_sum_previous " + point_sum_previous + "\n"
                        + " point_sum_previous2 " + point_sum_previous2 + "\n"
                        + " point_sum_previous3 " + point_sum_previous3 + "\n");

                //parse point_con_win3
                if (point_sum > point_sum_previous &&
                        point_sum_previous > point_sum_previous2 &&
                        point_sum_previous2 > point_sum_previous3) {
                    System.out.println(nextRow1.getRowNum() + " Point Con Win3  ");
                    nextRow1.getCell(point_con_win3_column).setCellValue(1);
                    game_con_win3_count += 1;
                } else {
                    game_con_win3_count = 0;
                    nextRow1.getCell(point_con_win3_column).setCellValue(0);
                }
                //parse point_con_fall3
                if (point_sum <= point_sum_previous &&
                        point_sum_previous <= point_sum_previous2 &&
                        point_sum_previous2 <= point_sum_previous3) {
                    if (point_sum > 0) {
                        System.out.println(nextRow1.getRowNum() + " Point Con Fall3  ");
                        nextRow1.getCell(point_con_fall3_column).setCellValue(1);
                    } else {
                        nextRow1.getCell(point_con_fall3_column).setCellValue(0);
                    }
                } else {
                    nextRow1.getCell(point_con_fall3_column).setCellValue(0);
                }
                System.out.println(nextRow.getRowNum() + " game_win " + current_game_win + "\n"
                        + " previous_game_win " + previous_game_win + "\n"
                        + " previous2_game_win " + previous2_game_win + "\n"
                        + " previous3_game_win " + previous3_game_win + "\n");
                //parse game_con win3
                current_game_win = (int) nextRow.getCell(current_game_win_column).getNumericCellValue();
                if (current_game_win > previous_game_win &&
                        previous_game_win > previous2_game_win &&
                        previous2_game_win > previous3_game_win) {
                    System.out.println(nextRow1.getRowNum() + " Game Con Win3  ");
                    nextRow1.getCell(game_con_win3_column).setCellValue(1);
                    set_con_win3_count += 1;
                } else {
                    set_con_win3_count = 0;
                    nextRow1.getCell(game_con_win3_column).setCellValue(0);
                }
                //parse game_con_fall3
                if (current_game_win <= previous_game_win &&
                        previous_game_win <= previous2_game_win &&
                        previous2_game_win <= previous3_game_win) {
                    if (current_game_win > 0) {
                        System.out.println(nextRow1.getRowNum() + " Game Con Fall3  ");
                        nextRow1.getCell(game_con_fall3_column).setCellValue(1);
                    } else {
                        nextRow1.getCell(game_con_fall3_column).setCellValue(0);
                    }

                } else {
                    nextRow1.getCell(game_con_fall3_column).setCellValue(0);
                }
                System.out.println(nextRow.getRowNum() + " set_win " + current_set_win + "\n"
                        + " previous_set_win " + previous_set_win + "\n"
                        + " previous2_set_win " + previous2_set_win + "\n");
                //parse set_con_win2
                current_set_win = (int) nextRow.getCell(current_set_win_column).getNumericCellValue();
                if (current_set_win > previous_set_win &&
                        previous_set_win > previous2_set_win) {
                    System.out.println(nextRow1.getRowNum() + " Set Con Win2  ");
                    nextRow1.getCell(set_con_win2_column).setCellValue(1);
                } else {
                    nextRow1.getCell(set_con_win2_column).setCellValue(0);
                }
                //parse set_con_fall2
                if (current_set_win <= previous_set_win &&
                        previous_set_win <= previous2_set_win) {
                    if (current_set_win > 0) {
                        System.out.println(nextRow1.getRowNum() + " Set Con Fall2  ");
                        nextRow1.getCell(set_con_fall2_column).setCellValue(1);
                    } else {
                        nextRow1.getCell(set_con_fall2_column).setCellValue(0);
                    }
                } else {
                    nextRow1.getCell(set_con_fall2_column).setCellValue(0);
                }


                double rate = 0;
                if (server_count != 0) {
                    rate = (double) ace_sum / server_count;
                }
                System.out.println(nextRow.getRowNum() + " ace_sum " + ace_sum + "\n"
                        + " server_count " + server_count + "\n"
                        + " ace_sum/server_count " + rate + "\n");
                //parse ace/server_count
                nextRow1.getCell(ace_sum_d_server_count_column).setCellValue(rate);

                //parse game summary info
                System.out.println(nextRow.getRowNum() + " game_direct_serve_success_count " + game_direct_serve_success_count + "\n"
                        + " game_hit_success_count " + game_hit_success_count + "\n"
                        + " game_con_win3_count " + game_con_win3_count + "\n");
                nextRow1.getCell(game_direct_serve_success_count_column).setCellValue(game_direct_serve_success_count);
                nextRow1.getCell(game_hit_success_count_column).setCellValue(game_hit_success_count);
                nextRow1.getCell(game_con_win3_count_column).setCellValue(game_con_win3_count);

                //parse set summary info
                System.out.println(nextRow.getRowNum() + " set_direct_serve_success_count " + set_direct_serve_success_count + "\n"
                        + " set_hit_success_count " + set_hit_success_count + "\n"
                        + " set_con_win3_count " + set_con_win3_count + "\n");
                nextRow1.getCell(set_direct_serve_success_count_column).setCellValue(set_direct_serve_success_count);
                nextRow1.getCell(set_hit_success_count_column).setCellValue(set_hit_success_count);
                nextRow1.getCell(set_con_win3_count_column).setCellValue(set_con_win3_count);

                previousRow = nextRow;
                previousRow1 = nextRow1;

                //
                if (game_no == game_no_previous) {
                    if (previousRow1!=null) {
                        pendingRemoveRows.add(previousRow1);
                    }
                }
            }
            for (Row row : pendingRemoveRows) {
                sheet1.removeRow(row);
            }
            try (FileOutputStream outputStream = new FileOutputStream(dest_xlsx)) {
                workbook1.write(outputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        // Now, the 'head' variable points to the head of the linked list of DayMarketState objects
    }

    public static double calculateStandardDeviation(List<Double> data) {
        System.out.println("Data size " + data.size());
        int size = data.size();

        if (size < 2) {
            return 0;
        }

        double sum = 0.0;
        double sumSquaredDiff = 0.0;

        // Calculate the sum of the values
        for (double value : data) {
            sum += value;
        }

        // Calculate the mean (average) of the values
        double mean = sum / size;

        // Calculate the sum of the squared differences from the mean
        for (double value : data) {
            double diff = value - mean;
            sumSquaredDiff += diff * diff;
        }

        // Calculate the variance
        double variance = sumSquaredDiff / (size - 1);

        // Calculate the standard deviation as the square root of the variance
        return Math.sqrt(variance);
    }

    public static double calculateMean(List<Double> data) {
        int size = data.size();

        if (size == 0) {
            return 0;
        }

        double sum = 0.0;

        // Calculate the sum of the values
        for (double value : data) {
            sum += value;
        }

        // Calculate the mean (average) of the values
        return sum / size;
    }
}