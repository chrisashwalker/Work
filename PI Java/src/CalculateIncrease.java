import java.io.File;
import java.io.FileNotFoundException;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Scanner;
import java.util.HashMap;
import java.util.ArrayList;
import java.time.format.DateTimeFormatter;
import java.lang.StringBuilder;

public class CalculateIncrease{
    public static void main(String[] args){
        var fs_types = new HashMap<String, String>();
        var ca_types = new HashMap<String, String>();
        var gmp_types = new HashMap<String, String>();
        var fs_multipliers = new HashMap<String, String>();
        var ca_multipliers = new HashMap<String, String>();
        var rates_lines = new ArrayList<String>();
        var types_lines = new ArrayList<String>();
        var new_csv_lines = new ArrayList<String>();
        var new_gmp_lines = new ArrayList<String>();
        var new_rates_lines = new ArrayList<String>();
        var problems = new ArrayList<String>();
        var date_format = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        var increase_date = LocalDate.parse("10/04/2021", date_format);
        var increase_end_date = LocalDate.parse("31/12/9999", date_format);
        var gmp_date = LocalDate.parse("06/04/2021", date_format);
        var gmp_end_date = LocalDate.parse("31/12/9999", date_format);
        var gmp_rate = 0.01;
        var written_lines = "";
        var gmp_run = false;
        var pens_inc_run = false;
        var gmp_done = false;
        var pens_inc_done = false;

        if (increase_date.isAfter(gmp_date)) {
            increase_end_date = LocalDate.parse("31/12/9999", date_format);
            gmp_end_date = increase_date.minusDays(1);
            gmp_run = true;
        }
        else if (increase_date.isBefore(gmp_date)) {
            increase_end_date = gmp_date.minusDays(1);
            gmp_end_date = LocalDate.parse("31/12/9999", date_format);
            pens_inc_run = true;
        }

        try {
            var f = new File("C:\\Users\\Christopher\\OneDrive\\Documents\\Code\\Work\\PI Java\\multipliers.csv");
            var scan = new Scanner(f);
            while (scan.hasNextLine()) {
                rates_lines.addAll(Arrays.asList(scan.nextLine().split(",")));
            }
            rates_lines.remove(0);
        }
        catch (FileNotFoundException e) {
            System.out.println("File not found");
        }

        try {
            var f = new File("C:\\Users\\Christopher\\OneDrive\\Documents\\Code\\Work\\PI Java\\wage_types.csv");
            var scan = new Scanner(f);
            while (scan.hasNextLine()) {
                if (scan.nextLine().contains("FS")) {
                    var wage_type_builder = new StringBuilder();
                    var fed_wage_type_builder = new StringBuilder();
                    // TODO: Need a more robust way to get substrings
                    for (var i=0; i < 4; i++) {
                        wage_type_builder.append(scan.nextLine().charAt(i));
                        fed_wage_type_builder.append(scan.nextLine().charAt(i + 8));
                    }
                    var wage_type = wage_type_builder.toString();
                    var fed_wage_type = fed_wage_type_builder.toString();
                    fs_types.put(wage_type, fed_wage_type);
                }
                types_lines.addAll(Arrays.asList(scan.nextLine().split(",")));
            }
            types_lines.remove(0);
        }
        catch (FileNotFoundException e) {
            System.out.println("File not found");
        }

        try {
            var f = new File("C:\\Users\\Christopher\\OneDrive\\Documents\\Code\\Work\\PI Java\\pay_records.csv");
            var scan = new Scanner(f);
            while (scan.hasNextLine()) {
                new_csv_lines.addAll(Arrays.asList(scan.nextLine().split(",")));
            }
            new_csv_lines.remove(0);
        }
        catch (FileNotFoundException e) {
            System.out.println("File not found");
        }

        while (!pens_inc_done || !gmp_done) {
            var inc_list = new ArrayList<String>();
            if (pens_inc_run & ! pens_inc_done) {
                for (var i = 0; i <= new_csv_lines.size(); i++) {
                    for (var j = 0; j <= rates_lines.size(); j++) {
                        // TODO: Add in other comparison
                        if (LocalDate.parse(rates_lines.get(j).split(",")[0], date_format)
                                .isBefore(LocalDate.parse(new_csv_lines.get(i).split(",")[19])) ||
                                LocalDate.parse(rates_lines.get(j).split(",")[0], date_format)
                                        .isEqual(LocalDate.parse(new_csv_lines.get(i).split(",")[19]))) {
                            fs_multipliers.put(new_csv_lines.get(i).split(",")[0],
                                    rates_lines.get(j).split(",")[2]);
                        }
                        if (LocalDate.parse(rates_lines.get(j).split(",")[0], date_format)
                                .isBefore(LocalDate.parse(new_csv_lines.get(i).split(",")[19])) ||
                                LocalDate.parse(rates_lines.get(j).split(",")[0], date_format)
                                        .isEqual(LocalDate.parse(new_csv_lines.get(i).split(",")[19]))) {
                            fs_multipliers.put(new_csv_lines.get(i).split(",")[0],
                                    rates_lines.get(j).split(",")[2]);
                        }
                    }
                }

            }
        }

        }
        }
