import java.io.File;
import java.time.LocalDate;
import java.util.Scanner;
import java.util.HashMap;
import java.util.ArrayList;
import java.time.format.DateTimeFormatter;

public class CalculateIncrease{
    public static void main(String[] args){
        HashMap<String, String> fs_types = new HashMap<String, String>();
        HashMap<String, String> ca_types = new HashMap<String, String>();
        HashMap<String, String> gmp_types = new HashMap<String, String>();
        HashMap<String, String> fs_multipliers = new HashMap<String, String>();
        HashMap<String, String> ca_multipliers = new HashMap<String, String>();
        ArrayList<String> new_csv_lines = new ArrayList<String>();
        ArrayList<String> new_gmp_lines = new ArrayList<String>();
        ArrayList<String> new_rates_lines = new ArrayList<String>();
        ArrayList<String> problems = new ArrayList<String>();
        DateTimeFormatter date_format = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        LocalDate increase_date = LocalDate.parse("10/04/2021", date_format);
        LocalDate increase_end_date = LocalDate.parse("10/04/2021", date_format);
        LocalDate gmp_date = LocalDate.parse("10/04/2021", date_format);
        LocalDate gmp_end_date = LocalDate.parse("10/04/2021", date_format);
        double gmp_rate = 0.01;
        String written_lines = "";
        boolean gmp_run = false;
        boolean pens_inc_run = false;
        boolean gmp_done = false;
        boolean pens_inc_done = false;
        }
        }