package business;

import java.io.*;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.sql.*;
import java.util.*;
import java.util.function.Function;
import java.util.regex.*;

public class Business {

    private static Scanner sc;
    private static String[] args;
    private static boolean run = true;
    private static boolean conRun = false;
    private static boolean connected = false;
    private static String url;
    private static Statement stmt;
    private static List<Salesman> salesmen;

    public static void main(String[] arguments) throws Exception {
        run();
    }

    private static void run() throws Exception {
        sc = new Scanner(System.in);
        while (run) {
            takeInput();
        }
        sc.close();
        System.out.println("Application terminated.");
    }

    private static void takeInput() throws Exception {

        String input;
        System.out.println("\nWhat would you like to do? Connection status: " + getConStat());
        input = sc.nextLine();
        if (input.isEmpty())
            input = sc.nextLine();
        List<String> argsList = new ArrayList<>();
        Matcher m = Pattern.compile("([^\"]\\S*|\".+?\")\\s*").matcher(input);
        while (m.find()) {
            argsList.add(m.group(1).replaceAll("\"", ""));
        }
        args = new String[argsList.size()];
        args = argsList.toArray(args);
        readArgs();

    }

    private static void readArgs() throws Exception {
        String[] commands = {"collect", "export", "import", "print", "delete", "help", "exit", "quit", "connect", "disconnect", "url"};
        if (args[0].equalsIgnoreCase(commands[0])) {
            // Collect - iterate directory and add emails to database -->   Args: collect tableName "pathOfDirectory"
            collect();
        } else if (args[0].equalsIgnoreCase(commands[1])) {
            // Export - read database and export to CSV file -->            Args: export tableName "pathOfCsv"
            export();
        } else if (args[0].equalsIgnoreCase(commands[2])) {
            // Import - read CSV file and import into database -->          Args: import tableName "pathOfCsv"
            uploadCustomerList(importCsv(args[2]), args[1]);
        } else if (args[0].equalsIgnoreCase(commands[3])) {
            // Print - print database to command line -->                   Args: print tableName
            printList(getTable(args[1]));
        } else if (args[0].equalsIgnoreCase(commands[4])) {
            // Delete - delete specified database table -->                 Args: delete tableName
            deleteTable(args[1]);
        } else if (args[0].equalsIgnoreCase(commands[5])) {
            // Help - show list of commands -->                         	Args: help
            help();
        } else if (args[0].equalsIgnoreCase(commands[6]) | args[0].equalsIgnoreCase(commands[7])) {
            // Quit/Exit - end application -->								Args: exit | quit
            System.out.println("Exiting application...");
            conRun = run = false;
        } else if (args[0].equalsIgnoreCase(commands[8])) {
            // Connect - login and connect to MySQL database -->			Args: connect
            connect();
        } else if (args[0].equalsIgnoreCase(commands[9])) {
            // Disconnect - disconnect from MySQL database --> 				Args: disconnect
            disconnect();
        } else if (args[0].equalsIgnoreCase(commands[10])) {
            // Url - enter MySQL database url and create config.properties	Args: url
            createProperties();
        } else {
            // Invalid argument
            System.out.println("\"" + args[0] + "\" is not a valid argument. Type \"help\" for a list of commands.");
        }
    }

    private static void collect() throws Exception {
        if (connected) {
            if (args[1].equalsIgnoreCase("all")) {
                getSalesmen();
                for (Salesman currentSalesman : salesmen) {
                    uploadCustomerList(iterateSpreadsheets(currentSalesman.getDirectory()), currentSalesman.getTableName());
                    System.out.println("\nCollection finished! All collected emails added to database table: " + currentSalesman.getDirectory() + "\n");
                }
            } else {
                if (args.length <3) { // isn't that cute
                    getSalesmen();
                    for (Salesman currentSalesman : salesmen) {
                        if (currentSalesman.getTableName().equalsIgnoreCase((args[1]))) {
                            uploadCustomerList(iterateSpreadsheets(currentSalesman.getDirectory()), args[1]);
                            System.out.println("\nCollection finished! All collected emails added to database table: " + currentSalesman.getDirectory() + "\n");
                        }
                    }
                } else {
                    uploadCustomerList(iterateSpreadsheets(args[2]), args[1]);
                    System.out.println("\nCollection finished! All collected emails added to database table: " + args[2] + "\n");
                }
            }

        } else {
            System.out.println("You must be connected to a MySQL database to execute this command.");
        }
    }

    private static void export() throws Exception {
        if (connected) {
            if (args[1].equalsIgnoreCase("all")) {
                getSalesmen();
                for (Salesman currentSaleman : salesmen) {
                    exportList(getTable(currentSaleman.getTableName()), currentSaleman.getExport());
                }
                System.out.println("\nFull export complete!");
            } else {
                if (args.length <3) {
                    getSalesmen();
                    for ( Salesman currentSalesman : salesmen) {
                        if (currentSalesman.getTableName().equalsIgnoreCase(args[1])) {
                            exportList(getTable(args[1]), currentSalesman.getExport());
                        }
                    }
                } else {
                    exportList(getTable(args[1]), args[2]);
                }
            }
        } else {
            System.out.println("You must be connected to a MySQL database to execute this command.");
        }
    }

    private static List<Salesman> getSalesmen() throws SQLException {
        salesmen = new ArrayList<>();
        ResultSet rs = stmt.executeQuery("SELECT * FROM salesmen");
        while (rs.next()) {
            Salesman currentSalesman = new Salesman(rs.getString("tableName"),
                    rs.getString("dirPath").replace("\"", ""),
                    rs.getString("csvName").replace("\"", ""));
            salesmen.add(currentSalesman);
        }
        return salesmen;
    }

    private static String getConStat() {
        String conStat;
        if (connected) {
            conStat = "Connected";
        } else {
            conStat = "Not Connected";
        }
        return conStat;
    }

    private static void connect() throws Exception {
        if (!connected) {
            loadProperties();
            System.out.println("Enter your username:");
            String username = sc.next();
            System.out.println("Enter yout password:");
            String password = sc.next();
            System.out.println("Connecting database...");
            try {
                Connection con = DriverManager.getConnection(url, username, password);
                stmt = con.createStatement();

                connected = conRun = true;

                System.out.println("Database connected!");

                conRun();

                con.close();
                stmt.close();
                connected = false;
                System.out.println("Database disconnected.");

            } catch (SQLException e) {
                if (e.getMessage().contains("Access denied")) {
                    System.out.println("Access denied. Invalid username or password.");
                } else if (e.getMessage().contains("url")) {
                    System.out.println("Connection failed. Please make sure that your \"config.properties\" file is located in \"C:\\emaildb\\\" and that it contains the line \"URL=[yourDatabaseUrl]\"");
                } else {
                    throw new IllegalStateException("Cannot connect database!", e);
                }
            }
        } else {
            System.out.println("You are already connected to the database.");
        }
    }

    private static void createProperties() throws IOException {
        Properties props = new Properties();
        String propUrl;
        System.out.println("Please enter your database url. It should look something like: jdbc:mysql://<host>:<port>/<database_name>");
        propUrl = sc.next();
        props.setProperty("URL", propUrl);
        File dir = new File("C:\\emaildb");
        File configFile = new File("C:\\emaildb\\config.properties");
        if (!dir.exists()) {
            System.out.println("Creating directory: \"" + dir.getAbsolutePath() + "\"...");
            boolean result = false;
            try {
                dir.mkdir();
                result = true;
            } catch (SecurityException se) {
                System.out.println("This program does not have permission to create the directory. Either close and run as administrator, or create dir \"C:\\emaildb\" and try again.");
            }
            if (result) {
                System.out.println("Directory created!");
            }
        }
        if (!configFile.exists()) {
            System.out.println("Writing properties file...");
            try (FileWriter writer = new FileWriter(configFile)) {
                props.store(writer, "Database settings");
                System.out.println("Properties file created!");
            }
        } else {
            System.out.println("\"" + configFile.getAbsolutePath() + "\" already exists.");

        }
    }

    private static void disconnect() {
        System.out.println("Disconnecting database...");
        conRun = false;
    }

    private static void conRun() throws Exception {
        while (conRun) {
            takeInput();
        }
    }

    private static void loadProperties() {
        Properties prop = new Properties();
        try (InputStream input = new FileInputStream("C:\\Emaildb\\config.properties")) {
            prop.load(input);
            url = prop.getProperty("URL");
        } catch (IOException e) {
            System.out.println("Problem reading 'config.properties' file!");
        }

    }

    private static List<Customer> iterateSpreadsheets(String directory) throws Exception {
        if (connected) {
            List<Customer> customerList = new ArrayList<Customer>();
            File dir = new File(directory);
            File[] directoryListing = dir.listFiles();
            if (directoryListing != null) {
                for (int i = 0; i < directoryListing.length; i++) {
                    double all = directoryListing.length;
                    double percent = (double) i / all * 100;
                    String perc = new DecimalFormat("#.##").format(percent);
                    System.out.println("Processing: " + perc + "% finished. Now proccessing file " + i + " / " +
                            directoryListing.length + " : " + directoryListing[i].getName());
                    Customer currentCustomer;
                    if (directoryListing[i].getName().contains(".xlsx") & !directoryListing[i].getName().contains("~$")) {
                        currentCustomer = findCustomerInfo(directoryListing[i].toString());
                        if (currentCustomer != null) {
                            customerList.add(currentCustomer);
                        }
                    }
                }
            } else {
                System.out.println("There is an error with this directory. It may empty or you may be wrong about something that's going on here.");
            }
            return customerList;

        } else {
            return null;
        }
    }

    private static void uploadCustomerList(List<Customer> customerList, String tableName) throws SQLException {
        createTableIfNotExists(tableName);
        for (int i = 0; i < customerList.size(); i++) {
            Customer currentCustomer = customerList.get(i);
            double all = customerList.size();
            double percent = (double) i / all * 100;
            String perc = new DecimalFormat("#.##").format(percent);
            System.out.println("Uploading: " + perc + "% finished. Adding " + currentCustomer.getName());
            insertSql(currentCustomer, tableName);
        }
    }

    private static void insertSql (Customer customer, String tableName) throws SQLException {

        customer.setName(customer.getName().replace("'", "''"));
        customer.setCompany(customer.getCompany().replace("'", "''"));
        stmt.executeUpdate("REPLACE INTO " + tableName + " VALUES ('" + customer.getName() + "', '" + customer.getCompany()
                + "', '" + customer.getEmail() + "');");
        System.out.println(customer.getName() + " added to database!");
    }

    private static void createTableIfNotExists(String tableName) throws SQLException {
        stmt.executeUpdate("CREATE TABLE IF NOT EXISTS " + tableName + " (\n" +
                "\tname varchar(255) NOT NULL,\n" +
                "    company varchar(255) NOT NULL,\n" +
                "    email varchar(255) NOT NULL,\n" +
                "    PRIMARY KEY (email)\n" +
                ");");
    }

    private static Customer findCustomerInfo(String filename) throws IOException, NotOfficeXmlFileException {

        File excel =  new File (filename);

        // Initialize customer variables
        String company = "";
        String name ="";
        String email = "";
        Customer currentCustomer = null;

        try (
                FileInputStream fis = new FileInputStream(excel);
                XSSFWorkbook wb = new XSSFWorkbook(fis)
        ){

            for (Sheet currentSheet : wb) {
                if (currentSheet.getSheetName().equals("ORIGINAL")) {
                    XSSFSheet ws = wb.getSheet("ORIGINAL");
                    int rowNum = ws.getLastRowNum() + 1;
                    int colNum = ws.getRow(0).getLastCellNum();
                    // Initialize cell data variables
                    String value = "";
                    // Iterate sheet
                    for(int i = 0; i <rowNum; i++){
                        XSSFRow row = ws.getRow(i);
                        for (int j = 0; j < colNum; j++){
                            XSSFCell cell = row.getCell(j);
                            if (cell != null) {
                                // Check cell type and get value (handles formulaic cells)
                                value = getCellValue(cell);
                                // Find company and name
                                if(value.contains("Sold to")) {
                                    XSSFCell compCell = row.getCell(2);
                                    company = getCellValue(compCell);
                                } else if(value.contains("Ordered By")) {
                                    XSSFCell nameCell = row.getCell(4);
                                    name = getCellValue(nameCell);
                                }
                                // Find email and build customer if found
                                if (Validator.validateEmail(value) != null) {
                                    value = Validator.validateEmail(value);
                                    email = value;
                                    currentCustomer = new Customer(company, name, email);
                                    System.out.println(filename + "\n" + company + "\n" + name + "\n" + email + "\n");
                                }
                            }
                        }
                    }
                    return currentCustomer;
                }
            }
            return null;
        } catch (Exception e){
            e.printStackTrace();
            System.out.println("There was a problem reading the file: " + filename);
            return currentCustomer;
        }
    }

    @SuppressWarnings({ "deprecation", "static-access" })
    private static String getCellValue(XSSFCell cell) {
        String value = "";
        if(cell.getCellType() == cell.CELL_TYPE_FORMULA) {
            if (cell.getCachedFormulaResultType() == cell.CELL_TYPE_NUMERIC) {
                value = String.valueOf(cell.getNumericCellValue());
            } else if (cell.getCachedFormulaResultType() == cell.CELL_TYPE_STRING){
                value = cell.getStringCellValue();
            }
        } else{
            value = cell.toString();
        }
        return value;
    }

    private static List<Customer> getTable(String tableName) throws SQLException {
        List<Customer> customerList = new ArrayList<>();
        String query = "SELECT * FROM " + tableName;
        ResultSet rs = stmt.executeQuery(query);
        while (rs.next()) {
            Customer newCustomer = new Customer(rs.getString("company"),
                    rs.getString("name"), rs.getString("email"));
            customerList.add(newCustomer);
        }
        return customerList;
    }

    private static void exportList(List<Customer> customerList, String fileName) throws Exception {
        final String DIR_PATH = "c:\\emaildb\\exports\\";
        final String FILE_HEADER = "\"Company\",\"Customer Name\",\"Email Address\",\"" + "Exported: " + getDateTime() + "\"";
        final String COMMA_QUOTE_DELIMITER = "\",\"";
        final String QUOTE_DELIMITER = "\"";
        final String NEW_LINE_SEPARATOR = "\n";

        File dir = new File(DIR_PATH);
        if (!dir.exists()) dir.mkdir();
        String fullPath = DIR_PATH + fileName;

        try (FileWriter fileWriter = new FileWriter(fullPath, false)){
            fileWriter.append(FILE_HEADER);
            fileWriter.append(NEW_LINE_SEPARATOR);
            for (Customer currentCustomer : customerList) {
                fileWriter.append(QUOTE_DELIMITER);
                fileWriter.append(currentCustomer.getName());
                fileWriter.append(COMMA_QUOTE_DELIMITER);
                fileWriter.append(currentCustomer.getCompany());
                fileWriter.append(COMMA_QUOTE_DELIMITER);
                fileWriter.append(currentCustomer.getEmail());
                fileWriter.append(QUOTE_DELIMITER);
                fileWriter.append(NEW_LINE_SEPARATOR);
            }
        } catch (Exception e) {
            System.out.println("There was an error writing your file.");
            e.printStackTrace();
        } finally {
            try {
                customerList.clear();
            } catch (Exception e) {
                System.out.println("Error closing/flushing file.");
                e.printStackTrace();
            }
        }
        System.out.println(fileName + " exported!");
    }

    private static List<Customer> importCsv(String fileName) {
        boolean stop = false;
        final String FILE_HEADER = "\"Company\",\"Customer Name\",\"Email Address\",\"";
        final String COMMA_QUOTE_DELIMITER = "\",\"";
        List<Customer> customerList = new ArrayList<>();
        while (!stop) {
            try (BufferedReader fileReader = new BufferedReader(new FileReader(fileName))) {
                String line;
                String header;
                header = fileReader.readLine();
                if (!header.contains(FILE_HEADER)) {
                    System.out.println("This file cannot be imported.");
                    stop = true;
                    return null;
                }
                //Read the file line by line starting from the second line
                while ((line = fileReader.readLine()) != null) {
                    //Get all tokens available in line
                    String[] currentTokens = line.split(COMMA_QUOTE_DELIMITER);
                    if (currentTokens.length > 0) {
                        Customer currentCustomer = new Customer(currentTokens[1], currentTokens[0].replace("\"", ""),currentTokens[2].replace("\"", ""));
                        customerList.add(currentCustomer);
                    }
                }
            }
            catch (Exception e) {
                System.out.println("Error in CsvFileReader !!!");
                e.printStackTrace();
            }
            stop = true;
        }
        System.out.println("Done!");
        return customerList;
    }

    private static void printList(List<Customer> customerList) {

        final String LINE_BREAK = "+---------------------------------+--------------------------------------------+------------------------------+";
        System.out.println(LINE_BREAK);
        System.out.println("\t      name        \t\t\t\t       company          \t\t            email");
        System.out.println(LINE_BREAK);
        for (Customer customer : customerList) {
            System.out.print("\t\t  " + customer.getName());
            printTabs(customer.getName());
            System.out.print(customer.getCompany());
            printTabs(customer.getCompany());
            System.out.println(customer.getEmail());
        }
        System.out.println(LINE_BREAK);
    }

    private static void printTabs(String s) {
        if (s.length() <= 5) {
            System.out.print("\t\t\t\t\t\t\t\t");
        } else if (s.length() <= 9) {
            System.out.print("\t\t\t\t\t\t\t");
        } else if (s.length() <= 13){
            System.out.print("\t\t\t\t\t\t");
        } else if (s.length() <= 17){
            System.out.print("\t\t\t\t\t");
        } else if (s.length() <= 21){
            System.out.print("\t\t\t\t");
        } else if (s.length() <= 25){
            System.out.print("\t\t\t");
        } else if (s.length() <= 29){
            System.out.print("\t\t");
        } else {
            System.out.print("\t");
        }
    }

    private static void deleteTable(String tableName) throws SQLException {
        System.out.println("Are you sure you want to delete the table " + tableName + "?\n" +
                "This table along with all data will be lost.");
        try (Scanner sc = new Scanner(System.in)) {
            String input = sc.next();
            Pattern pattern = Pattern.compile("y|yes|yeah|uh-hu|positive|affirmative|indeed|definitely", Pattern.CASE_INSENSITIVE);
            Matcher matcher = pattern.matcher(input);
            if (matcher.find()) {
                stmt.executeUpdate("DROP TABLE "  + tableName);
                System.out.println("Table " + tableName + " has been deleted!");
            }
        }

    }

    private static void help() {
        System.out.println("This program is for working with email databases. It works by giving it arguments to complete specific tasks.\n");
        System.out.println("Connect     [connect] - Connect to the email database. You will be prompted for your username and password");
        System.out.println("Disconnect  [disconnect] - Disconnect from the email database");
        System.out.println("Collect     [collect tableName \"directoryPath\"] - Collects all emails from work orders in directory and stores them in the database table");
        System.out.println("Exoprt      [export tableName \"csvFilePath\"] - Exports database table to CSV file");
        System.out.println("Import      [import tableName \"csvFilePath\"] - Imports CSV file into database table");
        System.out.println("Print       [print tableName] - Prints database table in the command line");
        System.out.println("Delete      [delete tableName] - Deletes database table");
        System.out.println("Url         [url] - Enter your MySQL database url and create a config.properties file (required for connecting to your database)");
        System.out.println("Help        [help] - Helpful information on how to use this program");
        System.out.println("Exit        [exit | quit] - Exits the application");
        System.out.println();
    }

    private static String getDateTime() {
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyy/MM/dd HH:mm:ss");
        LocalDateTime now = LocalDateTime.now();
        return dtf.format(now);
    }
}