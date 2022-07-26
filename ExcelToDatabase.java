package wfr.sys.plugins;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.hxtt.b.e;

import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.postgresql.util.PSQLException;

public class ExcelToDatabase {

  public static void main(String[] args)
      throws IOException, SQLException, EncryptedDocumentException, InvalidFormatException {
    // Database connection
    String jdbcUrl = "jdbc:postgresql:ems?user=postgres&password=Maker@1";
    String path = "C:\\Users\\LENOVO\\Downloads\\v.xlsx";
    convertDataXlsx(path, jdbcUrl);
  }

  public static void convertDataXlsx(String path, String jdbcUrl)
      throws IOException, SQLException, EncryptedDocumentException, InvalidFormatException {

    // Criando instância do workbook para xlsx ou xls
    Workbook wb = WorkbookFactory.create(new File(path));

    for (int index = 0; index < wb.getNumberOfSheets(); index++) {
      // Lista para armazenar os valores das celulas do tipo String
      ArrayList<String> list = sweppLeafAllStrings(index, path);

      // Lista com os valores das colunas para criar a tabela
      ArrayList<String> columns = regexStringValidation(list, index, path);

      // Lista para armazenar os valores de todas as celulas
      ArrayList<Object> allValues = sweppLeafAllValues(index, path);

      // Variavel pora contar o número de colunas na primeira linha
      int noOfColumns = sweppLeafAllCollumns(index, path);

      // criando o objeto para receber a tabela
      Sheet sheet = wb.getSheetAt(index);

      // Verificando os tipos de cada coluna
      ArrayList<Object> dataTypesVerify = new ArrayList<>();
      /*
       * for (int i = 0; i < noOfColumns; i++) {
       * dataTypesVerify.add(sheet.getRow(1).getCell(i));
       * }
       */
      for (int i = noOfColumns; i < noOfColumns * 2; i++) {
        dataTypesVerify.add(allValues.get(i));
      }

      // Comando SQL para criar a tabela POSTGRES
      StringBuilder sql = new StringBuilder(
          "CREATE TABLE IF NOT EXISTS " + sheet.getSheetName() + "( " + "ID SERIAL PRIMARY KEY NOT NULL, ");

      Integer vrfCellSize = varcharSize(list); // Verificar a quantidade de caracteres nas celulas para criar um tamanho
                                               // no VARCHAR
      // Alimentando o comando CREATE para gerar a tabela
      int s = noOfColumns - 1;
      for (int j = 0; j < noOfColumns; j++) {
        String aux = columns.get(j);

        if (j == s) {
          if (isNumeric(dataTypesVerify.get(j).toString()) == true) {
            sql.append(aux + " float )");
          } else if (isNumeric(dataTypesVerify.get(j).toString()) == false) {
            if (checkIfDateIsValid(dataTypesVerify.get(j).toString())) {
              sql.append(aux + " date )");
            } else {
              sql.append(aux + " varchar(" + vrfCellSize + ") " + ")");
            }

          }

        } else {
          if (isNumeric(dataTypesVerify.get(j).toString()) == true) {
            sql.append(aux + " float, ");
          } else if (isNumeric(dataTypesVerify.get(j).toString()) == false) {
            if (checkIfDateIsValid(dataTypesVerify.get(j).toString())) {
              sql.append(aux + " date, ");
            } else {
              sql.append(aux + " varchar(" + vrfCellSize + "), ");
            }

          }

        }
      }

      // Comando SQL para inserir valores na tabela POSTGRES
      StringBuilder insert = new StringBuilder("INSERT INTO " + sheet.getSheetName() + "(");
      // Alimentando o INSERT
      int auxI = noOfColumns;
      for (int i = 0; i < noOfColumns; i++) {
        if (i == s) {
          insert.append(columns.get(i) + ")");
        } else {
          insert.append(columns.get(i) + ",");
        }
      }

      insert.append(" VALUES ( ");
      int count = allValues.size() - 1;
      int y = 2;
      for (int i = noOfColumns; i < allValues.size(); i++) {
        if (i == count) {
          if (isNumeric(allValues.get(i).toString()) == false) {
            insert.append("'" + allValues.get(i) + "'" + ");");
          } else {
            insert.append(allValues.get(i) + ");");
          }

        } else if (i == (noOfColumns * y) - 1) {
          if (emptyCellValidation(allValues.get(i)) == false && isNumeric(allValues.get(i).toString()) == false) {
            insert.append("'" + allValues.get(i) + "'" + "),(");
          } else {
            insert.append(allValues.get(i) + "),(");
          }
          y += 1;
        } else if (emptyCellValidation(allValues.get(i)) == false) {
          if (isNumeric(allValues.get(i).toString()) == false) {
            insert.append("'" + allValues.get(i) + "'" + ",");
          } else {
            insert.append(allValues.get(i) + ",");
          }
        } else if (emptyCellValidation(allValues.get(i)) == true) {
          insert.append("null" + ",");
        }
      }

      System.out.println();
      System.out.println(insert);
      System.out.println();
      System.out.println();
      System.out.println(sql);
      System.out.println();

      /*
       * Connection conn = DriverManager.getConnection(jdbcUrl);
       * Statement stat = conn.createStatement();
       * creatConnection(conn, stat, jdbcUrl, sql.toString(), insert.toString());
       */
    }
  }

  // Varrer a folha para receber todos os valores
  private static ArrayList<Object> sweppLeafAllValues(int index, String path) throws IOException {
    FileInputStream fis = new FileInputStream(new File(path));
    ArrayList<Object> allValues = new ArrayList<>();
    // criando uma instância do workbook
    XSSFWorkbook wb = new XSSFWorkbook(fis);
    // Capturando o index da folha
    XSSFSheet sheet = wb.getSheetAt(index);

    int addlimit = emptyRowValidation(sheet);

    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
    // para avaliar o tipo de célula
    FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
    int aux = 0;
    for (org.apache.poi.ss.usermodel.Row row : sheet) {
      for (Cell cell : row) {
        switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
          case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
              allValues.add(formatter.format(cell.getDateCellValue()).toString());
            } else {
              allValues.add(cell.getNumericCellValue());
            }
            break;
          case STRING:
            allValues.add(cell.getStringCellValue());
            break;
          case BLANK:
            allValues.add(null);
            break;
          case BOOLEAN:
            allValues.add(cell.getBooleanCellValue());
            break;
          case FORMULA:
            allValues.add(cell.getCellFormula());
          default:
            break;
        }
      }
      aux += 1;
      if (aux == addlimit + 1) {
        break;
      }
    }
    return allValues;
  }

  // Varrer a folha para receber todos as Strings
  private static ArrayList<String> sweppLeafAllStrings(int index, String path) throws IOException {
    FileInputStream fis = new FileInputStream(new File(path));
    ArrayList<String> list = new ArrayList<>();
    // criando uma instância do workbook
    XSSFWorkbook wb = new XSSFWorkbook(fis);
    // Capturando o index da folha
    XSSFSheet sheet = wb.getSheetAt(index);

    // para avaliar o tipo de célula
    FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();

    for (org.apache.poi.ss.usermodel.Row row : sheet) {
      for (Cell cell : row) {
        switch (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum()) {
          case STRING:
            list.add(cell.getStringCellValue());
            break;
          case BLANK:
            list.add(null);
            break;
          default:
            break;
        }
      }
    }
    return list;
  }

  // Varrer a folha para retornar o número de colunas
  private static Integer sweppLeafAllCollumns(int index, String path) throws IOException {
    int noOfColumns = 0;
    FileInputStream fis = new FileInputStream(new File(path));
    // criando uma instância do workbook
    XSSFWorkbook wb = new XSSFWorkbook(fis);
    // Capturando o index da folha
    XSSFSheet sheet = wb.getSheetAt(index);

    for (org.apache.poi.ss.usermodel.Row row : sheet) {
      for (Cell cell : row) {
        noOfColumns = sheet.getRow(0).getLastCellNum();
      }
    }
    return noOfColumns;
  }

  // Função para verificar se uma String é numérico
  private static boolean isNumeric(String strNum) {
    if (strNum == null) {
      return false;
    }
    try {
      double d = Double.parseDouble(strNum);
    } catch (NumberFormatException nfe) {
      return false;
    }
    return true;
  }

  private static boolean validateUrl(String url) throws IOException {
    try {
      FileInputStream fis = new FileInputStream(new java.io.File(url));
      XSSFWorkbook excel = new XSSFWorkbook(fis);
      fis.close();
      excel.close();
      return true;
    } catch (IOException e) {
      return false;
    }
  }

  private static int emptyRowValidation(Sheet sheet) {
    int aux = 0;
    int rowStart = sheet.getFirstRowNum();
    int rowEnd = sheet.getLastRowNum();
    for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
      Row r = sheet.getRow(rowNum);
      if (r == null) {
        break;
      }
      if (r != null) {
        aux = r.getRowNum();
      }
    }
    return aux;
  }

  private static boolean emptyCellValidation(Object obj) {
    if (obj == null) {
      return true;
    }
    return false;
  }

  private static Integer varcharSize(ArrayList<String> list) {
    Integer vrfCellSize = 0;
    Integer maxValue = 0;
    for (int i = 0; i < list.size(); i++) {
      if (emptyCellValidation(list.get(i)) == false) {
        if (list.get(i).length() > maxValue) {
          maxValue = list.get(i).length();
        }
      }
    }
    if (maxValue > 0 && maxValue < 4000) {
      vrfCellSize = maxValue * 3;
    } else {
      vrfCellSize = 4000;
    }
    return vrfCellSize;
  }

  private static ArrayList<String> regexStringValidation(ArrayList<String> list, int index, String path)
      throws IOException {
    Pattern spc = Pattern.compile("[^a-zA-Z0-9]", Pattern.CASE_INSENSITIVE);
    Pattern space = Pattern.compile(" ", Pattern.CASE_INSENSITIVE);
    Pattern kb = Pattern.compile("/", Pattern.CASE_INSENSITIVE);

    int noOfColumns = sweppLeafAllCollumns(index, path);
    // Lista com os valores das colunas para criar a tabela
    ArrayList<String> columns = new ArrayList<>();

    String x = "";
    for (int i = 0; i < noOfColumns; i++) {
      Matcher matcher = spc.matcher(list.get(i));
      Matcher matcher2 = space.matcher(list.get(i));
      Matcher matcher3 = kb.matcher(list.get(i));

      boolean find = matcher.find();
      boolean find2 = matcher2.find();
      boolean find3 = matcher3.find();

      if (find) {
        x = list.get(i).replaceAll("[^a-zA-Z0-9]", "_");

      } else if (find2) {
        x = list.get(i).replaceAll(" ", "__");
      } else if (find3) {
        x = list.get(i).replaceAll("/", "_");
      } else {
        x = list.get(i);
      }
      columns.add(x);
    }
    return columns;
  }

  //// Database connection Testes
  private static void creatConnection(Connection conn, Statement stat, String url, String sqlCreate, String sqlInsert)
      throws SQLException {
    conn = DriverManager.getConnection(url);
    stat = conn.createStatement();
    try {
      stat.executeQuery(sqlCreate.toString());
    } catch (Exception e) {
      e.printStackTrace();
    }

    try {
      stat.executeQuery(sqlInsert.toString());
    } catch (PSQLException e) {
      e.printStackTrace();
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      stat.close();
    }
  }

  private static boolean checkIfDateIsValid(String date) {
    SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
    format.setLenient(false);
    try {
      format.parse(date);
    } catch (ParseException e) {
      return false;
    }
    return true;
  }
}