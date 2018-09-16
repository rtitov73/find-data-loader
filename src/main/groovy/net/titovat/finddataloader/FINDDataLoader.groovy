package net.titovat.finddataloader

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFRichTextString

class FINDDataLoader {

  private static final int FIRST_ROW = 1
  private static final String[] SPECIES = ["", "P. falciparum only", "P. falciparum + other species", "PAN only", "P. vivax only"]
  private static final String[] FORMAT = ["Hybrid", "Cassette", "Dipstick", "Card"]

  private static final String BOILERPLATE = '''
    This is a companion program for loading new rounds of testing data into the FIND
    Interactive guide for high-quality malaria RDT selection. This companion program takes
    an Excel file in a predefined format and converts it into a set of SQL insert statements
    that can be executed against the database in order to import the new data.
  '''

  private static String columnFootnote = 'null'

  static void main(String[] args) {
    String inputFile = System.getProperty('input')
    String outputFile = System.getProperty('output')
    String tableName = System.getProperty('tableName')

    if (!tableName) {
      tableName = 'malaria_rdt_tests_rnd_1_8'
    }

    println BOILERPLATE

    if (!inputFile || !outputFile) {
      println '''
        Usage: gradle run -Dinput=<input_file> -Doutput=<output_file> [-DtableName=<tableName>]
        
        Example: gradle run -Dinput="f:\\R8.xlsx" -Doutput="f:\\output.sql" -DtableName=malaria_rdt_tests_rnd_1_8
      '''
      System.exit(1)
    }

    Sheet sheet = loadWorksheet(inputFile)
    File output = new File(outputFile)
    Writer writer = output.newWriter()

//    output << "\n-- Identify products which exist in the current table but don't exist in the new spreadsheet\n\n"
//    output << prepareVerificationSql(sheet, tableName)
    writer << prepareInsertSqls(sheet, tableName)
    writer.close()
  }

  static String prepareVerificationSql(Sheet sheet, String tableName) {
    List<String> productIds = new ArrayList<>()

    for (Row row = sheet.getRow(FIRST_ROW); row?.getCell(0) != null; row = sheet.getRow(row.getRowNum() + 1)) {
      productIds.add("'" + getStringValue(row, 0) + "'")
    }
    return """Select product_id, product, manufacturer from ${tableName} where product_id not in (
        ${productIds.join(',')}
    ) order by product_id;
    """
  }

  static String prepareInsertSqls(Sheet sheet, String tableName) {
    int rowIndex = FIRST_ROW
    StringBuilder result = new StringBuilder()

    result << "-- Insert statements\n\n"

    while (true) {
      Row row = sheet.getRow(rowIndex++)
      StringBuilder s = new StringBuilder()

      if (row?.getCell(0) == null)
        break

      columnFootnote = 'null'

      sqlHeader s, tableName
      str s, row, 0
      str s, row, 1
      str s, row, 2
      str s, row, 3
      index s, SPECIES, row, 4
      index s, FORMAT, row, 5
      inum s, row, 6
      d3 s, row, 7
      d3 s, row, 8
      d3 s, row, 9
      d3 s, row, 10
      d3 s, row, 11
      d3 s, row, 12
      d3 s, row, 13
      d3 s, row, 14
      d3 s, row, 15
      num s, row, 16      // Invalid rate
      d3 s, row, 17
      d3 s, row, 18
      d3 s, row, 19
      d3 s, row, 20
      d3 s, row, 21
      d3 s, row, 22
      d3 s, row, 23
      d3 s, row, 24
      d3 s, row, 25
      d3 s, row, 26
      d3 s, row, 27
      d3 s, row, 28
      d3 s, row, 29
      d3 s, row, 30
      d3 s, row, 31
      d3 s, row, 32
      d3 s, row, 33
      d3 s, row, 34
      d3 s, row, 35
      d3 s, row, 36
      d3 s, row, 37
      d3 s, row, 38
      d3 s, row, 39
      d3 s, row, 40
      nstr s, row, 41
      nstr s, row, 42
      nstr s, row, 43   // Test line 3
      inum s, row, 44
      inum s, row, 45
      dashnum s, row, 46
      dashnum s, row, 47
      inum s, row, 48
      s.append("${columnFootnote}, ")
      s.append(getValue(row, 49) == 'yes' ? "'Y'" : "'N'")
      s.append(", 'Y', ${rowIndex - 2});")

      result.append(shrinkSpaces(s.toString())).append("\n")
    }
    return result.toString()
  }

  static String shrinkSpaces(String s) {
    s = s.replaceAll('\n', ' ')
    10.times { s = s.replaceAll('  ', ' ') }
    return s.trim()
  }

  static void d3(StringBuilder s, Row row, int offset) {
    String[] data = getValue(row, offset).replaceAll(' ', '').split(/[\(\)\/]/)

    for (int i = 0; i < data.length; i++) {
      data[i] = fmt(data[i])

      if ((data[i] != 'null') && (!data[i].isNumber())) {
        columnFootnote = "'" + data[i].replaceAll("[^A-Za-z]","") + "'"
        data[i] = data[i].replaceAll("[A-Za-z]","")
      }
    }

    if (data.length == 1)
      s.append(data[0]).append(', null, null, ')
    else if (data.length == 2)
      s.append(data[0]).append(', ').append(data[1]).append(', null, ')
    else if (data.length == 3)
      s.append(data[0]).append(', ').append(data[1]).append(', ').append(data[2]).append(', ')
  }

  static String fmt(String data) {
    return (data == 'NA') ? '-999' : (data == 'ND') ? 'null' : data;
  }

  static void nstr(StringBuilder s, Row row, int offset) {
    String val = getValue(row, offset)
    s.append((val == 'NA') ? 'null, ' : "'" + val + "', ")
  }

  static void dashnum(StringBuilder s, Row row, int offset) {
    String val = getValue(row, offset)
    s.append((val == '-') ? '0, ' : (row.getCell(offset).getNumericCellValue() as Integer).toString() + ', ')
  }

  static void str(StringBuilder s, Row row, int offset) {
    s.append("'" + getStringValue(row, offset) + "', ")
  }

  static String getStringValue(Row row, int offset) {
    Cell cell = row.getCell(offset)

    if (cell.getCellTypeEnum() == CellType.STRING) {
      RichTextString richText = (XSSFRichTextString)cell.getRichStringCellValue()

      if (richText.hasFormatting())
      {
        int plainTextLength = richText.getLengthOfFormattingRun(0)
        int superscriptLength = richText.getLengthOfFormattingRun(1)

        String plainText = richText.toString().substring(0, plainTextLength)
        String superscript = richText.toString().substring(plainTextLength, plainTextLength + superscriptLength)

        return plainText + '<sup>' + superscript + '</sup>'
      } else {
        return cell.getStringCellValue()
      }
    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
      return (cell.getNumericCellValue() as Long).toString()
    } else {
      return ''
    }
  }

  static double getNumericValue(Row row, int offset) {
    Cell cell = row.getCell(offset)

    if (cell.getCellTypeEnum() == CellType.STRING) {
      return Double.parseDouble(cell.getStringCellValue())
    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
      return cell.getNumericCellValue()
    }
  }

  static void num(StringBuilder s, Row row, int offset) {
    s.append(getValue(row, offset) + ", ")
  }

  static void inum(StringBuilder s, Row row, int offset) {
    s.append((getNumericValue(row, offset) as Integer) + ", ")
  }

  static void index(StringBuilder s, String[] ref,  Row row, int offset) {
    s.append(ref.findIndexOf { it == getValue(row, offset) }).append(', ')
  }

  static void sqlHeader(StringBuilder s, String tableName) {
    s.append """
      INSERT into ${tableName} (
      product_id, product, manufacturer, test_type, plasmodium_species, format, round_id,
      detection_rate_200_pf, detection_rate_200_pf_1, detection_rate_200_pf_2,
      detection_rate_200_pv, detection_rate_200_pv_1, detection_rate_200_pv_2,
      detection_rate_5000_pf, detection_rate_5000_pf_1, detection_rate_5000_pf_2,
      detection_rate_5000_pv, detection_rate_5000_pv_1, detection_rate_5000_pv_2,
      false_positive_rate_200_pf_non_pf, false_positive_rate_200_pf_non_pf_1, false_positive_rate_200_pf_non_pf_2,
      false_positive_rate_5000_pf_non_pf, false_positive_rate_5000_pf_non_pf_1, false_positive_rate_5000_pf_non_pf_2,
      false_positive_rate_200_pv_pf, false_positive_rate_200_pv_pf_1, false_positive_rate_200_pv_pf_2,
      false_positive_rate_5000_pv_pf, false_positive_rate_5000_pv_pf_1, false_positive_rate_5000_pv_pf_2,
      total_false_positive_rate, total_false_positive_rate_1, total_false_positive_rate_2,
      invalid_rate,
      heat_stability_200_pf_baseline, heat_stability_200_pf_baseline_1, heat_stability_200_pf_baseline_2,
      heat_stability_200_pf_35, heat_stability_200_pf_35_1, heat_stability_200_pf_35_2,
      heat_stability_200_pf_45, heat_stability_200_pf_45_1, heat_stability_200_pf_45_2,
      heat_stability_200_pan_baseline, heat_stability_200_pan_baseline_1, heat_stability_200_pan_baseline_2,
      heat_stability_200_pan_35, heat_stability_200_pan_35_1, heat_stability_200_pan_35_2,
      heat_stability_200_pan_45, heat_stability_200_pan_45_1, heat_stability_200_pan_45_2,
      heat_stability_2000_pf_baseline, heat_stability_2000_pf_baseline_1, heat_stability_2000_pf_baseline_2,
      heat_stability_2000_pf_35, heat_stability_2000_pf_35_1, heat_stability_2000_pf_35_2,
      heat_stability_2000_pf_45, heat_stability_2000_pf_45_1, heat_stability_2000_pf_45_2,
      heat_stability_2000_pan_baseline, heat_stability_2000_pan_baseline_1, heat_stability_2000_pan_baseline_2,
      heat_stability_2000_pan_35, heat_stability_2000_pan_35_1, heat_stability_2000_pan_35_2,
      heat_stability_2000_pan_45, heat_stability_2000_pan_45_1, heat_stability_2000_pan_45_2,
      heat_stability_200_pv_baseline, heat_stability_200_pv_baseline_1, heat_stability_200_pv_baseline_2,
      heat_stability_200_pv_35, heat_stability_200_pv_35_1, heat_stability_200_pv_35_2,
      heat_stability_200_pv_45, heat_stability_200_pv_45_1, heat_stability_200_pv_45_2,
      heat_stability_2000_pv_baseline, heat_stability_2000_pv_baseline_1, heat_stability_2000_pv_baseline_2,
      heat_stability_2000_pv_35, heat_stability_2000_pv_35_1, heat_stability_2000_pv_35_2,
      heat_stability_2000_pv_45, heat_stability_2000_pv_45_1, heat_stability_2000_pv_45_2,
      heat_stability_200_pv_pan_baseline, heat_stability_200_pv_pan_baseline_1, heat_stability_200_pv_pan_baseline_2,
      heat_stability_200_pv_pan_35, heat_stability_200_pv_pan_35_1, heat_stability_200_pv_pan_35_2,
      heat_stability_200_pv_pan_45, heat_stability_200_pv_pan_45_1, heat_stability_200_pv_pan_45_2,
      heat_stability_2000_pv_pan_baseline, heat_stability_2000_pv_pan_baseline_1, heat_stability_2000_pv_pan_baseline_2,
      heat_stability_2000_pv_pan_35, heat_stability_2000_pv_pan_35_1, heat_stability_2000_pv_pan_35_2,
      heat_stability_2000_pv_pan_45, heat_stability_2000_pv_pan_45_1, heat_stability_2000_pv_pan_45_2,
      test_line_1, test_line_2, test_line_3, blood_volume, buffer, min_read_time, max_read_time, protocol_group,
      column_footnote,
      who_qualified, is_visible, id) values (
    """
  }

  static String getValue(Row row, int offset) {
    return getValue(row.getCell(offset))
  }

  static String getValue(Cell cell) {
    String s = "???"
    if (cell.getCellTypeEnum() == CellType.STRING) {
      s = cell.getStringCellValue()
    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
      s = cell.getNumericCellValue()
    }
    return s
  }

  static Sheet loadWorksheet(String fileName) {
    return WorkbookFactory.create(new File(fileName)).getSheetAt(0)
  }

}
