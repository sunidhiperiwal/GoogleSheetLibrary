import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.util.List;

public interface GoogleSheetsService {
    void createSheet(String spreadsheetId, String sheetName) throws IOException;
    void updateSheet(String spreadsheetId, String sheetName, String value, String cell) throws IOException;
    void deleteSheet(String spreadsheetId, int sheetId) throws IOException;
    List<List<Object>> readSheet(String spreadsheetId, String range) throws IOException;
    void appendToSheet(String spreadsheetId, String range, List<List<Object>> values) throws IOException;
    String createSpreadsheet(String title) throws IOException;

    void insertRows(String spreadsheetId, String range, List<List<Object>> values) throws IOException;
    void clearRange(String spreadsheetId, String range) throws IOException;
    void findAndReplace(String spreadsheetId, String sheetName, String find, String replace) throws IOException;
    void batchUpdateValues(String spreadsheetId, List<ValueRange> data) throws IOException;
    Spreadsheet getSpreadsheetInfo(String spreadsheetId) throws IOException;
}