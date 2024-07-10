import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class GoogleSheetLibrary implements GoogleSheetsService {

    private static final String APPLICATION_NAME = "Google Sheets API Library";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private final Sheets service;

    public GoogleSheetLibrary(String credentialsFilePath, String tokensDirectoryPath) throws GeneralSecurityException, IOException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT, credentialsFilePath, tokensDirectoryPath))
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT, String credentialsFilePath, String tokensDirectoryPath) throws IOException {
        FileInputStream in = new FileInputStream(credentialsFilePath);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, Collections.singletonList(SheetsScopes.SPREADSHEETS))
                .setDataStoreFactory(new FileDataStoreFactory(new File(tokensDirectoryPath)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    @Override
    public void createSheet(String spreadsheetID,String sheetName) throws IOException {
        BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest().setRequests(
                Collections.singletonList(new Request().setAddSheet(new AddSheetRequest().setProperties(
                        new SheetProperties().setTitle(sheetName)))));
        service.spreadsheets().batchUpdate(spreadsheetID, request).execute();

    }

    @Override
    public void updateSheet(String spreadsheetId, String sheetName,  List<Object> value, String cell,String cell2) throws IOException {
        ValueRange body = new ValueRange().setValues(Collections.singletonList(value));
        service.spreadsheets().values().update(spreadsheetId, sheetName + "!" + cell+":"+cell2, body)
                .setValueInputOption("RAW")
                .execute();
    }

    @Override
    public void deleteSheet(String spreadsheetId, int sheetId) throws IOException {
        BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest().setRequests(
                Collections.singletonList(new Request().setDeleteSheet(new DeleteSheetRequest().setSheetId(sheetId))));
        service.spreadsheets().batchUpdate(spreadsheetId, request).execute();
    }

    @Override
    public List<List<Object>> readSheet(String spreadsheetId, String range) throws IOException {
        ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
        return response.getValues();
    }

    @Override
    public void appendToSheet(String spreadsheetId, String range, List<List<Object>> values) throws IOException {
        ValueRange body = new ValueRange().setValues(values);
        service.spreadsheets().values().append(spreadsheetId, range, body)
                .setValueInputOption("RAW")
                .execute();
    }

    @Override
    public String createSpreadsheet(String title) throws IOException {
        Spreadsheet spreadsheet = new Spreadsheet()
                .setProperties(new SpreadsheetProperties()
                        .setTitle(title));
        Spreadsheet response = service.spreadsheets().create(spreadsheet).execute();
        return response.getSpreadsheetId();
    }

//    @Override
//    public void insertRows(String spreadsheetId, String range, List<List<Object>> values) throws IOException {
//        ValueRange body = new ValueRange().setValues(values);
//        service.spreadsheets().values().update(spreadsheetId, range, body)
//                .setValueInputOption("RAW")
//                .execute();
//    }

    @Override
    public void clearRange(String spreadsheetId, String range) throws IOException {
        ClearValuesRequest requestBody = new ClearValuesRequest();
        service.spreadsheets().values().clear(spreadsheetId, range, requestBody).execute();
    }

    @Override
    public void findAndReplace(String spreadsheetId, int sheetId, String find, String replace) throws IOException {
        FindReplaceRequest findReplaceRequest = new FindReplaceRequest()
                .setFind(find)
                .setReplacement(replace)
                .setSheetId(sheetId);

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(Collections.singletonList(new Request().setFindReplace(findReplaceRequest)));
        service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
    }

    @Override
    public void batchUpdateValues(String spreadsheetId, List<List<Object>> data,String sheetName) throws IOException {
        ClearValuesRequest clearValuesRequest = new ClearValuesRequest();
        service.spreadsheets().values().clear(spreadsheetId, sheetName, clearValuesRequest).execute();

        // Define the range to update (starting from A1)
        String updateRange = sheetName + "!A1";

        // Create the ValueRange object
        ValueRange valueRange = new ValueRange()
                .setRange(updateRange)
                .setValues(data);

        // Create a BatchUpdateValuesRequest
        BatchUpdateValuesRequest batchUpdateRequest = new BatchUpdateValuesRequest()
                .setValueInputOption("RAW")
                .setData(Collections.singletonList(valueRange));

        service.spreadsheets().values().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
    }

    @Override
    public Spreadsheet getSpreadsheetInfo(String spreadsheetId) throws IOException {
        return service.spreadsheets().get(spreadsheetId).execute();
    }

    private int getSheetId(String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = getSpreadsheetInfo(spreadsheetId);
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IOException("Sheet not found: " + sheetName);
    }
    public void insertRow(String spreadsheetId, int sheetId, int startIndex, boolean insertAfter) throws IOException {
        int index = insertAfter ? startIndex + 1 : startIndex;
        DimensionRange dimensionRange = new DimensionRange()
                .setSheetId(sheetId)
                .setDimension("ROWS")
                .setStartIndex(index)
                .setEndIndex(index + 1);

        InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest()
                .setRange(dimensionRange)
                .setInheritFromBefore(index != 0);

        Request request = new Request().setInsertDimension(insertDimensionRequest);
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(Arrays.asList(request));

        service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    }

    public void insertColumn(String spreadsheetId, int sheetId, int startIndex, boolean insertAfter) throws IOException {
        int index = insertAfter ? startIndex + 1 : startIndex;
        DimensionRange dimensionRange = new DimensionRange()
                .setSheetId(sheetId)
                .setDimension("COLUMNS")
                .setStartIndex(index)
                .setEndIndex(index + 1);

        InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest()
                .setRange(dimensionRange)
                .setInheritFromBefore(index != 0);

        Request request = new Request().setInsertDimension(insertDimensionRequest);
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(Arrays.asList(request));

        service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    }
}
