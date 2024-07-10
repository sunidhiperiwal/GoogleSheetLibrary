import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class TestGoogleSheet {


  public static void main(String args[]) throws IOException, GeneralSecurityException {
    System.out.println(System.getProperty("user.dir"));
    String credentialsFilePath = "C:/Users/Sunidhi Periwal/IdeaProjects/TestGoogleSheet/src/main/resources/credentials.json";
    String tokenDirectoryPath = System.getProperty("user.dir") + "/tokens/path";
    String id = "1p1XbBYkLWdCZW_DTZe-VwYt7yYkTbJ7r-CQ-omgoBfM";
    GoogleSheetsService googleSheetsService = new GoogleSheetLibrary(credentialsFilePath, tokenDirectoryPath);
    googleSheetsService.createSheet(id, "abc");
    List<Object> newData = List.of(789, 387.5);
    googleSheetsService.updateSheet(id, "Sheet1", newData, "F4", "G4");
    googleSheetsService.deleteSheet(id, 1448697540);
  List<List<Object>> datas=googleSheetsService.readSheet(id,"Sheet1!A1:G7");
   for (List row : datas) {
      System.out.printf("%s, %s, %s, %s, %s, %s, %s\n", row.get(0), row.get(1), row.get(2), row.get(3), row.get(4), row.get(5), row.get(6));
    }
   List<List<Object>> values = Arrays.asList(
           Arrays.asList("2/14/2021", "Cents", "Sisss", "Pint", 124, 20.01, 1000.6));
googleSheetsService.appendToSheet(id,"Sheet1!A8:G8",values);
    googleSheetsService.createSpreadsheet("abrad");
      googleSheetsService.insertRow(id, 0, 6, false);
      googleSheetsService.insertColumn(id, 0, 0, true);
    googleSheetsService.clearRange(id,"Sheet1!F3:G3");
    googleSheetsService.findAndReplace(id,0,"19.99","12");

    googleSheetsService.batchUpdateValues(id,values,"Shot");
  }
}
