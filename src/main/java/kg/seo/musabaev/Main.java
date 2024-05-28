//Ежедневные лимиты:
//
//        Лимит запросов: 100 000 запросов в день.
//        Скоростные лимиты:
//
//        Лимит запросов в секунду: 100 запросов в секунду на проект.
//        Лимит запросов в секунду на пользователя: 60 запросов в секунду на пользователя.
//
//С самого начала таймаут через 5 мин, потом 15 секунд на нахождения элемента таблицы
//
//i + 2 потому что итерация начинается со строки 3

package kg.seo.musabaev;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static java.lang.Integer.parseInt;

public class Main {

    final static String spreadsheetId = "1sFcoLNCxFA4imDPjceSdQCUeosrL0uxeDmGfAim9DjU";
    final static String defaultSheetName = "Sheet34";
    final static String range = defaultSheetName + "!B2:B1576";
    final static String[] letters = "CDEFGHI".split("");
    final static String WORDSTAT_BASE_URL = "https://wordstat.yandex.ru/?region=207&view=graph&words=";
    final static Sheets sheets;
    final static int offset = 2;
    final static List<Request> reqs = new ArrayList<>();

    static {
        try {
            sheets = Google.getSheets();
        } catch (GeneralSecurityException | IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String... args) throws IOException, InterruptedException {

        WebDriverManager.chromedriver().setup();
        WebDriver web = new ChromeDriver();
        ValueRange response = sheets.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            for (int i = 0; i < 1576; i++) {
                final String keyword = (String) values.get(i).get(0);
                final String url = WORDSTAT_BASE_URL + keyword.replaceAll("\\s", "%20");

                web.get(url);
                List<WebElement> elements;
                if (i == 0) {
                    WebDriverWait wait = new WebDriverWait(web, Duration.ofMinutes(5));
                    elements = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.tagName("table")));
                } else {
                    try {
                        WebDriverWait wait = new WebDriverWait(web, Duration.ofSeconds(15));
                        elements = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.tagName("table")));
                    } catch (TimeoutException e) {
                        ValueRange body = new ValueRange().setValues(List.of(Collections.nCopies(7, 0)));
                        sheets.spreadsheets().values()
                                .update(spreadsheetId, defaultSheetName + "!C" + (i + offset) + ":I" + (i + offset), body)
                                .setValueInputOption("RAW")
                                .execute();
                        continue;
                    }
                }
                var list = Arrays.stream(elements.get(0).getText().split("\\n")).toList();
                var list2 = list.subList(22, list.size());
                System.out.println(keyword);
                for (int j = 0; j < list2.size(); j++) {
                    var a = list2.get(j).split("\\s");
                    String data;
                    if (a.length == 4) data = "%s".formatted(a[2]);
                    else data = "%s%s".formatted(a[2], a[3]);
                    System.out.println(data);
                    ValueRange body = new ValueRange().setValues(Collections.singletonList(
                            Collections.singletonList(data)));
                    sheets.spreadsheets().values()
                            .update(spreadsheetId, defaultSheetName + "!" + letters[j] + (i + offset), body)
                            .setValueInputOption("RAW")
                            .execute();
                    TimeUnit.SECONDS.sleep(1);
                }
                GridRange gridRange = new GridRange()
                        .setSheetId(1968414648)  // ID листа (листа по умолчанию 0)
                        .setStartRowIndex(i + offset)
                        .setEndRowIndex(i + offset + 1)
                        .setStartColumnIndex(8)
                        .setEndColumnIndex(9);
                CellFormat cellFormat = new CellFormat()
                        .setBackgroundColor(getColor(i));
                CellData cellData = new CellData().setUserEnteredFormat(cellFormat);
                RowData rowData = new RowData().setValues(Collections.singletonList(cellData));
                UpdateCellsRequest updateCellsRequest = new UpdateCellsRequest()
                        .setRange(gridRange)
                        .setRows(Collections.singletonList(rowData))
                        .setFields("userEnteredFormat.backgroundColor");
                reqs.add(new Request().setUpdateCells(updateCellsRequest));
                System.out.println(getColor(i).toPrettyString());
            }
        }
        web.close();

        System.out.println("начинаю закрашивать");
        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(reqs);
        sheets.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
    }

    private static Color getColor(int i) throws IOException {
        int octoberData = parseInt(((String) sheets.spreadsheets().values()
                .get(spreadsheetId, defaultSheetName + "!C" + (i + offset))
                .execute()
                .getValues()
                .get(0)
                .get(0))
                .replace(" ", ""));
        int aprilData = parseInt(((String) sheets.spreadsheets().values()
                .get(spreadsheetId, defaultSheetName + "!I" + (i + offset))
                .execute()
                .getValues()
                .get(0)
                .get(0))
                .replace(" ", ""));
        int difference = Math.abs(octoberData - aprilData);
        // Определение меньшего из двух чисел
        int smallerNumber = Math.min(octoberData, aprilData);
        // Вычисление процентной разницы
        double percentageDifference = (double) difference / smallerNumber * 100;
        if (percentageDifference > 15.0) return new Color().setRed(1f).setGreen(0f).setBlue(0f);
        return new Color().setRed(0f).setGreen(1f).setBlue(0f);
    }
}
