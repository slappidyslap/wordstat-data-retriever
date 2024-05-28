//Ежедневные лимиты:
//
//        Лимит запросов: 100 000 запросов в день.
//        Скоростные лимиты:
//
//        Лимит запросов в секунду: 100 запросов в секунду на проект.
//        Лимит запросов в секунду на пользователя: 60 запросов в секунду на пользователя.
//
//С самого начала таймаут через 5 мин, потом 15 секунд на нахождения элемента таблицы

package kg.seo.musabaev;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;
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
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class Main {

    final static String spreadsheetId = "1sFcoLNCxFA4imDPjceSdQCUeosrL0uxeDmGfAim9DjU";
    final static String range = "Sheet34!B3:B1576";
    final static String[] letters = "CDEFGHI".split("");
    final static String WORDSTAT_BASE_URL = "https://wordstat.yandex.ru/?region=207&view=graph&words=";
    final static Sheets sheets;

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
                        ValueRange body = new ValueRange().setValues(List.of(Collections.nCopies(7, "'0")));
                        sheets.spreadsheets().values()
                                .update(spreadsheetId, "Sheet34!C" + (i + 3) + ":I" + (i + 3), body)
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
                            .update(spreadsheetId, "Sheet34!" + letters[j] + (i + 3), body)
                            .setValueInputOption("RAW")
                            .execute();
                }
            }
            TimeUnit.SECONDS.sleep(5);
        }
        web.close();
    }
}
