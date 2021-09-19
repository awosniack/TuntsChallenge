/*
    *   André Bernardo Wosniack - 19/09/2021
*/


import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.util.logging.Logger;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;

public class TuntsChallenge {


    private static Logger logger;
    private static Sheets sheetsService;

    private static String APPLICATION_NAME = "Desafio estágio Tunts";
    private static String SPREADSHEET_ID = "1qXksRkrkK1CRrwUNEnatiWzxr05IZR79r9LZcGa9VP4";

    // Creating OAuth exchange to grant access to google sheets
    private static Credential authorize() throws IOException, GeneralSecurityException {

        logger.info("Starting OAuth authorization");

        InputStream in = TuntsChallenge.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
            JacksonFactory.getDefaultInstance(), new InputStreamReader(in)
        );
        
        // Scopes to get access
        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
            GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
            clientSecrets, scopes)
            .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
            .setAccessType("offline")
            .build();
        Credential credential = new AuthorizationCodeInstalledApp(
            flow, new LocalServerReceiver())
            .authorize("user");
        return credential;
    }

    /*
        *   Returns Google Spreadsheet Service
    */
    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        logger.info("Starting SheetsService");
        Credential credential = authorize();
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
            JacksonFactory.getDefaultInstance(), credential)
            .setApplicationName(APPLICATION_NAME)
            .build();
    }

    /*
        *   Converts 3 Cell Objects into floats, calculates their mean and ceiling, and returns the result as integer.
    */
    private static Integer mean(Object n1,Object n2,Object n3) {
        float p1 = Float.parseFloat(n1.toString());
        float p2 = Float.parseFloat(n2.toString());
        float p3 = Float.parseFloat(n3.toString());
        return (int) Math.ceil((p1+p2+p3)/3);
    }

    /*
        *   Returns the Situation of a student based on their mean
    */
    private static String situation(int media){
        return media < 50 ? "Reprovado por Nota" : media < 70 ? "Exame final" : "Aprovado";
    }


    /*
        *   Returns the minimun score required for a student on their final exam
    */
    private static Integer calcNaf(int mean) {
        return (100 - mean);
    }

    /*
        *   Updates the spreadsheet with students situations and "Nota para aprovação Final"
    */
    private static void updateSheet(ArrayList<String> situations, ArrayList<Integer> naf) throws IOException, GeneralSecurityException {
        logger.info("Creating update request body");
        List<List<Object>> values = Arrays.asList(
            Arrays.asList(
                situations.toArray()
            ),
            Arrays.asList(
                naf.toArray()
            )
        );

        List<ValueRange> data = new ArrayList<>();
        data.add(new ValueRange()
            .setRange("G4:H27")
            .setValues(values)
            .setMajorDimension("COLUMNS"));

        BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
            .setValueInputOption("RAW")
            .setData(data);
                
        logger.info("Sending update request");
        BatchUpdateValuesResponse result = sheetsService.spreadsheets().values().batchUpdate(SPREADSHEET_ID, body).execute();
        logger.info(String.format("Updated %d cells", result.getTotalUpdatedCells()));

    }

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        
        // initializing log
        logger = Logger.getLogger(TuntsChallenge.class.getName());
        
        // initializing googleSheets service
        sheetsService = getSheetsService();

        // range to read from the spreadsheet
        String range = "B4:F27";

        ValueRange response = sheetsService.spreadsheets().values()
            .get(SPREADSHEET_ID, range)
            .execute();

        List<List<Object>> values = response.getValues();

        if (values == null || values.isEmpty()) {
            logger.warning("No data found");
        } else {
            ArrayList<String> situations = new ArrayList<>();
            ArrayList<Integer> naf = new ArrayList<>();

            for (List<Object> row : values) {
                Integer mean = mean(row.get(2), row.get(3), row.get(4));
                String currentSituation = situation(mean);
                int currentNaf = currentSituation.equals("Exame final") ? calcNaf(mean) : 0;


                if (Integer.parseInt(row.get(1).toString()) > 15) {
                    currentSituation = "Reprovado por Falta";
                    currentNaf = 0;
                }


                /*
                System.out.printf("Aluno(a) %s ficou com média %d, situação %s", row.get(0), mean, currentSituation);
                if (currentNaf != 0) {
                    System.out.printf(" precisando de %d para ser aprovado", currentNaf);
                }
                System.out.println();
                */
                
                naf.add(currentNaf); 
                situations.add(currentSituation);
            }

            updateSheet(situations, naf);
        }

    }

}