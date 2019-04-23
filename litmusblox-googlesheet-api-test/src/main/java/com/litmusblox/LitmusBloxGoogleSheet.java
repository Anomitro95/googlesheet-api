package com.litmusblox;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;

import java.text.ParseException;

public class LitmusBloxGoogleSheet {

	/** Application name. */
	private static final String APPLICATION_NAME = "Google Sheets API LitmusBlox";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"),
			".credentials/sheets.googleapis.com-java-quickstart");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	private static double totalFixedCost;
	private static long totalActualHours;

	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public static Credential authorize() throws IOException {
// Load client secrets.
		InputStream in = LitmusBloxGoogleSheet.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES).setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * 
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public static Sheets getSheetsService() throws IOException {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(APPLICATION_NAME)
				.build();
	}

	public static void main(String[] args) throws IOException, ParseException {
// Build a new authorized API client service.
		Sheets service = getSheetsService();

// Fetching data from the given googlesheet
// https://docs.google.com/spreadsheets/d/1irzwIzOdKXqw94fmz21glxm4f7YJyj8wSLX1hiFsQio/edit?ts=5cbbfa4d#gid=0
		String spreadsheetId = "1irzwIzOdKXqw94fmz21glxm4f7YJyj8wSLX1hiFsQio";
		String range = "Sheet1!A1:K6";
		ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
		List<List<Object>> values = response.getValues();
		if (values == null || values.size() == 0) {
			System.out.println("No data found.");
		} else {

			for (List row : values) {

				System.out.printf("%20s%20s%20s%20s%20s\n", row.get(0), row.get(1), row.get(2), row.get(8),
						row.get(10));

			}
			System.out.println("\n");
			String rangeData = "Sheet1!A2:K6"; // modified range to fetch the data(number) cells of the googlesheet
			ValueRange responseData = service.spreadsheets().values().get(spreadsheetId, rangeData).execute();
			List<List<Object>> valuesData = responseData.getValues();

			for (List row : valuesData) {

				String fixedCost = ((String) row.get(8)).substring(1).trim(); // removing the dollar sign from fixed cost
				String actualHours = ((String) row.get(10)).trim();
				fixedCost = fixedCost.replace(",", "");		//extracting the cost as number in string format
				Double cost = Double.parseDouble(fixedCost);
				Double hours = Double.parseDouble(actualHours);
				totalFixedCost += cost;
				totalActualHours += hours;
			}
			System.out.println("The total fixed Cost is : $" + totalFixedCost);

			System.out.println("The total actual hours is : " + totalActualHours);

		}
	}
}