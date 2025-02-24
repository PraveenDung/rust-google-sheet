use dotenv::dotenv;
use jsonwebtoken::{encode, Algorithm, EncodingKey, Header};
use reqwest::Client;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::env;
use std::fs::File;
use std::io::Write;
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Serialize, Deserialize)]
struct Claims {
    iss: String,   // Service account email
    scope: String, // Google Sheets API scope
    aud: String,   // Token URL
    exp: u64,      // Expiration time
    iat: u64,      // Issued at time
}

// Function to get Google OAuth2 token
async fn get_google_access_token() -> Result<String, Box<dyn std::error::Error>> {
    dotenv().ok(); // Load .env variables

    let client_email = env::var("SERVICE_ACCOUNT_EMAIL")?;
    let private_key = env::var("PRIVATE_KEY")?.replace("\\n", "\n"); // Convert escaped \n to actual newlines

    let now = SystemTime::now().duration_since(UNIX_EPOCH)?.as_secs();
    let claims = Claims {
        iss: client_email,
        scope: "https://www.googleapis.com/auth/spreadsheets".to_string(), // Full access needed to write
        aud: "https://oauth2.googleapis.com/token".to_string(),
        exp: now + 3600,
        iat: now,
    };

    let jwt = encode(
        &Header::new(Algorithm::RS256),
        &claims,
        &EncodingKey::from_rsa_pem(private_key.as_bytes())?,
    )?;

    let client = Client::new();
    let response = client
        .post("https://oauth2.googleapis.com/token")
        .form(&[
            ("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer"),
            ("assertion", &jwt),
        ])
        .send()
        .await?
        .json::<Value>()
        .await?;

    Ok(response["access_token"].as_str().unwrap().to_string())
}

// Function to read Google Sheets data
async fn read_google_sheet(
    access_token: &str, column_index1: usize, filter_value1: &str, column_index2: usize, filter_value2: &str) -> Result<(), Box<dyn std::error::Error>> {
    dotenv().ok();
    let sheet_id = env::var("SHEET_ID")?;
    let range = "RETURNS MAIN"; // Reads entire sheet

    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
        sheet_id, range
    );

    let client = Client::new();
    let response = client
        .get(&url)
        .bearer_auth(access_token)
        .send()
        .await?
        .json::<Value>()
        .await?;

    print!("{}", response);

    let mut filtered_data = Vec::new();
    let mut count = 0;
    if let Some(values) = response["values"].as_array() {
        // println!(
        //     "Filtered Rows where Column {} = '{}':",
        //     column_index + 1,
        //     filter_value
        // );

        // ✅ Print & Store Header Row
        let header = &values[0];
        println!("📌 Header: {:?}", header);
        for row in values.iter().skip(1) {
            let match_col1 = row.get(column_index1).map_or(false, |cell| cell.as_str() == Some(filter_value1));
            let match_col2 = row.get(column_index2).map_or(false, |cell| cell.as_str() == Some(filter_value2));

            if match_col1 && match_col2 {
                println!("{:?}", row);
                filtered_data.push(row.clone());
                count += 1;
            }
        }
        println!("Total Matching Rows: {}", count);
        // ✅ Save to JSON file
        let json_output = json!({
            "header": header,
            "filtered_data": filtered_data,
            "count": count
        });

        let mut file = File::create("output.json")?;
        file.write_all(json_output.to_string().as_bytes())?;
        println!("✅ Data saved to 'output.json'");
    } else {
        println!("No data found!");
    }

    Ok(())
}

// Function to append a row to Google Sheets
async fn append_row_to_google_sheet(
    access_token: &str,
    new_row: Vec<String>,
) -> Result<(), Box<dyn std::error::Error>> {
    dotenv().ok();
    let sheet_id = env::var("SHEET_ID")?;
    let range = "Sheet1"; // Adjust based on sheet name

    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}:append?valueInputOption=RAW",
        sheet_id, range
    );

    let client = Client::new();
    let body = serde_json::json!({
        "values": [new_row] // Data to be inserted
    });

    let response = client
        .post(&url)
        .bearer_auth(access_token)
        .json(&body)
        .send()
        .await?
        .json::<Value>()
        .await?;

    println!("✅ Row added: {:#?}", response);
    Ok(())
}

// Function to update a specific row
async fn update_row_in_google_sheet(
    access_token: &str,
    row_index: usize,
    values: Vec<String>,
) -> Result<(), Box<dyn std::error::Error>> {
    dotenv().ok();
    let sheet_id = env::var("SHEET_ID")?;
    let range = format!("Sheet1!A{}:Z{}", row_index, row_index); // Adjust based on column range

    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}?valueInputOption=RAW",
        sheet_id, range
    );

    let body = serde_json::json!({
        "values": [values]
    });

    let client = Client::new();
    let response = client
        .put(&url)
        .bearer_auth(access_token)
        .json(&body)
        .send()
        .await?;

    println!("Update row status: {}", response.status());
    Ok(())
}

// Function to delete a row from Google Sheets
async fn delete_row_from_google_sheet(
    access_token: &str,
    row_index: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    dotenv().ok();
    let sheet_id = env::var("SHEET_ID")?;

    let delete_url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}:batchUpdate",
        sheet_id
    );

    let body = serde_json::json!({
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": 0, // Sheet ID (0 usually refers to the first sheet)
                        "dimension": "ROWS",
                        "startIndex": row_index - 1,  // Google Sheets uses zero-based index
                        "endIndex": row_index
                    }
                }
            }
        ]
    });

    let client = Client::new();
    let response = client
        .post(&delete_url)
        .bearer_auth(access_token)
        .json(&body)
        .send()
        .await?;

    println!("Delete row status: {}", response.status());
    Ok(())
}

#[tokio::main]
async fn main() {
    match get_google_access_token().await {
        Ok(token) => {
            println!("🔑 Token retrieved!");

            let column_index1: usize = 1; // Column B (CHANNEL VLOOKUP)
            let filter_value1 = "DEBENHAMS";

            let column_index2: usize = 9; // Column J (Refunded)
            let filter_value2 = "FALSE"; 

            // Read existing data
            if let Err(e) = read_google_sheet(&token, column_index1, filter_value1, column_index2, filter_value2).await {
                eprintln!("Error reading sheet: {}", e);
            }

            // // Append a new row
            // let new_data = vec![
            //     "SaveEfforts".to_string(),
            //     "saveefforts@gmail.com".to_string(),
            //     "192783568".to_string(),
            // ];
            // if let Err(e) = append_row_to_google_sheet(&token, new_data).await {
            //     eprintln!("Error appending row: {}", e);
            // }

            // // Update row 2 (change values as needed)
            // let updated_row = vec!["Jane Doe".to_string(), "janedoe@gmail.com".to_string(), "1234567890".to_string()];
            // if let Err(e) = update_row_in_google_sheet(&token, 2, updated_row).await {
            //     eprintln!("Error updating row: {}", e);
            // }

            // // Delete a row (Example: Delete Row 3)
            // if let Err(e) = delete_row_from_google_sheet(&token, 3).await {
            //     eprintln!("Error deleting row: {}", e);
            // }
        }
        Err(e) => eprintln!("Error getting token: {}", e),
    }
}
