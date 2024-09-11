mod spreadsheet;

use axum::http::{header, StatusCode};
use std::io::Cursor;

use axum::extract::Multipart;
use axum::response::{Html, IntoResponse};
use axum::routing::{get, post};
use axum::Router;
use chrono::prelude::*;

fn log(msg: String) {
    println!("{} - {}", Utc::now(), msg);
}

async fn upload_page() -> Html<&'static str> {
    Html(include_str!("../templates/upload.html"))
}

async fn process(mut multipart: Multipart) -> impl IntoResponse {
    log("received processing request".to_string());

    // Extract the XLSX file (hopefully)
    if let Ok(Some(field)) = multipart.next_field().await {
        let name = field.file_name().unwrap_or_default().to_string();
        let data = field.bytes().await.unwrap_or_default();

        log(format!("└─ processing {name}"));

        let cursor = Cursor::new(data);

        let res = match spreadsheet::process_varuvo_export(cursor) {
            Ok(res) => res,
            Err(e) => {
                log(format!("└─ failed during spreadsheet processing: {e}"));
                return Err((
                    StatusCode::INTERNAL_SERVER_ERROR,
                    "Er ging iets mis tijdens het bewerken van het Excel bestand",
                ));
            }
        };

        let headers = [
            (
                header::CONTENT_TYPE,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".to_owned(),
            ),
            (
                header::CONTENT_DISPOSITION,
                format!("attachment; filename=\"BEREKEND-{}\"", name),
            ),
        ];

        Ok((headers, res))
    } else {
        log("└─ no file found".to_string());
        Err((StatusCode::BAD_REQUEST, "No file upload found"))
    }
}

#[tokio::main]
async fn main() {
    let app = Router::new()
        .route("/", get(upload_page))
        .route("/process", post(process));

    let listener = tokio::net::TcpListener::bind("0.0.0.0:3000").await.unwrap();
    axum::serve(listener, app).await.unwrap();
}
