mod spreadsheet;

use anyhow::{anyhow, Result};
use axum::extract::multipart::Field;
use axum::http::header;
use std::fs::File;
use std::io::{Cursor, Write};

use axum::extract::Multipart;
use axum::response::{Html, IntoResponse};
use axum::routing::{get, post};
use axum::Router;

async fn upload_page() -> Html<&'static str> {
    Html(include_str!("../templates/upload.html"))
}

async fn process(mut multipart: Multipart) -> impl IntoResponse {
    println!("Received a request to process a file");

    // Extract the XLSX file (hopefully)
    if let Ok(Some(field)) = multipart.next_field().await {
        let name = field.file_name().unwrap_or_default().to_string();
        let data = field.bytes().await.unwrap_or_default();

        let cursor = Cursor::new(data);

        let res = spreadsheet::process_varuvo_export(cursor).unwrap();

        let headers = [
            (
                header::CONTENT_TYPE,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".to_owned(),
            ),
            (
                header::CONTENT_DISPOSITION,
                format!("attachment; filename=\"new-{}\"", name),
            ),
        ];

        Ok((headers, res))
    } else {
        Err("No file upload found")
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
