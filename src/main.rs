use warp::{Filter, Reply};
use std::convert::Infallible;
use xlsxwriter::{DateTime as XLSDateTime, Format, Workbook, Worksheet};
use uuid::Uuid;
use chrono::{Utc, DateTime};
use lazy_static::lazy_static;
use tokio::time::Instant;
use serde::{Deserialize, Serialize};
mod excel;

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Thing {
    pub title: String,
    pub number1: f64,
    pub number2: f64,
}

fn create_things() -> Vec<Thing> {
    let mut result: Vec<Thing> = vec![];
    for _ in 0..10 {
        result.push(Thing {
            title: random_string(),
            number1: 43.1_f64,
            number2: 43.1_f64,
        });
    }
    result
}

fn random_string() -> String {
    Uuid::new_v4().to_string()
}

#[tokio::main]
async fn main() {
    let static_route = warp::fs::dir("static");

    let report_route = warp::get()
        .and(warp::path("report.xlsx"))
        .and(warp::query::<Thing>())
        .and_then(report_handler);

    let routes = warp::get().and(
        static_route
        .or(report_route)
    );

    println!("Server started at localhost:8080");
    warp::serve(routes).run(([0, 0, 0, 0], 8080)).await;
}

async fn report_handler(thing: Thing) -> Result<impl Reply, Infallible> {
    let now = Instant::now();
    let result = tokio::task::spawn_blocking(move || excel::create_xlsx(vec![thing]))
        .await
        .expect("can create result");
    println!("report took: {:?}", now.elapsed());
    Ok(result)
}
