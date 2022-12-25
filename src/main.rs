use warp::{Filter, Reply};
use std::convert::Infallible;
use tokio::time::Instant;
use serde::{Deserialize, Serialize};
mod excel;

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Thing {
    pub title: String,
    pub number1: f64,
    pub number2: f64,
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
