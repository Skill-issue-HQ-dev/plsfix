//! 09.10.2026: bank PCs never deliver an HTTP PUT, so the add-in saves a link with POST to the
//! same address. POST is the same write as PUT (revisions, owner check, write limits); PUT stays
//! for add-ins loaded before the change.

use plsfix_server::{relay::*, store::Store};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
    response::Response,
    Router,
};
use http_body_util::BodyExt;
use tower::ServiceExt;

const ID: &str = "0123456789abcdef0123456789abcdef";
const MINE: &str = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"; // 43 chars
const THEIRS: &str = "BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

fn app() -> Router {
    routes(std::sync::Arc::new(AppState::new(
        Store::in_memory().unwrap(),
    )))
}

fn call(method: &str, auth: &str, body: &[u8]) -> Request<Body> {
    Request::builder()
        .method(method)
        .uri(format!("/api/links/{ID}"))
        .header(header::AUTHORIZATION, format!("Bearer {auth}"))
        .header(header::CONTENT_TYPE, "application/octet-stream")
        .body(Body::from(body.to_vec()))
        .unwrap()
}

async fn send(app: &Router, request: Request<Body>) -> Response {
    app.clone().oneshot(request).await.unwrap()
}

#[tokio::test]
async fn a_link_saved_by_post_is_the_same_write_as_put() {
    let app = app();
    let first = send(&app, call("POST", MINE, b"one")).await;
    assert_eq!(first.status(), StatusCode::OK);
    let body = first.into_body().collect().await.unwrap().to_bytes();
    let answer: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(answer["rev"], 1, "answer: {answer}");

    assert_eq!(send(&app, call("PUT", MINE, b"two")).await.status(), StatusCode::OK);
    let head = send(&app, call("GET", MINE, b"")).await;
    assert_eq!(head.headers().get(header::ETAG).unwrap(), "\"2\"");

    let foreign = send(&app, call("POST", THEIRS, b"x")).await;
    assert_eq!(foreign.status(), StatusCode::FORBIDDEN);
}
